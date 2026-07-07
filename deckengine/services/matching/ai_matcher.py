# -*- coding: utf-8 -*-
"""
ai_matcher.py  --  Step 02 AI layer: refine transcript matching with judgment.

Keyword matching (in matcher.py) gives candidates; this asks the LLM to:
  - pick the case studies that GENUINELY fit the meeting (drop false positives
    that only share a generic keyword but are about a different topic), and
  - decide which OPTIONAL slides the transcript actually calls for.

Leader / people slides are handled by the caller, never here — they are never
auto-picked.

Provider: OpenAI (per project owner's choice). Key read from .env.
"""

import json

from deckengine.services.infra import load_env

MODEL = "gpt-4o-mini"


def _client():
    load_env()
    from openai import OpenAI
    return OpenAI()                      # reads OPENAI_API_KEY from the environment


def refine(transcript, candidates_by_wt, optional_slides, top_n=3):
    """
    transcript        : the pasted meeting text
    candidates_by_wt  : {work_type: [{slide_id, title, keywords}]}
    optional_slides   : [{slide_id, title}]
    Returns {"cases": {work_type: [slide_id]}, "optional": [slide_id]}.
    """
    lines = [
        "A salesperson pasted this meeting transcript:",
        '"""', transcript[:6000], '"""', "",
        "Pick the case studies that GENUINELY match what this meeting is about.",
        "Do NOT pick a slide that only shares a generic keyword but is about a "
        "different industry or topic than the meeting.",
        "If a listed case study clearly relates to the meeting, you MUST include "
        "it — only return an empty list for a work type when none are relevant.",
        f"Pick at most {top_n} per work type. Use only the slide IDs shown.",
    ]
    for wt, rows in candidates_by_wt.items():
        lines.append(f"\nCASE STUDIES for {wt}:")
        for r in rows:
            lines.append(f"  {r['slide_id']}: {r['title']} — keywords: {r['keywords']}")
    if optional_slides:
        lines.append("\nOPTIONAL slides (include only if the transcript clearly calls for them):")
        for r in optional_slides:
            lines.append(f"  {r['slide_id']}: {r['title']}")
    lines.append(
        '\nReturn ONLY this JSON shape: '
        '{"cases": {"WORKFORCE": [ids], "AI_POD": [ids], "MS": [ids]}, "optional": [ids]}'
    )

    resp = _client().chat.completions.create(
        model=MODEL,
        temperature=0,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": "You select the most relevant sales slides "
                                          "for a meeting. Reply with one JSON object only."},
            {"role": "user", "content": "\n".join(lines)},
        ],
    )
    try:
        data = json.loads(resp.choices[0].message.content)
    except (json.JSONDecodeError, AttributeError, IndexError):
        return {"cases": {}, "optional": []}

    # The model is asked for plain slide IDs, but occasionally returns objects
    # (e.g. {"slide_id": "CS70"}). Normalise everything to a list of ID STRINGS
    # so the caller never has to guess (and never hits 'unhashable type: dict').
    def _ids(v):
        out = []
        for x in (v or []):
            if isinstance(x, str):
                out.append(x)
            elif isinstance(x, dict):
                sid = x.get("slide_id") or x.get("id")
                if sid:
                    out.append(str(sid))
        return out

    raw_cases = data.get("cases") if isinstance(data.get("cases"), dict) else {}
    cases = {wt: _ids(v) for wt, v in raw_cases.items()}
    return {"cases": cases, "optional": _ids(data.get("optional"))}


def extract_profile(profile_text, max_items=8):
    """From a stakeholder's profile (LinkedIn/bio), pull WHAT THIS PERSON DOES:
    their function, key skills, and above all their CURRENT-ROLE mandate — so we
    pitch things relevant to their day-to-day. Returns [{"name","description"}]
    focus areas (e.g. 'Procurement', 'GCC / capability-center setup'). Fails []."""
    if not (profile_text or "").strip():
        return []
    prompt = (
        "Below is the professional profile of the person we are meeting. First read "
        "WHO they are, and especially their CURRENT role at their current company and "
        "what they are doing in it right now.\n"
        "Then list their SPECIFIC FUNCTIONAL DOMAINS — the concrete areas they work "
        "in that a case study could prove we understand. Name each by its DOMAIN, e.g. "
        "'Procurement', 'Contract management', 'Corporate real estate & facilities', "
        "'GCC / capability-center setup', 'Vendor management', 'Supply chain'.\n"
        "STRICT RULES:\n"
        "- Put the CURRENT-role domains FIRST.\n"
        "- Do NOT return generic management/soft labels (project management, program "
        "management, budget management, risk management, stakeholder management, "
        "leadership, communication) — return the DOMAIN they manage instead.\n"
        "- Each name = a specific function (1-3 words). description = 1-2 lines on what "
        "they do in it in their CURRENT role, grounded in the profile.\n"
        f"- At most {max_items}.\n"
        "PROFILE:\n\"\"\"\n" + profile_text[:9000] + "\n\"\"\"\n"
        'Return ONLY this JSON: {"items": [{"name": "...", "description": "..."}]}'
    )
    try:
        resp = _client().chat.completions.create(
            model=MODEL, temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You profile a buyer's real function "
                 "and current mandate from their bio. Reply with one JSON object only."},
                {"role": "user", "content": prompt},
            ],
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        return []
    out = []
    for m in (data.get("items") if isinstance(data, dict) else []) or []:
        if isinstance(m, dict) and (m.get("name") or "").strip():
            out.append({"name": m["name"].strip(),
                        "description": (m.get("description") or "").strip()})
    return out[:max_items]


def extract_accelerators(notes, max_items=8):
    """From the meeting notes + deep-research brief, list the named accelerators /
    capabilities / solution areas the ACCOUNT needs — EXTRACTION ONLY.

    Whether we already have a case for each is decided separately by SEMANTIC
    match against the store (reliable), NOT by the model eyeballing the library
    (which mis-judges coverage). Returns [{"name","description"}]. Fails safe []."""
    if not (notes or "").strip():
        return []
    prompt = (
        "From these client meeting notes + research brief, list the named "
        "accelerators, capabilities, or solution areas the ACCOUNT needs (named "
        "explicitly, or a clearly implied need). For each, give a 1-2 line "
        "description of what it is and the problem it solves, grounded ONLY in the "
        "text (no invented metrics or facts).\n"
        f"- At most {max_items}. Prefer specific, named accelerators over generic themes.\n"
        "- Skip individual tools/languages (Selenium, Python, Docker); name the capability.\n"
        "NOTES:\n\"\"\"\n" + notes[:9000] + "\n\"\"\"\n"
        'Return ONLY this JSON: {"items": [{"name": "...", "description": "..."}]}'
    )
    try:
        resp = _client().chat.completions.create(
            model=MODEL,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You extract a client's needed "
                 "capabilities from meeting notes. Reply with one JSON object only."},
                {"role": "user", "content": prompt},
            ],
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        return []
    out = []
    for m in (data.get("items") if isinstance(data, dict) else []) or []:
        if isinstance(m, dict) and (m.get("name") or "").strip():
            out.append({"name": m["name"].strip(),
                        "description": (m.get("description") or "").strip()})
    return out[:max_items]


def extract_client_context(notes):
    """For the "Client Context" / "Tailored Approach" slide pair (Workforce-only,
    First/Second stage) -- pull the account's REAL talent-challenge facts from
    the notes, grounded ONLY in what's actually there. Never invents a number or
    a challenge/solution that isn't named or clearly implied.

    Returns a dict (see schema below) or None if the notes don't give enough to
    fill the slides honestly -- FAILS CLOSED: the caller must not build either
    slide with fewer than 2 real challenges/solutions or no client name (a half-
    filled template with leftover brackets is worse than no slide at all)."""
    if not (notes or "").strip():
        return None
    prompt = (
        "From these client meeting notes, extract facts for a 'Client Context' + "
        "'Tailored Approach' slide pair about their TALENT/HIRING challenges. Ground "
        "every field ONLY in the notes -- if a fact isn't stated or clearly implied, "
        "leave it as an empty string. Never invent a number, a client name, or a "
        "challenge/solution that isn't real.\n"
        "Return:\n"
        "- client_name: the client's name.\n"
        "- date: the discovery session/meeting date, if mentioned (empty if not).\n"
        "- offer_drop_pct: their offer-drop-rate number as a bare number+unit (e.g. "
        "'35%'), if mentioned (empty if not).\n"
        "- org_size, hiring_range, hiring_year, junior_pct, city, hq: short bare "
        "values if mentioned (e.g. org_size='450', hiring_range='40-60', "
        "hiring_year='2026', junior_pct='60%', city='Bengaluru', hq='London') -- "
        "empty string for any not mentioned.\n"
        "- challenges: up to 4 REAL talent/hiring challenges this account actually "
        "has, each {\"title\": short heading (3-6 words), \"body\": one factual "
        "sentence (max ~30 words) explaining it, grounded in the notes}.\n"
        "- solutions: up to 4 REAL approaches/solutions discussed for those "
        "challenges, each {\"title\": short heading, \"body\": one factual sentence, "
        "\"solves\": which challenge(s) it addresses, in a few words}. Order solutions "
        "to correspond to the challenges list where possible.\n"
        "If the notes don't give at least 2 real challenges or don't name the "
        "client, return empty lists / empty client_name -- do NOT pad with "
        "generic or invented content.\n"
        "NOTES:\n\"\"\"\n" + notes[:9000] + "\n\"\"\"\n"
        'Return ONLY this JSON: {"client_name":"...","date":"...","offer_drop_pct":"...",'
        '"org_size":"...","hiring_range":"...","hiring_year":"...","junior_pct":"...",'
        '"city":"...","hq":"...","challenges":[{"title":"...","body":"..."}],'
        '"solutions":[{"title":"...","body":"...","solves":"..."}]}'
    )
    try:
        resp = _client().chat.completions.create(
            model=MODEL, temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You extract real, grounded talent-"
                 "challenge facts for a client-context slide. Reply with one JSON "
                 "object only. Never invent facts not in the notes."},
                {"role": "user", "content": prompt},
            ],
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        return None
    if not isinstance(data, dict):
        return None

    def _pairs(key, extra=()):
        out = []
        for m in (data.get(key) or []):
            if isinstance(m, dict) and (m.get("title") or "").strip():
                rec = {"title": m["title"].strip(), "body": (m.get("body") or "").strip()}
                for k in extra:
                    rec[k] = (m.get(k) or "").strip()
                out.append(rec)
        return out[:4]

    result = {
        "client_name": (data.get("client_name") or "").strip(),
        "date": (data.get("date") or "").strip(),
        "offer_drop_pct": (data.get("offer_drop_pct") or "").strip(),
        "org_size": (data.get("org_size") or "").strip(),
        "hiring_range": (data.get("hiring_range") or "").strip(),
        "hiring_year": (data.get("hiring_year") or "").strip(),
        "junior_pct": (data.get("junior_pct") or "").strip(),
        "city": (data.get("city") or "").strip(),
        "hq": (data.get("hq") or "").strip(),
        "challenges": _pairs("challenges"),
        "solutions": _pairs("solutions", extra=("solves",)),
    }
    # fail closed: need a client name + at least 2 real challenges AND solutions
    if not result["client_name"] or len(result["challenges"]) < 2 or len(result["solutions"]) < 2:
        return None
    return result


def extract_skill_profile(notes, max_per_category=6):
    """From the meeting notes ('more information'), pull the SKILLS/requirements
    the client is hiring/staffing against and categorise each into exactly one of
    three buckets for the "Target Skill Profile" slide (Workforce-only):
      - domain_expertise: business/domain knowledge areas (e.g. "Banking Regulation
        & Compliance", "Financial Services & Lending")
      - technical_stack: tools/platforms/technologies (e.g. "Cloud & Data Engineering",
        "Test Automation")
      - academic_professional: qualifications/experience/soft requirements (e.g.
        "Engineering & CS Graduates", "0-12 Years Experience")

    Each item is {"name": short heading (2-5 words), "description": one short line}.
    Returns {} if the notes name no skills/requirements (nothing to show) or on any
    API error — the caller treats an empty result as "no slide", never a placeholder."""
    if not (notes or "").strip():
        return {}
    prompt = (
        "From these client meeting notes, identify the SKILLS, technologies, domain "
        "knowledge, and background the client is hiring / staffing / building a team "
        "against. This powers a 'Target Skill Profile' slide with three columns.\n"
        "Categorise EVERY item into exactly one bucket:\n"
        "- domain_expertise: business/industry/domain knowledge areas (e.g. "
        "'Banking Regulation & Compliance', 'Financial Services & Lending', "
        "'Business Process Domain').\n"
        "- technical_stack: concrete tools, platforms, languages, technologies (e.g. "
        "'Cloud & Data Engineering', 'Test Automation', 'Security & IAM Tooling').\n"
        "- academic_professional: qualifications, experience level, certifications, "
        "soft/behavioural requirements (e.g. 'Engineering & CS Graduates', "
        "'0-12 Years Experience', 'High Learning Agility').\n"
        "For each item: name = a short heading (2-5 words), description = one short "
        "line (max ~8 words) naming the specifics (e.g. tools, standards, range).\n"
        f"At most {max_per_category} items per bucket. Ground every item in the notes "
        "— do not invent skills that aren't named or clearly implied. If the notes "
        "name no skills/requirements at all, return every bucket empty.\n"
        "NOTES:\n\"\"\"\n" + notes[:9000] + "\n\"\"\"\n"
        'Return ONLY this JSON: {"domain_expertise": [{"name":"...","description":"..."}], '
        '"technical_stack": [...], "academic_professional": [...]}'
    )
    try:
        resp = _client().chat.completions.create(
            model=MODEL, temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You categorise a client's hiring/staffing "
                 "requirements into a skill profile. Reply with one JSON object only."},
                {"role": "user", "content": prompt},
            ],
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        return {}
    if not isinstance(data, dict):
        return {}

    def _items(key):
        out = []
        for m in (data.get(key) or []):
            if isinstance(m, dict) and (m.get("name") or "").strip():
                out.append({"name": m["name"].strip(),
                            "description": (m.get("description") or "").strip()})
        return out[:max_per_category]

    result = {
        "domain_expertise": _items("domain_expertise"),
        "technical_stack": _items("technical_stack"),
        "academic_professional": _items("academic_professional"),
    }
    if not any(result.values()):
        return {}
    return result


def extract_asks(transcript):
    """Pull the SPECIFIC capability / skill / technology asks the CLIENT made.

    Returns a short list of concise phrases, e.g. ['ADAS', 'fraud detection'].
    This is what lets the engine flag "X was asked but isn't in the deck" even
    for a topic that exists in NO slide yet (the whole point of the gap fix).
    Conservative by design — concrete asks only, capped, never generic words.
    Fails safe to [] on any error (the caller still has keyword detection)."""
    if not (transcript or "").strip():
        return []
    prompt = (
        "A salesperson pasted this client meeting transcript:\n"
        '"""\n' + transcript[:6000] + '\n"""\n\n'
        "List the SUBSTANTIAL capabilities, solutions, or domains the CLIENT asked "
        "for that would each justify a DEDICATED sales slide. Rules:\n"
        "- Capability/solution THEMES only (e.g. 'ADAS', 'fraud detection', "
        "'predictive maintenance', 'blockchain traceability', 'test automation').\n"
        "- Do NOT list individual tools, libraries, frameworks, or programming "
        "languages (e.g. Cucumber, Selenium, JUnit, React, Python, Docker, Jenkins). "
        "These are details, not deck themes. If the client only named tools, infer "
        "the capability they belong to (Selenium/Cucumber -> test automation) or skip.\n"
        "- Do NOT include generic words (software, team, quality, support, project, "
        "solution, technology, help, service).\n"
        "- Only things the client wants delivered — not background chit-chat.\n"
        "- At most 6 items. If none clearly merit their own slide, return an empty list.\n"
        "- Keep acronyms/proper nouns as written (ADAS, SAP); otherwise lowercase.\n"
        'Return ONLY this JSON: {"asks": ["...", "..."]}'
    )
    try:
        resp = _client().chat.completions.create(
            model=MODEL,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You extract concrete client asks from "
                                              "a sales meeting. Reply with one JSON object only."},
                {"role": "user", "content": prompt},
            ],
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        return []
    asks = data.get("asks") if isinstance(data, dict) else None
    out = []
    for a in (asks or []):
        if isinstance(a, str) and a.strip():
            out.append(a.strip())
        elif isinstance(a, dict):
            v = a.get("ask") or a.get("topic") or a.get("name")
            if v:
                out.append(str(v).strip())
    return out[:6]


def extract_brief(research="", profile="", transcript=""):
    """Structured matching brief from the deep-research brief + stakeholder profile +
    transcript — the signals the ranker uses DIRECTLY (not just descriptive text):

        {"needs": [{"name","description","domain","use_case"}],
         "avoid": [{"capability","reason"}],       # the brief's mismatch flags
         "expressed_interest": ["..."],            # accelerators discussed in the transcript
         "account": {"industry","role","company_context"}}

    Reference priority: a deep-research brief is PRIMARY; the transcript adds prior
    interest; profile-only derives needs from the role/company. Fails safe to {} so
    the caller can fall back to the name-only extractors."""
    research = (research or "").strip()
    profile = (profile or "").strip()
    transcript = (transcript or "").strip()
    if not (research or profile or transcript):
        return {}
    parts = []
    if research:
        parts.append("DEEP RESEARCH BRIEF (PRIMARY — includes synergy mapping and "
                     "mismatch flags):\n\"\"\"\n" + research[:9000] + "\n\"\"\"")
    if profile:
        parts.append("STAKEHOLDER PROFILE:\n\"\"\"\n" + profile[:6000] + "\n\"\"\"")
    if transcript:
        parts.append("MEETING TRANSCRIPT (prior sales context — expressed interest):"
                     "\n\"\"\"\n" + transcript[:6000] + "\n\"\"\"")
    prompt = (
        "From the sources below, extract a STRUCTURED brief for matching this client to "
        "a case-study library. Rules:\n"
        "- needs: capabilities/solution areas this account GENUINELY needs. For each: "
        "name (1-3 words, a business capability), a 1-2 line description, the DOMAIN / "
        "industry it applies to, and the specific USE_CASE. Ground every field in the "
        "text; never invent.\n"
        "- avoid: MISMATCH FLAGS — capabilities that look related by keyword but are the "
        "WRONG fit for this account (wrong domain / use-case). Give the capability and a "
        "one-line reason. Only clear or explicitly-flagged misfits.\n"
        "- expressed_interest: specific accelerators/topics the CLIENT actually discussed "
        "(from the transcript). Empty if none.\n"
        "- account: {industry, role, company_context} from the profile/brief.\n"
        "Reference priority: if a deep-research brief is present treat it as PRIMARY and "
        "the transcript as prior-interest only; if only a profile is present, derive needs "
        "from the role and company context.\n"
        "- Skip individual tools/languages (Selenium, Python, Docker); name the capability.\n\n"
        + "\n\n".join(parts) + "\n\n"
        'Return ONLY this JSON: {"needs":[{"name":"...","description":"...","domain":"...",'
        '"use_case":"..."}],"avoid":[{"capability":"...","reason":"..."}],'
        '"expressed_interest":["..."],"account":{"industry":"...","role":"...",'
        '"company_context":"..."}}'
    )
    try:
        resp = _client().chat.completions.create(
            model=MODEL, temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You extract a structured client-matching "
                 "brief from sales research. Reply with one JSON object only."},
                {"role": "user", "content": prompt},
            ],
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        return {}
    if not isinstance(data, dict):
        return {}

    def _needs(v):
        out = []
        for m in (v or []):
            if isinstance(m, dict) and (m.get("name") or "").strip():
                out.append({"name": m["name"].strip(),
                            "description": (m.get("description") or "").strip(),
                            "domain": (m.get("domain") or "").strip(),
                            "use_case": (m.get("use_case") or "").strip()})
        return out[:8]

    def _avoid(v):
        out = []
        for m in (v or []):
            if isinstance(m, dict) and (m.get("capability") or "").strip():
                out.append({"capability": m["capability"].strip(),
                            "reason": (m.get("reason") or "").strip()})
        return out[:8]

    ei = [str(x).strip() for x in (data.get("expressed_interest") or []) if str(x).strip()][:8]
    acct = data.get("account") if isinstance(data.get("account"), dict) else {}
    return {"needs": _needs(data.get("needs")), "avoid": _avoid(data.get("avoid")),
            "expressed_interest": ei,
            "account": {"industry": (acct.get("industry") or "").strip(),
                        "role": (acct.get("role") or "").strip(),
                        "company_context": (acct.get("company_context") or "").strip()}}


def explain_picks(brief, items):
    """Write ONE short line + a signal explaining why each ALREADY-PICKED case fits its
    client need. The SELECTION is done algorithmically (semantic capability match) — this
    only EXPLAINS, so the model cannot mis-pick or hallucinate a different case.

    items = [{"need","id","title","blurb"}]. brief supplies account + expressed interest.
    Returns {case_id: {"reason","signal"}}. Fails safe to {} (caller uses a template)."""
    items = list(items or [])
    if not items:
        return {}
    brief = brief or {}
    acct = brief.get("account") or {}
    ei = brief.get("expressed_interest") or []
    lines = [
        "For each PICKED case study below, write ONE short line (max ~18 words) on why it "
        "fits the client need, and a signal tag. Be specific; do not invent case facts.",
        "signal: 'capability' (proves the needed capability), 'industry' (also same "
        "industry), 'interest' (a topic the client discussed), or 'role' (fits the "
        "stakeholder's function).",
        "",
    ]
    if acct.get("industry") or acct.get("role"):
        lines.append(f"ACCOUNT: industry={acct.get('industry','')} role={acct.get('role','')}")
    if ei:
        lines.append("EXPRESSED INTEREST: " + ", ".join(ei))
    lines.append("")
    for it in items:
        lines.append(f'NEED "{it.get("need","")}" -> {it["id"]}: {it.get("title","")}. '
                     f'{(it.get("blurb","") or "")[:160]}')
    lines.append('\nReturn ONLY JSON mapping each case id to '
                 '{"reason":"one line","signal":"capability|industry|interest|role"}')
    try:
        resp = _client().chat.completions.create(
            model=MODEL, temperature=0.2,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You explain why a proof point fits a client "
                 "need. Reply with one JSON object only."},
                {"role": "user", "content": "\n".join(lines)},
            ],
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        return {}
    if not isinstance(data, dict):
        return {}
    valid = {it["id"] for it in items}
    out = {}
    for cid, v in data.items():
        if cid in valid and isinstance(v, dict):
            out[cid] = {"reason": (v.get("reason") or "").strip(),
                        "signal": (v.get("signal") or "").strip().lower()}
    return out
