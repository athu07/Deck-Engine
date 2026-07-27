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
import logging

from deckengine.services.infra import load_env

MODEL = "gpt-4o-mini"
# explain_picks()/validate_bet_fit() do real JUDGMENT work now (the fit check: is this
# case a genuine mechanism+subject-matter match, or just a capability-label coincidence),
# not just formatting/summarizing -- confirmed 2026-07-23 (Vanderlande, MSS117): given
# fully correct information (an accurate account description AND the case's real content
# spelling out "vehicle documents"), gpt-4o-mini's own generated reasoning named the exact
# mismatch in its own words and still returned fit=true. The prompt/data were right; the
# model's judgment reliability on this specific class of call wasn't. Everything else in
# this file (extraction, refine, accelerators) stays on MODEL -- those are simpler, more
# mechanical tasks where mini has tested reliably; only the two fact-checking calls below
# get the stronger model, to bound the added cost/latency to where it's actually needed.
EXPLAIN_MODEL = "gpt-4o"
logger = logging.getLogger(__name__)

# gpt-4o-mini's context window is 128K tokens (~500K chars); a deep-research brief
# truncated at a few thousand chars was silently dropping over half of real, dense
# briefs (confirmed 2026-07-23: a 19,241-char brief was cut mid-section-7, discarding
# the entire stakeholder-background section). This cap is sized generously above any
# brief seen in practice while keeping prompt + profile + transcript comfortably
# inside the model's window.
RESEARCH_CHAR_CAP = 40000


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
        "— do not invent skills that aren't named or clearly implied. Only extract "
        "items that are genuine hiring/staffing/team-build requirements (roles or "
        "skills the client wants to hire or deploy against). General company or "
        "industry background (e.g. 'they are an FMCG company', 'they operate in "
        "food & agriculture') is NOT a skill requirement — do not turn it into one. "
        "If the notes name no real hiring/staffing skills or requirements, return "
        "every bucket empty.\n"
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
                     "mismatch flags):\n\"\"\"\n" + research[:RESEARCH_CHAR_CAP] + "\n\"\"\"")
    if profile:
        parts.append("STAKEHOLDER PROFILE:\n\"\"\"\n" + profile[:6000] + "\n\"\"\"")
    if transcript:
        parts.append("MEETING TRANSCRIPT (prior sales context — expressed interest):"
                     "\n\"\"\"\n" + transcript[:6000] + "\n\"\"\"")
    prompt = (
        "From the sources below, extract a STRUCTURED brief for matching this client to "
        "a case-study library. Rules:\n"
        "- needs: capabilities/solution areas this account GENUINELY needs. Ground these "
        "in the person's SUBSTANTIVE experience -- certifications, hands-on technical "
        "work, and anything emphasised or repeated across their career -- not just a "
        "shallow 'skills' tag list at the top of a profile (LinkedIn's own 'Top Skills' "
        "section is often generic/self-selected and far less reliable than the actual "
        "work-history descriptions and certifications below it). A stakeholder's own "
        "deep, certified, career-defining expertise is a STRONG signal of what they'd "
        "care about, even when their current title is a leadership role they've since "
        "been promoted into -- do NOT dismiss it as \"past experience that doesn't apply "
        "now\"; someone who spent a decade doing X and still holds a current "
        "certification in X is exactly the person who'd want to see X modernised, "
        "augmented, or scaled. For each need: name (1-3 words, a business capability), a "
        "1-2 line description, the DOMAIN / industry it applies to, and the specific "
        "USE_CASE. Ground every field in the text; never invent.\n"
        "- GRANULARITY: do not collapse someone's expertise into one umbrella need. If a "
        "person's background spans several genuinely distinct capabilities (e.g. a "
        "platform-migration specialism, a specific in-memory/performance specialism, a "
        "low-code/extension-development specialism, an AI-assisted-tooling angle, a "
        "cross-functional/governance angle), list EACH as its own separate need with its "
        "own name and use_case -- even if they all relate to the same underlying platform "
        "or product family. A broad bucket like \"SAP Development\" hides everything "
        "underneath it once ONE matching case study is found for it, so the more specific "
        "capabilities inside it never get checked against the library on their own -- "
        "always prefer the more specific, separately-checkable need over a broad one that "
        "would swallow it. Extract up to 12 needs when the sources support that many "
        "distinct, real capabilities; do not pad with near-duplicates.\n"
        "- NAMED ACCELERATORS: if the deep-research brief itself proposes specific, named, "
        "buildable accelerators (e.g. \"Accelerator 1: SAP BTP Extension for Real-time "
        "Analytics\") -- these are already the single most specific, vetted, actionable "
        "line items in the whole document, written for exactly this purpose. Extract EACH "
        "one as its own distinct need, using its own given name and description near-"
        "verbatim -- do not paraphrase it into a broader theme, rename it into a generic "
        "category, or let a different, vaguer mention of the same technology elsewhere in "
        "the brief (e.g. a passing reference under a general 'synergies' section) take its "
        "place instead.\n"
        "- EXPLICIT FRAMING INSTRUCTIONS: a research brief often tells the salesperson "
        "directly how to pitch, not just what the account cares about -- \"any pitch here "
        "should be framed around X, Y and Z\", \"position around A, B, C\", \"this creates "
        "demand for P, Q and R\". Each named item in a sentence like this is a real, "
        "specific, separately-checkable need in its own right -- extract EACH ONE (P, Q, "
        "AND R, not just the sentence's general topic) exactly as named, even if a "
        "broader-sounding need elsewhere in the brief would seem to already cover it. "
        "Confirmed miss (owner-reported, 2026-07-23, Vanderlande): a brief said pitches "
        "\"should be framed around the 3DEXPERIENCE ecosystem, EBOM/MBOM continuity, and "
        "the split Veghel-Pune PLM team\" and separately \"creates demand for CAD/design "
        "automation, BOM harmonization, and engineering-process automation\" -- extraction "
        "produced a \"PLM Automation\" need but never \"BOM Harmonization\" on its own, so "
        "the one case in the library actually built around BOM/EBOM/MBOM never got checked "
        "on its own terms and lost the generic \"PLM\" need to an unrelated case whose "
        "TITLE merely happened to contain the word \"PLM\". Named items in these "
        "instruction sentences are exactly the failure mode GRANULARITY above warns "
        "about -- treat them with the same discipline as a NAMED ACCELERATOR.\n"
        "- avoid: MISMATCH FLAGS — capabilities that look related by keyword but are "
        "genuinely the WRONG fit for this ACCOUNT (wrong domain / wrong use-case for the "
        "COMPANY's business). Never use this to suppress a stakeholder's own real, "
        "substantive background just because they now hold a senior/leadership title -- "
        "that is normally a MATCH, not a mismatch (see the needs rule above). Give the "
        "capability and a one-line reason. Only clear or explicitly-flagged misfits.\n"
        "- expressed_interest: specific accelerators/topics the CLIENT actually discussed "
        "(from the transcript). Empty if none.\n"
        "- account: {industry, role, company_context} from the profile/brief. "
        "company_context is NOT the company's name -- it's a 1-2 sentence description of "
        "what the account ACTUALLY DOES/SELLS/OPERATES (e.g. \"baggage-handling and parcel-"
        "sortation systems for airports and logistics operators\", not just \"Vanderlande\"). "
        "This is the yardstick a later step uses to judge whether a proof point genuinely "
        "fits or is a superficial capability-label match to an unrelated business -- a bare "
        "company name gives that step nothing to judge against (owner-reported, 2026-07-23: "
        "a case about validating VEHICLE REGISTRATION documents passed as \"compliance "
        "automation\" for a warehouse-automation company because company_context was just "
        "the company's name, with no description of what the account does to compare "
        "against). Ground it in the sources; never invent details not present.\n"
        "- prefer_high_impact: true ONLY if the transcript/notes EXPLICITLY ask for case "
        "studies with strong/proven/high-impact numbers (e.g. \"show cases with the best "
        "ROI\", \"need proof points with real margin improvement\", \"cases where we've "
        "moved the needle\") -- a request about the STRENGTH of the evidence, not just "
        "that a capability is needed. False by default; do not infer this from the mere "
        "presence of numbers or business language elsewhere in the sources.\n"
        "- asks_differentiation: true if the client's own conversation shows they want to "
        "understand HOW our engagement/delivery MODEL is structurally different from a "
        "typical vendor (not just what we can technically do) -- e.g. asking how "
        "engagement, risk, or accountability actually works differently, not just a "
        "capability list.\n"
        "- asks_why_not_big_si: true if the client's own conversation raises, even "
        "implicitly, why they'd choose us over a Big 4 firm, a large systems integrator, "
        "or an established incumbent vendor -- a genuine competitive-positioning question, "
        "not just the account happening to already use a large vendor.\n"
        "Both default false; only set true from a genuine signal in the sources, never "
        "guessed from the account being large or enterprise-sized alone.\n"
        "Reference priority: if a deep-research brief is present treat it as PRIMARY and "
        "the transcript as prior-interest only; if only a profile is present, derive needs "
        "from the role and company context.\n"
        "- Skip individual tools/languages (Selenium, Python, Docker); name the capability.\n\n"
        + "\n\n".join(parts) + "\n\n"
        'Return ONLY this JSON: {"needs":[{"name":"...","description":"...","domain":"...",'
        '"use_case":"..."}],"avoid":[{"capability":"...","reason":"..."}],'
        '"expressed_interest":["..."],"account":{"industry":"...","role":"...",'
        '"company_context":"..."},"prefer_high_impact":false,"asks_differentiation":false,'
        '"asks_why_not_big_si":false}'
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
        logger.warning("extract_brief: OpenAI call/parse failed — matching falls back "
                       "to generic ranking with no research-driven picks", exc_info=True)
        return {}
    if not isinstance(data, dict):
        logger.warning("extract_brief: model returned non-JSON-object content: %r", data)
        return {}

    def _needs(v):
        out = []
        for m in (v or []):
            if isinstance(m, dict) and (m.get("name") or "").strip():
                out.append({"name": m["name"].strip(),
                            "description": (m.get("description") or "").strip(),
                            "domain": (m.get("domain") or "").strip(),
                            "use_case": (m.get("use_case") or "").strip()})
        return out[:12]

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
                        "company_context": (acct.get("company_context") or "").strip()},
            "prefer_high_impact": bool(data.get("prefer_high_impact")),
            "asks_differentiation": bool(data.get("asks_differentiation")),
            "asks_why_not_big_si": bool(data.get("asks_why_not_big_si"))}


def infer_strategic_fit(research="", profile="", transcript="", brief=None):
    """Deep-research-style INFERENCE, deliberately separate from extract_brief()'s
    literal grounding: proposes up to 4 capability bets that are NOT stated anywhere
    in the sources, but that a sharp account researcher would infer by connecting the
    STAKEHOLDER's own career/role pattern to the COMPANY's business.

    Owner's spec (2026-07-08, built from a real example: Pankaj Kumar Pant / Waaree
    Group): nothing in his profile says "we want predictive maintenance for solar
    assets," but 2 years running a solar ingot/wafer plant plus a career built on
    defect-catching (FMEA, paint QC) makes that a reasoned bet -- the kind of
    connection a deep-research pass makes but literal keyword/need extraction can't.
    A bet only earns a spot if BOTH sides hold up: it has to be something THIS
    PERSON would personally champion (their own role/career, not a generic pitch)
    AND something THIS COMPANY would strategically fund (its actual business, not
    "any manufacturer could use this"). Either side unconvincing and it's dropped --
    caller-side too, via the stakeholder_why/company_why presence check below, not
    just prompt instruction.

    brief = the already-extracted literal brief (needs/avoid/account), so this never
    re-proposes what extract_brief already covered. Returns a list of
    {"name","description","stakeholder_why","company_why"}, capped at 4. Fails safe
    to [] -- an inferred pick is a bonus on top of the literal ones, never required."""
    research = (research or "").strip()
    profile = (profile or "").strip()
    transcript = (transcript or "").strip()
    if not (research or profile):
        return []      # nothing to extrapolate a CAREER PATTERN or COMPANY context from
    brief = brief or {}
    acct = brief.get("account") or {}
    covered = sorted({n["name"] for n in (brief.get("needs") or [])})
    avoid = brief.get("avoid") or []

    parts = []
    if profile:
        parts.append("STAKEHOLDER PROFILE (their career/role):\n\"\"\"\n" + profile[:6000] + "\n\"\"\"")
    if research:
        parts.append("COMPANY / DEEP RESEARCH:\n\"\"\"\n" + research[:RESEARCH_CHAR_CAP] + "\n\"\"\"")
    if transcript:
        parts.append("MEETING NOTES:\n\"\"\"\n" + transcript[:4000] + "\n\"\"\"")
    if acct.get("industry") or acct.get("company_context"):
        parts.append(f"ACCOUNT CONTEXT: industry={acct.get('industry','')} "
                     f"company_context={acct.get('company_context','')}")

    prompt = (
        "You are doing STRATEGIC ACCOUNT RESEARCH, not literal extraction. The literally "
        "stated needs have already been pulled out separately -- your job is different: "
        "propose up to 4 capability bets that are NOT stated anywhere in the sources, but "
        "that a sharp account researcher would infer by connecting the stakeholder's own "
        "career pattern to the company's business.\n\n"
        "For each bet you must argue BOTH sides, or don't propose it:\n"
        "- stakeholder_why: why THIS SPECIFIC PERSON, given their actual role/career "
        "history, would personally want it -- something that helps THEM do their job or "
        "look good to their own leadership. Cite the actual role/project detail you're "
        "extrapolating from.\n"
        "- company_why: why THIS COMPANY, given its actual business/industry, would "
        "strategically fund it -- not just 'it's generally useful,' but a reason tied to "
        "what this specific company does.\n"
        "A bet that only satisfies one side is not good enough -- skip it rather than force "
        "a weak connection. Do not propose a generic capability that could apply to any "
        "account; the whole point is that it's specific to THIS person at THIS company.\n"
        + (f"Already covered by literal needs, don't repeat: {', '.join(covered)}.\n" if covered else "")
        + (("Known mismatches, don't propose these: " +
           "; ".join(f"{a['capability']} ({a['reason']})" for a in avoid) + ".\n") if avoid else "")
        + "\n" + "\n\n".join(parts) + "\n\n"
        'Return ONLY this JSON: {"bets":[{"name":"...","description":"...",'
        '"stakeholder_why":"...","company_why":"..."}]} -- 0 to 4 items; fewer is '
        "fine if you can't honestly argue both sides for more."
    )
    try:
        resp = _client().chat.completions.create(
            model=MODEL, temperature=0.3,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You are a strategic account researcher who "
                 "makes reasoned, evidence-grounded inferences -- not wild guesses. Reply "
                 "with one JSON object only."},
                {"role": "user", "content": prompt},
            ],
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        logger.warning("infer_strategic_fit: OpenAI call/parse failed — no strategic "
                       "bets this build", exc_info=True)
        return []
    if not isinstance(data, dict):
        return []
    out = []
    for b in (data.get("bets") or [])[:4]:
        if not isinstance(b, dict):
            continue
        name = (b.get("name") or "").strip()
        sw = (b.get("stakeholder_why") or "").strip()
        cw = (b.get("company_why") or "").strip()
        if name and sw and cw:      # BOTH sides required -- enforced in code, not just prompt
            out.append({"name": name, "description": (b.get("description") or "").strip(),
                        "stakeholder_why": sw, "company_why": cw})
    return out


def validate_bet_fit(bet, case_title, case_challenge, case_solution):
    """A strategic bet's stakeholder_why/company_why (infer_strategic_fit, above) is
    written BEFORE any real case is matched to it -- the caller's shortlist step only
    picks the nearest case AFTER the narrative already exists, with nothing to stop a
    narrative written for one imagined capability ending up pinned to an unrelated real
    case. Confirmed 2026-07-23 (Kimberly-Clark audit): a bet's narrative about "supplier
    collaboration" survived unmodified onto MSS048, a banking finance-close/IT-ticket
    case with nothing to do with suppliers -- the bet was never checked against what the
    matched case actually contains.

    Re-grounds the two-angle narrative in the REAL matched case's own challenge/solution:
    tightens the wording to what the case actually does if it genuinely still supports
    the bet, or returns fit=false if the matched case doesn't really back this narrative
    (a keyword-level coincidence, not a real fit) so the caller can drop the bet rather
    than keep a mismatched story. Fails safe to fit=true, UNCHANGED narrative -- an AI
    outage here must never silently kill a bet that was already good."""
    case_title = (case_title or "").strip()
    case_challenge = (case_challenge or "").strip()
    case_solution = (case_solution or "").strip()
    if not (case_title or case_challenge or case_solution):
        return {"stakeholder_why": bet.get("stakeholder_why", ""),
                "company_why": bet.get("company_why", ""), "fit": True}
    prompt = (
        "A strategic-bet narrative was written proposing a capability BEFORE any real "
        "case study was matched to it. A real case has now been matched by keyword/semantic "
        "similarity -- your job is to check whether that REAL case's actual content still "
        "genuinely supports the narrative, and tighten the wording to the case's real facts.\n\n"
        f'PROPOSED CAPABILITY: "{bet.get("name","")}" -- {bet.get("description","")}\n'
        f'STAKEHOLDER ANGLE (as originally written): {bet.get("stakeholder_why","")}\n'
        f'COMPANY ANGLE (as originally written): {bet.get("company_why","")}\n\n'
        f"REAL CASE MATCHED TO THIS BET: {case_title}\n"
        f"CHALLENGE: {case_challenge[:300]}\n"
        f"SOLUTION: {case_solution[:300]}\n\n"
        "If the real case's actual mechanism genuinely supports the proposed capability, "
        "rewrite BOTH angles to reference what the case actually does (not the original "
        "guess) -- keep them tight, one line each. If the real case does NOT genuinely "
        "support this narrative (a different mechanism, a different business entirely, "
        "only a coincidental keyword overlap -- e.g. a banking finance-close case matched "
        "to a 'supplier collaboration' bet has nothing to do with suppliers), set "
        '"fit":false and leave the angle fields empty -- do not force a connection that '
        "isn't really there.\n\n"
        'Return ONLY JSON: {"stakeholder_why":"...","company_why":"...","fit":true}'
    )
    try:
        resp = _client().chat.completions.create(
            model=EXPLAIN_MODEL, temperature=0.2,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You fact-check a sales narrative against a "
                 "real case's actual content before it goes in front of a client. Reply "
                 "with one JSON object only."},
                {"role": "user", "content": prompt},
            ],
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        logger.warning("validate_bet_fit: OpenAI call/parse failed — keeping the bet's "
                       "original unvalidated narrative", exc_info=True)
        return {"stakeholder_why": bet.get("stakeholder_why", ""),
                "company_why": bet.get("company_why", ""), "fit": True}
    if not isinstance(data, dict):
        return {"stakeholder_why": bet.get("stakeholder_why", ""),
                "company_why": bet.get("company_why", ""), "fit": True}
    fit = data.get("fit") is not False
    sw = (data.get("stakeholder_why") or "").strip()
    cw = (data.get("company_why") or "").strip()
    return {"stakeholder_why": sw if (fit and sw) else bet.get("stakeholder_why", ""),
            "company_why": cw if (fit and cw) else bet.get("company_why", ""),
            "fit": fit}


def explain_picks(brief, items, profile="", research=""):
    """Write ONE short line + a signal explaining why each ALREADY-PICKED case fits its
    client need. The SELECTION is done algorithmically (semantic capability match) — this
    EXPLAINS, so the model cannot mis-pick or SUBSTITUTE a different case; it CAN say a
    pick doesn't actually hold up (fit=false, see below) so the caller can drop it, but it
    never gets to choose what replaces it.

    items = [{"need","id","title","blurb","solution"}]. brief supplies account + expressed
    interest. profile/research (owner-spec, 2026-07-23: "why is our tool not this powerful"
    — Kimberly-Clark test) let the model ground a reason in a SPECIFIC real fact about the
    stakeholder or account (a prior company, a named event, a role detail) when one
    genuinely exists, the same grounded-narrative move infer_strategic_fit() already makes
    for its bonus bets — now available to EVERY literal pick, not just those. Bounded
    slices only (this is explanation, not extraction); never invents beyond the text given.

    "solution" (owner-reported, 2026-07-23: the Kimberly-Clark audit) -- explain_picks used
    to see only a 160-char slice of the case's CHALLENGE, never its solution, and produced a
    fabricated "procurement chatbot implementation" for a case that was actually a document-
    processing pipeline with no chatbot anywhere in it. Passing the real solution text closes
    that gap: the model has actual content to describe instead of guessing a plausible-
    sounding mechanism. fit:false (owner-reported, same audit) is the companion fix for
    matches that clear the semantic/lexical bar on a shared word ("demand") but are a
    genuinely different mechanism in a genuinely different domain (a grid load-balancing
    case explained as "capacity and demand matching" for a manufacturer) -- rather than
    write reasoning that glosses over the gap, the model can say the case doesn't actually
    fit, and the caller drops it back to an honest gap.
    Returns {case_id: {"reason","signal","fit"}}. Fails safe to {} (caller uses a
    template AND treats an absent/failed response as fit=true, never auto-drops on
    an AI outage)."""
    items = list(items or [])
    if not items:
        return {}
    brief = brief or {}
    acct = brief.get("account") or {}
    ei = brief.get("expressed_interest") or []
    profile = (profile or "").strip()
    research = (research or "").strip()
    lines = [
        "For each PICKED case study below, write ONE short line (max ~22 words) on why it "
        "fits the client need, and a signal tag. Be specific; do not invent case facts -- "
        "every mechanism/detail you name must come from the case's own CHALLENGE/SOLUTION "
        "text below it, never guessed from the title or need name alone (owner-reported, "
        "2026-07-23: a case with no chatbot anywhere in its real content was explained as "
        "'procurement chatbot implementation' -- invented, not read from the case).",
        "FIT CHECK, before writing the line -- TWO separate questions, both must pass:\n"
        "  (a) MECHANISM: does this case's actual CHALLENGE/SOLUTION genuinely address the "
        "stated need, or does it only share a generic word with it (e.g. a power-grid "
        "'demand response' case matched to a manufacturer's 'material demand planning' need "
        "-- same word 'demand', different mechanism entirely)? Cross-INDUSTRY is fine if the "
        "actual MECHANISM transfers (a contract-intelligence case from construction "
        "genuinely proving contract intelligence for a manufacturer is a real fit); it's the "
        "MECHANISM match that matters, not the industry label.\n"
        "  (b) SUBJECT MATTER PLAUSIBILITY: even when (a) passes, would someone AT THIS "
        "ACCOUNT actually recognize the case's concrete subject as relevant to their "
        "business -- or is it a specific context that has nothing to do with what this "
        "account does, merely filed under the same abstract capability label? Confirmed "
        "miss (owner-reported, 2026-07-23, Vanderlande -- a warehouse/baggage-handling "
        "automation company): a case about validating VEHICLE REGISTRATION documents "
        "(invoice, chassis number, NOC) is a genuine instance of 'compliance automation' in "
        "the abstract (passes test (a)) but has nothing to do with what Vanderlande's "
        "compliance work would ever look like (fails test (b)) -- the label matched, the "
        "actual subject didn't. Passing (a) alone is NOT enough.\n"
        "If EITHER test fails, set \"fit\":false and a one-line \"reason\" saying why (it "
        "will be dropped, not shown to the client) -- do NOT write flattering-sounding but "
        "empty reasoning to paper over a mismatch. Default \"fit\":true only when BOTH "
        "genuinely hold up.",
        "signal: 'capability' (proves the needed capability), 'industry' (also same "
        "industry), 'interest' (a topic the client discussed), 'role' (fits the "
        "stakeholder's function), or 'narrative' (see below).",
        "CITATION CHECK, mandatory, in this order, BEFORE falling back to a generic "
        "capability line (owner-reported, 2026-07-23: our reasoning read as generic next to "
        "a competitor's freeform research, which explicitly quoted the brief's own words for "
        "every pick instead of just describing the case in the abstract):\n"
        "  1. Does the DEEP RESEARCH text below EXPLICITLY tell the seller how to position "
        "or frame THIS capability -- \"should be framed around X\", \"position around Y\", "
        "\"creates demand for Z\"? If yes: signal='narrative', and QUOTE the brief's own "
        "phrase in the reason (e.g. \"The brief itself frames this around EBOM/MBOM "
        "continuity\" beats a generic \"proves PLM capability\").\n"
        "  2. Otherwise, does the case genuinely connect to a SPECIFIC fact in the "
        "STAKEHOLDER BACKGROUND / DEEP RESEARCH text -- a named prior company or role, a "
        "specific project or certification, a named business event (a reorg, a spinoff, a "
        "new plant)? If yes: signal='narrative', and cite that fact plainly (e.g. 'Mirrors "
        "their P&G contract-manufacturing background' not 'Fits their procurement "
        "experience').\n"
        "  3. Only if NEITHER 1 nor 2 genuinely applies, fall back to capability/industry/"
        "interest/role, describing the case's own mechanism instead.\n"
        "Never invent a citation that isn't genuinely there just to avoid step 3 -- an "
        "honest capability line beats a fabricated brief quote every time. This is the SAME "
        "grounding discipline as the fit check above: real text only, never inferred.",
        "Each case below shows its OWN tagged industry in [brackets]. Use signal='industry' "
        "-- and only say the case 'relates to'/'aligns with'/'is relevant to' the account's "
        "industry in the reason text -- when that tagged industry genuinely matches the "
        "ACCOUNT's industry below. If the case's industry is different (or unspecified), "
        "explain it purely by the capability it proves and do NOT claim any industry "
        "connection, even a soft one (owner-reported, 2026-07-20: a case tagged Automotive "
        "was being described as 'relevant to FMCG operations' for an FMCG account -- never "
        "do this).",
        "",
    ]
    if acct.get("industry") or acct.get("role"):
        lines.append(f"ACCOUNT: industry={acct.get('industry','')} role={acct.get('role','')} "
                     f"company_context={acct.get('company_context','')}")
    if ei:
        lines.append("EXPRESSED INTEREST: " + ", ".join(ei))
    if profile:
        lines.append("STAKEHOLDER BACKGROUND (for narrative grounding and citing the "
                     "account's own stated asks -- never a case fact):\n\"\"\"\n"
                     + profile[:6000] + "\n\"\"\"")
    if research:
        # same cap as extract_brief() -- a smaller slice here would cut off exactly the
        # stakeholder-specific section (owner-reported, 2026-07-23: capping this at 6000
        # chars hid Kimberly-Clark's Lovish Jain / Arbex facts from every literal pick's
        # reasoning even though extract_brief() itself saw the full document).
        lines.append("DEEP RESEARCH (for narrative grounding and citing the account's own "
                     "stated asks -- never a case fact):\n"
                     "\"\"\"\n" + research[:RESEARCH_CHAR_CAP] + "\n\"\"\"")
    lines.append("")
    for it in items:
        lines.append(f'NEED "{it.get("need","")}" -> {it["id"]}: {it.get("title","")} '
                     f'[industry: {it.get("industry") or "unspecified"}]. '
                     f'CHALLENGE: {(it.get("blurb","") or "")[:220]} '
                     f'SOLUTION: {(it.get("solution","") or "")[:220]}')
    lines.append('\nReturn ONLY JSON mapping each case id to '
                 '{"reason":"one line","signal":"capability|industry|interest|role|narrative",'
                 '"fit":true}')
    try:
        resp = _client().chat.completions.create(
            model=EXPLAIN_MODEL, temperature=0.2,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You explain why a proof point fits a client "
                 "need, grounded strictly in the case's own challenge/solution text -- and "
                 "fact-check whether it genuinely fits at all before explaining it. Reply "
                 "with one JSON object only."},
                {"role": "user", "content": "\n".join(lines)},
            ],
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        logger.warning("explain_picks: OpenAI call/parse failed — picks keep their "
                       "templated fallback reason instead", exc_info=True)
        return {}
    if not isinstance(data, dict):
        return {}
    valid = {it["id"] for it in items}
    out = {}
    for cid, v in data.items():
        if cid in valid and isinstance(v, dict):
            out[cid] = {"reason": (v.get("reason") or "").strip(),
                        "signal": (v.get("signal") or "").strip().lower(),
                        "fit": v.get("fit") is not False}   # missing/anything-but-False -> True
    return out


def extract_mechanism_tags(case):
    """PILOT, additive (owner-spec, 2026-07-21): read a case's REAL content --
    challenge/solution/capabilities/results, not just its existing keywords or
    the industry it happened to be sold into -- and name the underlying,
    transferable MECHANISM/capability it proves. The gap this closes: MSS022's
    own keyword list is all automotive-flavoured (adas, ev, oem, stamping)
    even though its actual mechanism -- a CFD surrogate model that predicts
    drag/lift from geometry in seconds instead of hours -- is exactly as
    relevant to an aerospace airframe/engine program, but nothing in its tags
    said so, so a GE Aerospace build could never lexically find it.

    Returns {"mechanism_tags": [3-6 short phrases, the ACTUAL technique/
    approach, e.g. "CFD/aerodynamic simulation surrogate model" -- not vague
    restatements like "efficiency" or "innovation"], "transferable_to": one
    line naming 2-4 OTHER industries/domains where this SAME mechanism would
    apply and why -- grounded in the mechanism, never inventing a result or
    client that isn't in the source}. Returns {} on any failure or if the
    case has no real challenge/solution text to ground this in (fail-safe --
    never fabricate tags for a case we can't actually read)."""
    title = (case.get("title") or "").strip()
    challenge = (case.get("challenge") or "").strip()
    solution = (case.get("solution") or "").strip()
    if not (challenge and solution):
        return {}
    caps = case.get("capabilities") or []
    cap_lines = "\n".join(f"- {c.get('title','')}: {c.get('body','')}"
                          for c in caps if isinstance(c, dict))
    results = case.get("results") or []
    results_lines = "\n".join(f"- {r}" for r in results if isinstance(r, str))
    industry = (case.get("industry") or "").strip()
    prompt = (
        "Read this case study's REAL content below and name the underlying "
        "MECHANISM it proves -- the actual technique, approach or system, "
        "described concretely enough that someone could judge whether it "
        "would resonate with a DIFFERENT client in a DIFFERENT industry, not "
        "just a restatement of the industry it happened to be sold into.\n\n"
        f"TITLE: {title}\n"
        f"ORIGINAL INDUSTRY (do not just repeat this as a tag): {industry}\n"
        f"CHALLENGE: {challenge}\n"
        f"SOLUTION: {solution}\n"
        + (f"CAPABILITIES:\n{cap_lines}\n" if cap_lines else "")
        + (f"RESULTS:\n{results_lines}\n" if results_lines else "")
        + "\nReturn ONLY this JSON:\n"
        '{"mechanism_tags": ["...", "..."], "transferable_to": "one line"}\n\n'
        "Rules for mechanism_tags (3-6 items): each is the ACTUAL technique/"
        "approach (e.g. \"CFD/aerodynamic simulation surrogate model\", "
        "\"condition-based predictive maintenance for rotating assets\", "
        "\"RAG copilot over PLM/CAD documentation\") -- never a vague word "
        "like \"efficiency\", \"innovation\", \"automation\" alone, and never "
        "just the original industry restated as a tag.\n"
        "Rule for transferable_to: name 2-4 OTHER industries/domains where "
        "this SAME mechanism (not just a similar-sounding one) would "
        "genuinely apply, and say why in a few words each. If the mechanism "
        "is genuinely industry-specific and doesn't transfer, say so plainly "
        "-- do not force a connection that isn't real."
    )
    try:
        resp = _client().chat.completions.create(
            model=MODEL, temperature=0.2,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You analyse a case study's real content to "
                 "name its transferable underlying mechanism. Ground everything in the "
                 "text given; never invent a client, result or capability that isn't "
                 "there. Reply with one JSON object only."},
                {"role": "user", "content": prompt},
            ],
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        return {}
    if not isinstance(data, dict):
        return {}
    tags = [t.strip() for t in (data.get("mechanism_tags") or [])
           if isinstance(t, str) and t.strip()]
    transferable = (data.get("transferable_to") or "").strip()
    if not tags:
        return {}
    return {"mechanism_tags": tags[:6], "transferable_to": transferable}
    return out
