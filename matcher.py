# -*- coding: utf-8 -*-
"""
matcher.py  --  Box 1 (form-only, rules-only): client context -> slide_ids.

Given a context the sales team fills in (industry + which work types), this
returns the ordered list of slide IDs that should go into the deck, following
the registry's Selection Rules. NO transcript, NO AI, NO API cost.

Selection Rules (from the registry's READ ME):
  1. ALWAYS slides (CORE CS01-08 + brand CS20) -> always in.
  2. For each selected work type, include its STANDARD block + section divider.
  3. CASE_STUDY slides in a selected work type -> scored by industry / function /
     keyword fit; keep the top few.
  4. A case-study DIVIDER is included only if >=1 case under it was selected.
  5. OPTIONAL slides (leaders, extras) -> skipped here; they're decided later
     from the transcript / by asking the user.

The registry drives the rules (its include_rule column); the content library
(tagged_library.json) supplies each slide's pointer into the deck.
"""

import json
import re
import sys

import openpyxl

import synonyms   # equivalence groups so paraphrased topics still match
import personas   # buyer-role detection + persona-relevance scoring

MIDDOT = "·"   # the keyword-string separator on the slides

REGISTRY_XLSX = "J2W_CaseStudy_Portfolio_Metadata.xlsx"
LIBRARY_JSON = "tagged_library.json"

# Slides forced to the VERY END of the deck, in this exact order.
# Everything else stays in natural deck order. Edit this list to reorder the close.
PIN_TO_END = ["CS07", "CS08"]   # Next Steps -> Let's win together

# Slides to NEVER auto-include, even if a registry rule would pick them.
# These are DIVIDERS — reference-only markers used to classify slide types, NOT
# content to put in decks (per the deck owner). CS09/14/17 = the 01/02/03 section
# dividers; CS20 = J2W brand; CS21/35/45 = case-study section dividers.
EXCLUDE = {"CS09", "CS14", "CS17", "CS20", "CS21", "CS35", "CS45"}

# People / leader slides — never auto-picked; always surfaced for the user to add.
LEADER_IDS = {"CS61", "CS62"}


def _num(slide_id):
    """CS22 -> 22, for ordering."""
    try:
        return int("".join(ch for ch in slide_id if ch.isdigit()))
    except ValueError:
        return 9999


def load_registry():
    ws = openpyxl.load_workbook(REGISTRY_XLSX, data_only=True)["Slide Registry"]
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[0]
    col = {h: i for i, h in enumerate(hdr)}
    out = []
    for r in rows[1:]:
        out.append({k: r[col[k]] for k in col})
    return out


def _wts_of(row):
    raw = row.get("work_types") or ""
    return {w.strip().upper() for w in raw.replace("|", ",").split(",") if w.strip()}


def _kw_tags(keyword_string):
    """Split a slide's keyword string into individual lowercase tags."""
    return [t.strip().lower() for t in (keyword_string or "").split(MIDDOT) if t.strip()]


def _transcript_hits(tags, transcript):
    """Which of this slide's keyword tags actually appear in the transcript.
    A tag 'hits' if the tag OR any of its synonyms shows up as a whole word —
    so "automated QA" in the notes still matches a slide tagged "test automation".
    Tags under 3 chars are skipped to avoid false hits (AI, ML, QA)."""
    if not transcript:
        return []
    hits = []
    for t in tags:
        if len(t) < 3:
            continue
        if synonyms.hits_in(t, transcript):   # tag + all its synonyms
            hits.append(t)
    return hits


# Generic words that are never a real "capability ask" even if they appear.
_ASK_STOPWORDS = {
    "software", "team", "teams", "quality", "support", "project", "projects",
    "solution", "solutions", "technology", "tech", "help", "service", "services",
    "platform", "system", "systems", "tool", "tools", "data", "people", "process",
    "delivery", "engineering", "development", "work", "business", "company",
    "client", "customer", "product", "experience", "capability", "capabilities",
}

# Individual tools / libraries / frameworks / languages are DETAILS, not deck
# themes — never flag them as a missing capability (they belong to a capability
# we already cover, e.g. Selenium -> test automation). Belt-and-suspenders behind
# the AI extractor, which is also told to skip these.
_ASK_TOOLS = {
    # testing tools
    "cucumber", "selenium", "junit", "testng", "cypress", "playwright", "appium",
    "pytest", "katalon", "soapui", "specflow", "rest assured", "robot framework",
    "jmeter", "loadrunner", "postman", "mocha", "jasmine", "karate", "bdd", "tdd",
    # languages
    "java", "python", "javascript", "typescript", "c#", ".net", "c++", "golang",
    "go", "ruby", "php", "scala", "kotlin", "swift", "rust",
    # frameworks / libraries / dev tools
    "react", "angular", "vue", "svelte", "spring", "spring boot", "django",
    "flask", "node", "nodejs", "express", "rails", "hibernate", "maven", "gradle",
    "jenkins", "docker", "kubernetes", "ansible", "terraform", "git", "github",
    "gitlab", "jira",
}


def slugify(text):
    """ADAS / fraud detection -> a safe form-field id (adas / fraud-detection)."""
    s = re.sub(r"[^a-z0-9]+", "-", (text or "").lower()).strip("-")
    return s or "topic"


def _known_topics(reg):
    """Every distinct keyword tag across the registry (lowercased, >=3 chars).
    This is the controlled capability vocabulary the transcript is scanned against."""
    topics = set()
    for row in reg:
        for t in _kw_tags(row.get("keywords")):
            if len(t) >= 3 and t not in _ASK_STOPWORDS:
                topics.add(t)
    return topics


def _coverage_terms(topic):
    """Terms to search slides for when deciding if an ask is covered:
    the ask itself (+ its synonyms) PLUS any known concept that appears INSIDE
    the ask (+ its synonyms). This is what lets 'CI/CD setup' resolve to the
    'ci/cd' concept (the AI often appends words like setup/support/migration).
    Generic stopwords are never used as a coverage term."""
    low = (topic or "").strip().lower()
    terms = set(synonyms.expand(low))
    for kt in synonyms.known_terms():
        if len(kt) >= 4 and kt not in _ASK_STOPWORDS and \
                re.search(r"\b" + re.escape(kt) + r"\b", low):
            terms |= synonyms.expand(kt)
    return {t for t in terms if len(t) >= 3 and t not in _ASK_STOPWORDS}


def _slides_covering(topic, reg):
    """Slide IDs whose keywords OR title cover this topic (synonym-aware, and
    aware of a known concept embedded in a longer ask), excluding dividers/brand
    slides we never put in a deck."""
    terms = _coverage_terms(topic)
    out = []
    for row in reg:
        sid = row["slide_id"]
        if sid in EXCLUDE:
            continue
        text = ((row.get("keywords") or "") + " " + (row.get("title") or "")).lower()
        if any(re.search(r"\b" + re.escape(t) + r"\b", text) for t in terms):
            out.append(sid)
    return out


def _is_confident_ask(ask):
    """Conservative filter: keep multi-word asks or specific single terms; drop
    generic words and noise. Acronyms (ADAS, SAP) are kept."""
    a = (ask or "").strip()
    if len(a) < 3:
        return False
    low = a.lower()
    if low in _ASK_STOPWORDS or low in _ASK_TOOLS:   # generic word or a bare tool
        return False
    if " " in a:                       # multi-word phrase -> specific enough
        return True
    if a.isupper() and len(a) <= 6:    # acronym like ADAS, SAP, ETL
        return True
    return len(low) >= 4               # a single concrete word


def _dedupe_asks(asks):
    """Collapse asks that are the same concept (synonym-equivalent) or duplicates.
    Keeps the first-seen surface form."""
    out, seen_groups = [], []
    for a in asks:
        forms = synonyms.expand(a.lower())
        if any(forms & g for g in seen_groups):
            continue
        seen_groups.append(set(forms))
        out.append(a)
    return out


def _capability_gaps(asks, reg, chosen, max_missing=5, max_suggest=6):
    """Classify each client ask (transcript-first):
       - covered by a PICKED slide  -> answered, ignore
       - covered by an UNPICKED slide -> 'asked, add this existing slide'
       - covered by NO slide        -> a real gap ('asked but not in the deck')
    Returns (missing_topics, asked_existing) where asked_existing is
    [{'slide_id','topic','reason'}]."""
    picked = set(chosen)
    missing, asked_existing, used_slides = [], [], set()
    for ask in asks:
        slides = _slides_covering(ask, reg)
        if not slides:
            missing.append(ask)
        elif any(s in picked for s in slides):
            continue                                  # already answered in the deck
        else:
            sid = next((s for s in slides if s not in used_slides), slides[0])
            used_slides.add(sid)
            asked_existing.append({
                "slide_id": sid, "topic": ask,
                "reason": f"“{ask}” was asked — this slide covers it",
            })
    return missing[:max_missing], asked_existing[:max_suggest]


def plan(context, top_n=3, use_ai=False):
    """Returns {'picks','gaps','suggestions','ai_used'}. Picks = slides to include;
    gaps = slots with no good match; use_ai=True refines case picks via the LLM."""
    industry = (context.get("industry") or "").upper()
    functions = {f.strip().upper() for f in (context.get("functions") or []) if f.strip()}
    wanted = {w.strip().upper() for w in (context.get("work_types") or []) if w.strip()}
    transcript_raw = context.get("transcript") or ""
    transcript = transcript_raw.lower()

    # Persona = WHO we're meeting (recipient field) + anyone named in the notes.
    # Used to nudge persona-relevant case studies up the ranking.
    persona_codes = personas.detect(context.get("recipient", ""), transcript_raw)

    reg = load_registry()
    chosen = {}                 # slide_id -> reason
    cases_by_wt = {}            # work_type -> [rows]

    # ---- pass 1: always-in, std blocks, and gather candidate case studies ----
    for row in reg:
        sid = row["slide_id"]
        rule = (row["include_rule"] or "").strip()
        kind = row["kind"]
        row_wts = _wts_of(row)

        if rule.upper() == "ALWAYS":
            chosen[sid] = "core / always"
        elif rule.startswith("IF work_type includes"):
            if row_wts & wanted:
                chosen[sid] = f"standard block ({'/'.join(sorted(row_wts & wanted))})"
        elif kind == "CASE_STUDY":
            for wt in (row_wts & wanted):
                cases_by_wt.setdefault(wt, []).append(row)

    # ---- optional AI refinement from the transcript (opt-in, fails safe) ----
    ai_used, ai_cases, ai_optional = False, {}, []
    if use_ai and transcript.strip():
        try:
            import ai_matcher
            cand = {wt: [{"slide_id": r["slide_id"], "title": r["title"],
                          "keywords": r["keywords"]} for r in rows]
                    for wt, rows in cases_by_wt.items()}
            opt_rows = [r for r in reg
                        if (r["include_rule"] or "").strip().upper().startswith("OPTIONAL")
                        and r["slide_id"] not in LEADER_IDS]
            opt = [{"slide_id": r["slide_id"], "title": r["title"]} for r in opt_rows]
            res = ai_matcher.refine(transcript_raw, cand, opt, top_n=top_n)
            ai_cases = res.get("cases", {}) or {}
            ai_optional = res.get("optional", []) or []
            ai_used = True
        except Exception:
            ai_used = False                 # any failure -> fall back to keywords

    # ---- pass 2: choose case studies per selected work type ----
    selected_case_wts = set()
    best_by_wt = {}                      # work_type -> best case-study score seen
    all_scored = []                      # every scored candidate, for suggestions
    for wt, rows_ in cases_by_wt.items():
        scored = []
        for row in rows_:
            kw = (row["keywords"] or "").lower()
            hits = _transcript_hits(_kw_tags(row["keywords"]), transcript)
            score = 3 * len(hits)                 # transcript overlap = primary signal
            if industry and (row["primary_industry"] or "").upper() == industry:
                score += 2
            if industry and industry.lower() in kw:
                score += 1
            if functions and (row["primary_function"] or "").upper() in functions:
                score += 1
            p_boost, p_why = personas.score_boost(persona_codes, row)
            score += p_boost                      # persona relevance (capped)
            scored.append((score, hits, row, p_why))
        scored.sort(key=lambda x: -x[0])
        best_by_wt[wt] = scored[0][0] if scored else 0
        all_scored.extend(scored)

        rows_by_id = {r["slide_id"]: r for r in rows_}
        ai_pick_ids = ([s for s in ai_cases.get(wt, []) if s in rows_by_id][:top_n]
                       if ai_used else [])

        if ai_pick_ids:                      # AI made usable picks for this work type
            for sid in ai_pick_ids:
                hits = _transcript_hits(_kw_tags(rows_by_id[sid]["keywords"]), transcript)
                _, p_why = personas.score_boost(persona_codes, rows_by_id[sid])
                why = ("transcript: " + ", ".join(hits[:4])) if hits else "AI-selected"
                if p_why:
                    why += " · for " + "/".join(sorted(set(p_why)))
                chosen[sid] = f"case [{wt}] · {why} (AI)"
                selected_case_wts.add(wt)
        else:                                # keyword fallback (also if AI returned none)
            for score, hits, row, p_why in scored[:top_n]:
                why = []
                if hits:
                    why.append("transcript: " + ", ".join(hits[:4]))
                if industry and (row["primary_industry"] or "").upper() == industry:
                    why.append(industry)
                if p_why:
                    why.append("for " + "/".join(sorted(set(p_why))))
                note = " · ".join(why) if why else "WEAK match — review"
                chosen[row["slide_id"]] = f"case [{wt}] · {note} (score {score})"
                selected_case_wts.add(wt)

    # ---- pass 3: case-section dividers (only if a child case was chosen) ----
    for row in reg:
        rule = (row["include_rule"] or "").strip()
        if rule.startswith("IF >=1"):
            if _wts_of(row) & selected_case_wts:
                chosen[row["slide_id"]] = "case-section divider"

    # ---- AI-chosen OPTIONAL slides (transcript-relevant, non-leader) ----
    if ai_used and ai_optional:
        valid_opt = {r["slide_id"] for r in reg
                     if (r["include_rule"] or "").strip().upper().startswith("OPTIONAL")
                     and r["slide_id"] not in LEADER_IDS}
        for sid in ai_optional:
            if sid in valid_opt:
                chosen[sid] = "optional · AI (transcript-relevant)"

    # ---- drop any slides we never auto-include (e.g. the J2W brand divider) ----
    for sid in EXCLUDE:
        chosen.pop(sid, None)

    # ---- gaps: needed slots with no good match -> "needs to be created" ----
    titles = _title_lookup()
    lib_ids = set(titles)
    gaps = []
    GOOD = 2   # a good proof slide has at least an industry match or one transcript hit
    for wt in sorted(wanted):
        if best_by_wt.get(wt, 0) < GOOD:
            gaps.append({
                "type": "needs_case_study",
                "work_type": wt,
                "detail": f"No strong {wt} case study for "
                          f"{industry or 'this client'} in the library yet.",
            })
    for sid in chosen:                         # selected but not actually in the deck
        if sid not in lib_ids:
            gaps.append({
                "type": "missing_slide",
                "slide_id": sid,
                "detail": f"{sid} ({titles.get(sid, '?')}) is selected but not "
                          f"found in the deck — needs to be created.",
            })

    # ---- capability gaps: per-TOPIC asks (transcript-first, then suggest) ----
    # Closes the loophole where a client ask (e.g. "ADAS") silently goes
    # unanswered. Hybrid detection: known keyword/synonym hits + a conservative
    # AI extraction that can surface a brand-new ask absent from every slide.
    asked_existing = []
    if transcript.strip():
        det = sorted(t for t in _known_topics(reg) if synonyms.hits_in(t, transcript))
        ai_asks = []
        if use_ai:
            try:
                import ai_matcher
                ai_asks = ai_matcher.extract_asks(transcript_raw)
            except Exception:
                ai_asks = []
        # AI asks first (these can be genuinely novel/missing), then known hits.
        asks = _dedupe_asks([a for a in (ai_asks + det) if _is_confident_ask(a)])
        missing_topics, asked_existing = _capability_gaps(asks, reg, chosen)
        default_wt = (sorted(wanted)[0] if wanted else "AI_POD")
        for topic in missing_topics:
            gaps.append({
                "type": "missing_capability",
                "topic": topic,
                "slug": slugify(topic),
                "work_type": default_wt,
                "detail": f"“{topic}” was asked in the meeting but isn’t in the deck.",
            })

    # ---- order: natural deck order, except PIN_TO_END slides go last ----
    def order_key(sid):
        if sid in PIN_TO_END:
            return (1, PIN_TO_END.index(sid))   # pinned: after everything, in list order
        return (0, _num(sid))                   # normal: natural deck order
    ordered = sorted(chosen.items(), key=lambda kv: order_key(kv[0]))
    picks = [{"slide_id": sid, "reason": reason} for sid, reason in ordered]

    # ---- "you might also include": next-best related slides, ranked lower ----
    # Asks that an EXISTING-but-unpicked slide already answers go FIRST and are
    # flagged as "asked in the meeting" so the salesperson pulls them in.
    suggested, seen = [], set()
    for item in asked_existing:
        sid = item["slide_id"]
        if sid in chosen or sid in seen or sid in EXCLUDE:
            continue
        seen.add(sid)
        suggested.append({"slide_id": sid, "reason": item["reason"], "asked": True})

    for score, hits, row, p_why in sorted(all_scored, key=lambda x: -x[0]):
        sid = row["slide_id"]
        if score <= 0 or sid in chosen or sid in seen or sid in EXCLUDE:
            continue
        seen.add(sid)
        why = []
        if hits:
            why.append("matches your notes")
        if industry and (row["primary_industry"] or "").upper() == industry:
            why.append("same industry")
        if functions and (row["primary_function"] or "").upper() in functions:
            why.append("same function")
        if p_why:
            why.append("for " + "/".join(sorted(set(p_why))))
        suggested.append({"slide_id": sid, "reason": " · ".join(why) or "related"})
        if len(suggested) >= 6:
            break

    suggestions = [
        "Leader slides (CS61 Architects, CS62 The J2W Squad) are never "
        "auto-added — add them in the panel if you want to show people."
    ]
    if persona_codes:
        suggestions.insert(0,
            "Tuned for " + ", ".join(personas.labels(persona_codes)) +
            " — persona-relevant case studies were ranked higher.")
    return {"picks": picks, "gaps": gaps, "suggestions": suggestions,
            "suggested": suggested, "ai_used": ai_used,
            "persona_codes": persona_codes,
            "persona_labels": personas.labels(persona_codes)}


def match(context, top_n=3):
    """Backward-compatible: just the picks (used by the assembler)."""
    return plan(context, top_n)["picks"]


def _title_lookup():
    lib = json.load(open(LIBRARY_JSON, encoding="utf-8"))
    return {r["slide_id"]: r.get("title", "") for r in lib}


if __name__ == "__main__":
    # Demo context — edit these to try different scenarios.
    context = {
        "client_name": "Acme Bank",
        "industry": "BFSI",
        "work_types": ["AI_POD", "WORKFORCE"],
        "deck_phase": "intro",
        "recipient": "CTO",
    }
    titles = _title_lookup()
    picks = match(context)
    print("CONTEXT:", json.dumps(context))
    print(f"\nDECK PLAN — {len(picks)} slides:\n")
    for p in picks:
        print(f"  {p['slide_id']:5} {titles.get(p['slide_id'],'')[:46]:46} | {p['reason']}")
