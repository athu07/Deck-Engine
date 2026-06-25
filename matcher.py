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
    Whole-word match; tags under 3 chars skipped to avoid false hits (AI, ML, QA)."""
    if not transcript:
        return []
    hits = []
    for t in tags:
        if len(t) < 3:
            continue
        if re.search(r"\b" + re.escape(t) + r"\b", transcript):
            hits.append(t)
    return hits


def plan(context, top_n=3, use_ai=False):
    """Returns {'picks','gaps','suggestions','ai_used'}. Picks = slides to include;
    gaps = slots with no good match; use_ai=True refines case picks via the LLM."""
    industry = (context.get("industry") or "").upper()
    functions = {f.strip().upper() for f in (context.get("functions") or []) if f.strip()}
    wanted = {w.strip().upper() for w in (context.get("work_types") or []) if w.strip()}
    transcript_raw = context.get("transcript") or ""
    transcript = transcript_raw.lower()

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
            scored.append((score, hits, row))
        scored.sort(key=lambda x: -x[0])
        best_by_wt[wt] = scored[0][0] if scored else 0
        all_scored.extend(scored)

        rows_by_id = {r["slide_id"]: r for r in rows_}
        ai_pick_ids = ([s for s in ai_cases.get(wt, []) if s in rows_by_id][:top_n]
                       if ai_used else [])

        if ai_pick_ids:                      # AI made usable picks for this work type
            for sid in ai_pick_ids:
                hits = _transcript_hits(_kw_tags(rows_by_id[sid]["keywords"]), transcript)
                why = ("transcript: " + ", ".join(hits[:4])) if hits else "AI-selected"
                chosen[sid] = f"case [{wt}] · {why} (AI)"
                selected_case_wts.add(wt)
        else:                                # keyword fallback (also if AI returned none)
            for score, hits, row in scored[:top_n]:
                why = []
                if hits:
                    why.append("transcript: " + ", ".join(hits[:4]))
                if industry and (row["primary_industry"] or "").upper() == industry:
                    why.append(industry)
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
                          f"{industry or 'this client'} — needs to be created.",
            })
    for sid in chosen:                         # selected but not actually in the deck
        if sid not in lib_ids:
            gaps.append({
                "type": "missing_slide",
                "slide_id": sid,
                "detail": f"{sid} ({titles.get(sid, '?')}) is selected but not "
                          f"found in the deck — needs to be created.",
            })

    # ---- order: natural deck order, except PIN_TO_END slides go last ----
    def order_key(sid):
        if sid in PIN_TO_END:
            return (1, PIN_TO_END.index(sid))   # pinned: after everything, in list order
        return (0, _num(sid))                   # normal: natural deck order
    ordered = sorted(chosen.items(), key=lambda kv: order_key(kv[0]))
    picks = [{"slide_id": sid, "reason": reason} for sid, reason in ordered]

    # ---- "you might also include": next-best related slides, ranked lower ----
    suggested, seen = [], set()
    for score, hits, row in sorted(all_scored, key=lambda x: -x[0]):
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
        suggested.append({"slide_id": sid, "reason": " · ".join(why) or "related"})
        if len(suggested) >= 6:
            break

    suggestions = [
        "Leader slides (CS61 Architects, CS62 The J2W Squad) are never "
        "auto-added — add them in the panel if you want to show people."
    ]
    return {"picks": picks, "gaps": gaps, "suggestions": suggestions,
            "suggested": suggested, "ai_used": ai_used}


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
