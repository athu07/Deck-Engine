# -*- coding: utf-8 -*-
"""
case_library.py  --  The content-store case studies, served in the shape the
rest of the engine already understands.

WHY: case studies now come from case_study_content_store.json (160 records,
ids AIP/WFS/MSS), each rendered fresh from the shared case_study_v2 template at
build time. This module is the single seam between that store and:
  - matcher.py   (auto-pick: scores candidates per work type)
  - app.py       (display titles, and the browsable "add a case" panel)

It deliberately returns matcher-shaped "rows" (slide_id / keywords / primary_*)
so matcher's existing scoring, persona boost, and AI refinement keep working
unchanged — only the SOURCE of candidate case studies moved to the store.
"""

import json

CONTENT_STORE = "case_study_content_store.json"
MIDDOT = "·"   # the keyword separator matcher splits on

_cache = None


def _load():
    global _cache
    if _cache is None:
        try:
            with open(CONTENT_STORE, encoding="utf-8") as f:
                _cache = json.load(f)
        except (OSError, ValueError):
            _cache = []
    return _cache


def _search_text(rec):
    """The full body of a case, for content-based matching — not just the terse
    keyword tags. Titles + keywords are repeated so they weigh a little heavier."""
    caps = " ".join((c.get("title", "") + " " + c.get("body", ""))
                    for c in (rec.get("capabilities") or []))
    kws = " ".join(rec.get("keywords") or [])
    title = rec.get("title", "")
    parts = [
        title, title,                       # title twice (light emphasis)
        rec.get("domain", ""),
        kws,
        rec.get("challenge", ""),
        rec.get("solution", ""),
        caps,
        " ".join(rec.get("results") or []),
    ]
    return " ".join(p for p in parts if p)


def _as_row(rec):
    """Reshape a store record into the row dict matcher/personas read."""
    kws = rec.get("keywords") or []
    return {
        "slide_id":         rec["id"],
        "title":            rec.get("title", ""),
        "keywords":         (" " + MIDDOT + " ").join(kws),
        "primary_industry": rec.get("industry") or "",
        "primary_function": rec.get("function") or "",
        "work_types":       rec.get("work_type") or "",
        "search_text":      _search_text(rec),   # full body for content matching
        "_record":          rec,
    }


def candidate_rows(wanted):
    """{work_type -> [row,...]} for the selected work types (codes like AI_POD)."""
    want = {str(w).strip().upper() for w in (wanted or []) if str(w).strip()}
    out = {}
    for rec in _load():
        wt = (rec.get("work_type") or "").upper()
        if wt in want:
            out.setdefault(wt, []).append(_as_row(rec))
    return out


def all_rows():
    """Every case as a matcher-shaped row, keyed by work type — NO work-type gate.
    The rebuilt matcher scores across all cases and treats the salesperson's
    work-type selection as a boost, not a filter."""
    out = {}
    for rec in _load():
        wt = (rec.get("work_type") or "").upper()
        out.setdefault(wt, []).append(_as_row(rec))
    return out


def title_map():
    """{id -> title} for every store case (for display lookups)."""
    return {r["id"]: r.get("title", "") for r in _load()}


def record(case_id):
    return next((r for r in _load() if r["id"] == case_id), None)


def is_store_id(sid):
    return record(sid) is not None


def all_cases():
    """Light list for the browsable add-a-case panel."""
    out = []
    for r in _load():
        out.append({
            "id":        r["id"],
            "title":     r.get("title", ""),
            "domain":    r.get("domain", ""),
            "work_type": r.get("work_type", ""),
            "industry":  r.get("industry") or "",
            "function":  r.get("function") or "",
        })
    return out


if __name__ == "__main__":
    cases = all_cases()
    print(f"{len(cases)} case studies in the store")
    by_wt = {}
    for c in cases:
        by_wt[c["work_type"]] = by_wt.get(c["work_type"], 0) + 1
    for wt, n in sorted(by_wt.items()):
        print(f"  {wt:12} {n}")
