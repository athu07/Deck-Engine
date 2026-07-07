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
import re

from deckengine import config

CONTENT_STORE = config.CONTENT_STORE_JSON
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


# ---------------------------------------------------------------------------
# Auto-save an accepted "Create with AI" case study into the shared library
# (owner's choice: automatic on every accept, no extra step -- see decks.py's
# /finalize, which calls promote_ai_case() for every NEW:<id> in the final deck).
# ---------------------------------------------------------------------------
_WT_PREFIX = {"WORKFORCE": "WFS", "AI_POD": "AIP", "MS": "MSS"}
_WT_LABEL = {"WORKFORCE": "Workforce Solutions", "AI_POD": "AI Pods", "MS": "MS Solution"}


def _next_id(prefix):
    nums = []
    for r in _load():
        rid = r.get("id", "")
        if rid.startswith(prefix) and rid[len(prefix):].isdigit():
            nums.append(int(rid[len(prefix):]))
    return f"{prefix}{(max(nums) + 1 if nums else 1):03d}"


def _derive_tags(title, challenge, solution, capabilities, results, work_type_code):
    """The same style of function/persona/keyword tagging the Excel importer
    uses (see scripts/build_case_study_store.py), applied to ONE AI-drafted
    case. No industry inference needed here -- Create-with-AI already carries
    a clean form industry CODE, unlike the Excel importer's raw domain text."""
    from deckengine.services.matching.tagger import FUNCTION, _score
    from deckengine.services.matching import personas as _personas
    from deckengine.services.matching import synonyms as _synonyms
    from deckengine.services.rendering.fill_case_study import split_capability

    # capabilities may be plain "Title: body" strings (fresh AI draft) OR
    # {title, body} dicts (if the salesperson edited them on /review) --
    # split_capability() already handles both.
    cap_pairs = [split_capability(c) for c in (capabilities or [])]
    cap_titles = [t for t, _b in cap_pairs]
    cap_text = " ".join(t + " " + b for t, b in cap_pairs)
    fulltext = " ".join([title, challenge, solution, cap_text,
                        " ".join(results or [])]).lower()
    function, _votes = _score(" " + fulltext + " ", FUNCTION)
    persona_codes = _personas.tag_slide({
        "primary_function": function or "", "work_types": work_type_code,
        "keywords": title, "title": title,
    })
    words = set()
    for src in [title] + cap_titles:
        for w in re.findall(r"[A-Za-z][A-Za-z\-]+", src):
            if len(w) >= 3:
                words.add(w.lower())
    keywords = set()
    for w in words:
        keywords |= _synonyms.expand(w)
    return function, persona_codes, sorted(keywords)


def _append_to_excel(record):
    """Best-effort: append a row to the source Excel too, so a from-scratch
    rebuild (build_case_study_store.py) reproduces this case. Never blocks the
    deck finalize if the file is locked/unwritable (e.g. open in Excel)."""
    import openpyxl
    path = config.CASE_STUDIES_SOURCE_XLSX
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        kw_cell = ", ".join([record["domain"]] + record["keywords"][:10])
        # capabilities are stored as {title, body} dicts -- reflow to the Excel's
        # own "Title: description" one-per-line convention
        cap_lines = [f"{c['title']}: {c['body']}" if c.get("body") else c["title"]
                    for c in record["capabilities"]]
        ws.append([
            record["id"], record["work_type_label"], kw_cell, record["title"],
            record["challenge"], record["solution"],
            "\n".join(cap_lines), "; ".join(record["results"]), "Yes",
        ])
        wb.save(path)
        return True
    except Exception:
        return False   # fail safe -- the store/JSON entry already exists regardless


def _append_embedding(record):
    """Best-effort: one OpenAI embedding call so the new case is semantically
    matchable immediately (not just by keyword) without waiting for a full
    py build_case_embeddings.py rebuild. Fails safe -- relevance.py's embedding
    lookups are all .get()-based, so a case with no vector yet just scores on
    lexical/keyword signals alone until embeddings are next rebuilt."""
    try:
        from deckengine.services.infra import load_env
        load_env()
        from openai import OpenAI
        client = OpenAI()
        text = _search_text(record)[:8000]
        resp = client.embeddings.create(model="text-embedding-3-small", input=[text])
        vec = resp.data[0].embedding

        try:
            with open(config.CASE_EMBEDDINGS_JSON, encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError):
            data = {"model": "text-embedding-3-small", "dim": len(vec), "vectors": {}}
        data.setdefault("vectors", {})[record["id"]] = vec
        with open(config.CASE_EMBEDDINGS_JSON, "w", encoding="utf-8") as f:
            json.dump(data, f)
        return True
    except Exception:
        return False


def promote_ai_case(content, work_type_code, industry_code):
    """Auto-save an accepted 'Create with AI' case study into the shared content
    store (JSON + Excel + embeddings), so it's searchable/reusable across future
    projects. `content` is the staged draft (title/challenge/solution/
    capabilities/results, from slide_generator.draft_case_study). `work_type_code`
    (WORKFORCE/AI_POD/MS) decides the new id's prefix -- if the deck has more than
    one work type selected, the caller passes the FIRST one (best-effort; the
    content itself doesn't commit to a single work type). Returns the new id, or
    None if work_type_code is blank/unrecognised (never guesses a prefix)."""
    prefix = _WT_PREFIX.get((work_type_code or "").upper())
    if not prefix:
        return None

    from deckengine.services.rendering.fill_case_study import split_capability

    title = content.get("title", "")
    challenge = content.get("challenge", "")
    solution = content.get("solution", "")
    raw_capabilities = content.get("capabilities") or []
    results = content.get("results") or []
    domain = (industry_code or "").replace("_", " ").title()

    function, persona_codes, keywords = _derive_tags(
        title, challenge, solution, raw_capabilities, results, work_type_code.upper())
    # normalise capabilities to {title, body} dicts -- the same shape every other
    # entry in the store uses, regardless of whether this came in as plain
    # "Title: body" strings (fresh draft) or dicts (edited on /review)
    capabilities = [{"title": t, "body": b} for t, b in
                    (split_capability(c) for c in raw_capabilities)]

    new_id = _next_id(prefix)
    record = {
        "id": new_id,
        "work_type": work_type_code.upper(),
        "work_type_label": _WT_LABEL.get(work_type_code.upper(), work_type_code),
        "title": title,
        "raw_title": title,
        "client_descriptor": "",
        "domain": domain,
        "industry": (industry_code or "").upper(),
        "function": function,
        "personas": persona_codes,
        "curated_keywords": [],
        "keywords": keywords,
        "challenge": challenge,
        "solution": solution,
        "capabilities": capabilities,
        "results": results,
        "ai_generated": True,
        "source_row": None,
    }

    store = _load()
    store.append(record)
    with open(CONTENT_STORE, "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=2)
    global _cache
    _cache = store   # keep the in-process cache in sync so this build's own
                     # re-reads (and the next one) see the new case immediately

    _append_to_excel(record)
    _append_embedding(record)
    return new_id


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
