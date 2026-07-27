# -*- coding: utf-8 -*-
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from deckengine import config  # noqa: E402  (repo-root data paths)
"""
enrich_mechanism_tags.py  --  PILOT, additive (owner-spec, 2026-07-21).

Reads every case study for one work type (default: MS Solution, "MSS...") from
the content store, and for each asks ai_matcher.extract_mechanism_tags() to
name the case's underlying, transferable MECHANISM from its real challenge/
solution/capabilities/results text -- independent of which industry it was
originally sold into (e.g. MSS022 is tagged "automotive OEM" but its real
mechanism, a CFD surrogate model, is exactly as relevant to aerospace).

Writes ONLY to case_mechanism_tags.json (config.CASE_MECHANISM_TAGS_JSON) --
a brand new, separate file. Never touches case_study_content_store.json, the
existing "keywords" field, or the source Excel. Re-running is safe: it
overwrites only its own file, and only for the work type requested; entries
for other work types already in the file are preserved.

Not yet wired into matching (relevance.py) -- this is the PILOT generation
step, for the owner to review before deciding whether/how to use it.

Run (repo root):
    py scripts/enrich_mechanism_tags.py            # MS Solution only (pilot)
    py scripts/enrich_mechanism_tags.py MS AIP WFS  # any subset of work types
"""

import io
import json
import sys

from deckengine.services.matching import ai_matcher

STORE = config.CONTENT_STORE_JSON
OUT = config.CASE_MECHANISM_TAGS_JSON

_PREFIX_OF_WT = {"MS": "MSS", "AIP": "AIP", "WFS": "WFS"}


def main(work_types):
    prefixes = tuple(_PREFIX_OF_WT[w] for w in work_types)
    recs = json.load(open(STORE, encoding="utf-8"))
    targets = [r for r in recs if (r.get("id") or "").startswith(prefixes)]
    print(f"Enriching {len(targets)} case(s) for work type(s) {work_types} ...")

    try:
        existing = json.load(open(OUT, encoding="utf-8"))
    except (OSError, ValueError):
        existing = {}

    ok, skipped = 0, []
    for i, r in enumerate(targets, 1):
        cid = r["id"]
        result = ai_matcher.extract_mechanism_tags(r)
        if result:
            existing[cid] = {**result, "title": r.get("title", "")}
            ok += 1
        else:
            skipped.append(cid)
        print(f"  ...{i}/{len(targets)}  {cid}"
              f"{'  (skipped -- no usable content or API error)' if not result else ''}")

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(existing, f, indent=2, ensure_ascii=False)

    print(f"\nSaved {ok} enriched, {len(skipped)} skipped -> {OUT}")
    if skipped:
        print("Skipped:", ", ".join(skipped))


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    wts = [w.upper() for w in sys.argv[1:]] or ["MS"]
    bad = [w for w in wts if w not in _PREFIX_OF_WT]
    if bad:
        print(f"Unknown work type(s): {bad}. Use MS, AIP, and/or WFS.")
        sys.exit(1)
    main(wts)
