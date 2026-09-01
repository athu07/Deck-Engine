# -*- coding: utf-8 -*-
"""library.py  --  the /library page (all master slides + content-store cases)."""

import json

from flask import Blueprint, render_template

from deckengine import config
from deckengine.services.content import case_library
from deckengine.services.matching import matcher
from .view_helpers import shell, legacy_case_ids

bp = Blueprint("library", __name__)


@bp.route("/library")
def library():
    slides = []
    # 1) master slides EXCEPT the legacy case-study slides — those are superseded
    #    by the content store and shown below (rendered via the new branded
    #    template), so a case study is never the old master version anywhere.
    legacy = legacy_case_ids()
    try:
        recs = json.load(open(config.TAGGED_LIBRARY_JSON, encoding="utf-8"))
    except Exception:
        recs = []
    # The card heading should read as the slide's NAME, not whatever text build_
    # library.build() happened to pull from the first shape it found -- for any
    # slide whose visible text starts with a number/photo/label rather than a
    # title (e.g. CS09 = "01", CS61 = "Photo", CS140 = "BOT"), that auto-extract
    # is noise, not a name. The Slide Registry's own `title` column is the
    # human-curated name every one of these slides already has (owner-spec,
    # 2026-08-13) -- prefer it, falling back to the auto-extracted text only if
    # a slide somehow has no registry row or a blank title there.
    registry_titles = {row["slide_id"]: (row.get("title") or "").strip()
                       for row in matcher.load_registry()}
    for r in recs:
        if r["slide_id"] in legacy:
            continue
        t = r.get("tags", {})
        title = registry_titles.get(r["slide_id"]) or r.get("title", "")
        slides.append({
            "id": r["slide_id"], "title": title,
            "wt": t.get("work_type", {}).get("value") or "",
            "kind": t.get("kind", {}).get("value") or "",
            "ind": t.get("industry", {}).get("value") or "",
            "fn": t.get("function", {}).get("value") or "",
            "kw": r.get("keywords", [])[:6],
            "search": (r["slide_id"] + " " + title + " " + r.get("title", "") + " " +
                       " ".join(r.get("keywords", []))).lower(),
        })
    # 2) every content-store case study — these download/add as the NEW branded
    #    case_study_v2 template (see /slide/<id>/download and finalize).
    for rec in case_library._load():
        kws = rec.get("keywords") or []
        slides.append({
            "id": rec["id"], "title": rec.get("title", ""),
            "wt": rec.get("work_type") or "",
            "kind": "CASE_STUDY",
            "ind": rec.get("industry") or "",
            "fn": rec.get("function") or "",
            "kw": kws[:6],
            "search": (rec["id"] + " " + rec.get("title", "") + " " +
                       rec.get("domain", "") + " " + " ".join(kws)).lower(),
        })
    industries = sorted({s["ind"] for s in slides if s["ind"]})
    body = render_template("library.html", slides=slides, industries=industries, total=len(slides))
    return shell(body, active="library", crumb="<b>Library</b> / All slides")
