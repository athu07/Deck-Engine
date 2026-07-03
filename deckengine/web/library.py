# -*- coding: utf-8 -*-
"""library.py  --  the /library page (all master slides + content-store cases)."""

import json

from flask import Blueprint, render_template

from deckengine import config
from deckengine.services.content import case_library
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
    for r in recs:
        if r["slide_id"] in legacy:
            continue
        t = r.get("tags", {})
        slides.append({
            "id": r["slide_id"], "title": r.get("title", ""),
            "wt": t.get("work_type", {}).get("value") or "",
            "kind": t.get("kind", {}).get("value") or "",
            "ind": t.get("industry", {}).get("value") or "",
            "fn": t.get("function", {}).get("value") or "",
            "kw": r.get("keywords", [])[:6],
            "search": (r["slide_id"] + " " + r.get("title", "") + " " +
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
