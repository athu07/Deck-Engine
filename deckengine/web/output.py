# -*- coding: utf-8 -*-
"""output.py  --  serve built decks (/output/<file>) and single-slide downloads."""

import os

from flask import Blueprint, request, send_file, abort

from deckengine.constants import OUTPUT_DIR
from deckengine.services.content import case_library
from deckengine.services.rendering import assembler, fill_case_study
from .view_helpers import file_busy_page

bp = Blueprint("output", __name__)


@bp.route("/output/<path:fname>")
def output_file(fname):
    fname = os.path.basename(fname)
    path = os.path.join(OUTPUT_DIR, fname)
    if not os.path.exists(path):
        abort(404)
    return send_file(path, as_attachment=bool(request.args.get("dl")), download_name=fname)


@bp.route("/slide/<sid>/download")
def slide_download(sid):
    sid = sid.upper()
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, f"Slide_{sid}.pptx")
    # a content-store case study (AIP/WFS/MSS) -> render fresh from the NEW
    # branded case_study_v2 template, never the old master version
    rec = case_library.record(sid)
    if rec is not None:
        try:
            fill_case_study.fill_row(rec, path)
        except (PermissionError, OSError) as e:
            return file_busy_page(str(e))
        return send_file(path, as_attachment=True, download_name=f"{sid}.pptx")
    # otherwise a master (standard/structural) slide -> build from the master deck
    kept, _ = assembler.build_deck([sid], out=path)
    if not kept:
        abort(404)
    return send_file(path, as_attachment=True, download_name=f"{sid}.pptx")
