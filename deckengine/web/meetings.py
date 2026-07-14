# -*- coding: utf-8 -*-
"""meetings.py  --  the /meetings deck-repository page (search past decks)."""

import os

from flask import Blueprint, request, render_template

from deckengine import config
from deckengine import constants
from deckengine.services import meeting_log
from deckengine.constants import WORK_TYPES, PHASES, WT_LABELS
from .view_helpers import shell

bp = Blueprint("meetings", __name__)


@bp.route("/meetings")
def meetings():
    f_ind = request.args.get("industry", "").strip()
    f_wt = request.args.get("work_type", "").strip()
    f_phase = request.args.get("phase", "").strip()
    rows = meeting_log.all_meetings()          # every version, newest first
    if f_ind:
        rows = [r for r in rows if r.get("industry") == f_ind]
    if f_wt:
        rows = [r for r in rows if f_wt in r.get("work_types", [])]
    if f_phase:
        rows = [r for r in rows if r.get("phase") == f_phase]
    # "Reopen to edit" only makes sense on the LATEST version of a client+phase
    # (reopening always continues from the newest, never forks an older one)
    latest_version_by_key = {(r["client"], r["phase"]): r["version"] for r in meeting_log.all_latest()}
    for r in rows:
        r["is_latest"] = r.get("version") == latest_version_by_key.get((r.get("client"), r.get("phase")))
        # a logged version whose .pptx is no longer on disk (owner-reported,
        # 2026-07-14) -- flag it instead of a link that 404s when clicked
        r["file_missing"] = not os.path.isfile(os.path.join(config.OUTPUT_DIR, r.get("deck_file", "")))
    body = render_template("meetings.html", rows=rows, total=len(rows),
                                  industries=constants.all_industries(), work_types=WORK_TYPES,
                                  phases=PHASES, wt_labels=WT_LABELS,
                                  f_ind=f_ind, f_wt=f_wt, f_phase=f_phase)
    return shell(body, active="meetings", crumb="<b>Deck repository</b> / All created decks")
