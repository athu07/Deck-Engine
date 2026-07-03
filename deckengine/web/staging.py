# -*- coding: utf-8 -*-
"""staging.py  --  the /staging read-only AI-history page."""

from flask import Blueprint, render_template

from deckengine.services.rendering import staging
from .view_helpers import shell

bp = Blueprint("staging", __name__)


@bp.route("/staging")
def staging_page():
    # read-only history, newest first (records have no timestamp pre-this build -> keep order)
    items = list(reversed(staging.all_items()))
    body = render_template("staging.html", items=items)
    return shell(body, active="staging", crumb="<b>AI history</b>")
