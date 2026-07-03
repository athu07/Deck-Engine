# -*- coding: utf-8 -*-
"""templates.py  --  the /templates page (fill-on-demand slide templates)."""

import re

from flask import Blueprint, render_template

from deckengine.services.rendering import slide_generator
from .view_helpers import shell

bp = Blueprint("templates", __name__)


@bp.route("/templates")
def templates_page():
    items = []
    try:
        for name, slide in slide_generator.list_templates().items():
            markers = set()
            text = ""
            for sh in slide.shapes:
                if sh.has_text_frame:
                    markers.update(re.findall(r"\{\{[A-Z]+\}\}", sh.text_frame.text))
                    text += sh.text_frame.text
            status = "placeholder" if "Generated slide" in text else "active"
            items.append({"name": name, "markers": sorted(markers), "status": status})
    except Exception:
        items = []
    body = render_template("templates_page.html", items=items)
    return shell(body, active="templates", crumb="<b>Templates</b>")
