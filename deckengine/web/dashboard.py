# -*- coding: utf-8 -*-
"""dashboard.py  --  the /dashboard overview page."""

from flask import Blueprint, render_template

from .view_helpers import shell, dash_stats

bp = Blueprint("dashboard", __name__)


@bp.route("/dashboard")
def dashboard():
    # No tabs: "Overview / Activity" looked like a control and switched nothing. The shell
    # still supports `tabs` for a page that earns them.
    return shell(render_template("dashboard.html", **dash_stats()),
                 active="home", crumb="<b>Dashboard</b>")
