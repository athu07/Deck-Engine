# -*- coding: utf-8 -*-
"""dashboard.py  --  the /dashboard overview page."""

from flask import Blueprint, render_template

from .view_helpers import shell, dash_stats

bp = Blueprint("dashboard", __name__)


@bp.route("/dashboard")
def dashboard():
    return shell(render_template("dashboard.html", **dash_stats()),
                 active="home", crumb="<b>Dashboard</b> / Overview", tabs=["Overview", "Activity"])
