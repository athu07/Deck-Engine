# -*- coding: utf-8 -*-
"""
deckengine  --  the J2W Pre-sales Deck Engine web application (app factory).

Usage:
    from deckengine import create_app
    app = create_app()

Templates and static assets live at the repo root (templates/, static/); the
factory points Flask at them explicitly so the package can live in a subfolder.
"""

from flask import Flask

from deckengine import config


def create_app():
    app = Flask(
        "deckengine",
        template_folder=str(config.PROJECT_ROOT / "templates"),
        static_folder=str(config.PROJECT_ROOT / "static"),
    )
    # one blueprint per area (deckengine/web/*.py)
    from .web import decks, dashboard, library, staging, templates, meetings, output, api
    for module in (decks, dashboard, library, staging, templates, meetings, output, api):
        app.register_blueprint(module.bp)
    return app
