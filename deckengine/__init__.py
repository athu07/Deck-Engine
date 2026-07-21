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


def _asset(path):
    """A /static URL stamped with the file's modification time, e.g.
    /static/js/builder.js?v=1783674521.

    Without this, a browser holds the previous app.css / builder.js after a deploy and
    the user sees the OLD page against the NEW server -- which looks exactly like the
    change never shipped. There is no auto-reload here and no build step, so the mtime
    is the version.
    """
    full = config.PROJECT_ROOT / "static" / path
    try:
        return "/static/%s?v=%d" % (path, full.stat().st_mtime)
    except OSError:
        return "/static/%s" % path


def create_app():
    # Make sure every writable runtime folder exists before the first request. On a
    # host with a freshly-mounted persistent disk (Render) the disk starts empty, so
    # meetings/, output/, staging/ etc. would not exist until something happened to
    # write into them -- this creates them once, up front. Idempotent everywhere else.
    config.ensure_runtime_dirs()
    app = Flask(
        "deckengine",
        template_folder=str(config.PROJECT_ROOT / "templates"),
        static_folder=str(config.PROJECT_ROOT / "static"),
    )
    app.jinja_env.globals["asset"] = _asset
    # one blueprint per area (deckengine/web/*.py)
    from .web import (decks, dashboard, library, staging, templates, meetings, output,
                      api, builder)
    for module in (decks, dashboard, library, staging, templates, meetings, output,
                   api, builder):
        app.register_blueprint(module.bp)
    return app
