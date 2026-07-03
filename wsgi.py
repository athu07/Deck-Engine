# -*- coding: utf-8 -*-
"""
wsgi.py  --  production entrypoint for gunicorn: `gunicorn wsgi:app`.
"""

from deckengine import create_app

app = create_app()
