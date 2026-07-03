# -*- coding: utf-8 -*-
"""
app.py  --  entrypoint shim.

The application now lives in the `deckengine` package (app factory + blueprint +
services). This module keeps the historical `app:app` import working and lets you
run the dev server directly:

    py app.py            # http://127.0.0.1:5000

For production the Dockerfile uses `wsgi:app` (see wsgi.py).
"""

from deckengine import create_app

app = create_app()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
