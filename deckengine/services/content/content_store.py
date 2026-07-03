# -*- coding: utf-8 -*-
"""
content_store.py  --  cached {id -> record} access to the case-study content store.

The store (data/case_study_content_store.json) holds the 160 AIP/WFS/MSS case-study
records. case_library.py serves them in the matcher's row shape; this module gives
the web layer a simple id-keyed lookup for rendering picked cases at finalize.
"""

import json

from deckengine import config

_cache = None


def content_store():
    """{id -> record} for the content-store case studies. Cached; falls back to an
    empty dict if the store file is missing or unreadable."""
    global _cache
    if _cache is None:
        try:
            with open(config.CONTENT_STORE_JSON, encoding="utf-8") as f:
                _cache = {r["id"]: r for r in json.load(f)}
        except (OSError, ValueError):
            _cache = {}
    return _cache
