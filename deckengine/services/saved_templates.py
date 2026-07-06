# -*- coding: utf-8 -*-
"""
saved_templates.py  --  a small store of user-uploaded decks saved as reusable
templates (surfaced on the Templates page: download / delete).

Each saved template is a .pptx on disk plus one row in a JSON index. File-backed so
it survives restarts and is shared across gunicorn workers.
"""

import json
import os
import re
import time

from deckengine import config


def _index_path():
    return os.path.join(config.SAVED_TEMPLATES_DIR, "_index.json")


def _load():
    try:
        with open(_index_path(), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return []


def _write(items):
    os.makedirs(config.SAVED_TEMPLATES_DIR, exist_ok=True)
    with open(_index_path(), "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)


def _slug(name):
    s = re.sub(r"[^A-Za-z0-9]+", "-", (name or "").lower()).strip("-")
    return s or "template"


def save(name, data, slide_count=0):
    """Persist an uploaded deck (bytes `data`) under a display `name`. Returns the row."""
    os.makedirs(config.SAVED_TEMPLATES_DIR, exist_ok=True)
    tid = _slug(name) + "-" + str(int(time.time()))
    fname = tid + ".pptx"
    with open(os.path.join(config.SAVED_TEMPLATES_DIR, fname), "wb") as f:
        f.write(data)
    row = {"id": tid, "name": name.strip() or "Untitled template", "file": fname,
           "slides": int(slide_count or 0), "size_mb": round(len(data) / 1048576, 1)}
    items = _load()
    items.insert(0, row)          # newest first
    _write(items)
    return row


def all_templates():
    return _load()


def get(tid):
    return next((r for r in _load() if r["id"] == tid), None)


def file_path(tid):
    row = get(tid)
    if not row:
        return None
    p = os.path.join(config.SAVED_TEMPLATES_DIR, row["file"])
    return p if os.path.exists(p) else None


def delete(tid):
    items = _load()
    row = next((r for r in items if r["id"] == tid), None)
    if not row:
        return False
    try:
        os.remove(os.path.join(config.SAVED_TEMPLATES_DIR, row["file"]))
    except OSError:
        pass
    _write([r for r in items if r["id"] != tid])
    return True
