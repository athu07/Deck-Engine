# -*- coding: utf-8 -*-
"""
build_context.py  --  persist a build's rich context so the AI case-study generator
can synthesise from it AFTER /build.

The deep-research brief, stakeholder profile and full transcript are read once at
/build (for matching) and were previously discarded. Here they are saved to a small
JSON keyed by the build_id already generated on that page; /create_ai reloads them so
generation is grounded in the real client context (not a truncated transcript).

File-backed (not an in-memory dict) so it survives across gunicorn's worker
processes. Old entries are pruned so the folder can't grow without bound.
"""

import json
import os
import time

from deckengine import config

_MAX_AGE_SECONDS = 7 * 24 * 3600   # prune build contexts older than 7 days


def _path(build_id):
    safe = "".join(ch for ch in (build_id or "") if ch.isalnum())[:64]
    if not safe:
        return None
    return os.path.join(config.BUILD_CONTEXT_DIR, safe + ".json")


def _prune(now):
    try:
        for fn in os.listdir(config.BUILD_CONTEXT_DIR):
            if not fn.endswith(".json"):
                continue
            p = os.path.join(config.BUILD_CONTEXT_DIR, fn)
            try:
                if now - os.path.getmtime(p) > _MAX_AGE_SECONDS:
                    os.remove(p)
            except OSError:
                pass
    except OSError:
        pass


def save(build_id, ctx):
    """Persist {research, profile, transcript, industry, recipient, functions,
    client_name} for this build. Best-effort — failure never blocks the build."""
    p = _path(build_id)
    if not p:
        return
    try:
        os.makedirs(config.BUILD_CONTEXT_DIR, exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(ctx, f, ensure_ascii=False)
        _prune(time.time())
    except OSError:
        pass


def load(build_id):
    """Return the saved context dict for build_id, or {} if none / unreadable."""
    p = _path(build_id)
    if not p:
        return {}
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}
