# -*- coding: utf-8 -*-
"""
industries.py  --  salesperson-added industries (the "Other" field on the
New-deck form) that aren't in the built-in taxonomy (constants.INDUSTRIES).

Persisted to a flat JSON list so a custom industry typed once shows up in the
dropdown for every build after that, for every salesperson (shared, not per-
browser). NOT the matching taxonomy: constants.INDUSTRIES stays the source of
truth for industry-boost scoring; a custom industry here is a display label
only (no code, no industry-boost weight, no skills-slide Excel mapping) -- see
constants.all_industries() for how the two lists combine for the dropdown.
"""

import json
import os

from deckengine import config
from deckengine.services import jsonstore

_PATH = config.CUSTOM_INDUSTRIES_JSON


def load():
    """Every custom industry added so far, in the order first added."""
    try:
        with open(_PATH, encoding="utf-8") as f:
            data = json.load(f)
        return [s for s in data if isinstance(s, str) and s.strip()]
    except (OSError, ValueError):
        return []


def add(name):
    """Persist a new custom industry if it's non-empty and not already known
    (case-insensitively, against both the custom list and the caller's known
    built-in list). Returns True if it was newly added."""
    name = (name or "").strip()
    if not name:
        return False
    items = load()
    if name.lower() in {i.lower() for i in items}:
        return False
    items.append(name)
    os.makedirs(os.path.dirname(_PATH), exist_ok=True)
    jsonstore.write_json(_PATH, items)
    return True
