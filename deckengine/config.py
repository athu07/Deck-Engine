# -*- coding: utf-8 -*-
"""
config.py  --  the SINGLE source of truth for every file path the engine uses.

Before this module, ~25 filename strings were hardcoded (and duplicated) across
modules and resolved relative to the current working directory, so the app could
only run from the repo root and no data file could move without breaking imports.

Everything here is anchored to PROJECT_ROOT (the directory this file lives in), so
paths are absolute and CWD-independent. The writable runtime dirs (output/,
meetings/, staging/) stay directly under PROJECT_ROOT so they keep matching the
docker-compose volume mounts (./output:/app/output, etc.) unchanged.

To move a data file later, change ONE line here — every consumer imports from here.
"""

import os
from pathlib import Path

# Repo root == two levels up from this file (deckengine/config.py -> repo root).
PROJECT_ROOT = Path(__file__).resolve().parent.parent


def _p(*parts):
    """Absolute path under the project root, as a plain string (open()/pptx-friendly)."""
    return str(PROJECT_ROOT.joinpath(*parts))


# ── runtime READ inputs (live under data/) ────────────────────────────────────
MASTER_DECK             = _p("data", "decks", "WORKING_COPY_Master_Deck.pptx")   # living master deck
LIBRARY_JSON            = _p("data", "library.json")                    # slide records (untagged)
TAGGED_LIBRARY_JSON     = _p("data", "tagged_library.json")             # slide records + tags
CONTENT_STORE_JSON      = _p("data", "case_study_content_store.json")   # 160 case-study records
CASE_EMBEDDINGS_JSON    = _p("data", "case_embeddings.json")            # semantic vectors

# registry / footprint spreadsheets
REGISTRY_XLSX           = _p("data", "registry", "J2W_CaseStudy_Portfolio_Metadata.xlsx")
DELIVERY_FOOTPRINT_XLSX = _p("data", "registry", "J2W_Delivery_Footprint_Organized_Latest.xlsx")

# salesperson-added industries not in the built-in taxonomy (the "Other" field on
# the New-deck form) -- persisted so they show up in the dropdown from then on
CUSTOM_INDUSTRIES_JSON  = _p("data", "custom_industries.json")

# template decks used to render new slides
TEMPLATES_PPTX          = _p("data", "templates", "templates.pptx")
SKILLS_TEMPLATES_PPTX   = _p("data", "templates", "skills_templates.pptx")
CASE_TEMPLATE_PPTX      = _p("data", "templates", "case_study_v2.pptx")

# source spreadsheet for the content-store build SCRIPTS (not read at request time)
CASE_STUDIES_SOURCE_XLSX = _p("data", "registry", "Case_Studies_Master_IDed.xlsx")

# secrets (stays at repo root, git-ignored)
ENV_FILE                = _p(".env")

# ── writable runtime dirs ─────────────────────────────────────────────────────
# Default to PROJECT_ROOT (so they keep matching the docker-compose volume mounts
# ./output:/app/output etc.), but allow an env override — handy for deployment and
# for tests that must not touch the real working tree.
OUTPUT_DIR        = os.environ.get("DECK_OUTPUT_DIR")        or _p("output")
MEETINGS_DIR      = os.environ.get("DECK_MEETINGS_DIR")      or _p("meetings")
STAGING_DIR       = os.environ.get("DECK_STAGING_DIR")       or _p("staging")
STAGING_JSON      = os.path.join(STAGING_DIR, "staging.json")
# per-build context (deep research + profile + full transcript), keyed by build_id,
# so the AI case-study generator can synthesise from it after /build.
BUILD_CONTEXT_DIR = os.environ.get("DECK_BUILD_CONTEXT_DIR") or _p("build_context")
# "learn a template from a deck" -- one slide extracted + its role-map (Templates page).
LEARNED_TEMPLATES_DIR = os.environ.get("DECK_LEARNED_TEMPLATES_DIR") or _p("learned_templates")
# uploaded client logos, one PNG per client (background-removed), stamped onto the
# title slide next to the J2W wordmark at finalize time.
CLIENT_LOGOS_DIR = os.environ.get("DECK_CLIENT_LOGOS_DIR") or _p("client_logos")
# rendered slide-preview images (served from /static/renders/); regenerated on demand.
RENDERS_DIR = os.environ.get("DECK_RENDERS_DIR") or _p("static", "renders")

# ── AI models ─────────────────────────────────────────────────────────────────
# Matching / extraction stay on the cheap model; case-study GENERATION uses a
# stronger model because it is low-volume and quality-critical.
GEN_MODEL = os.environ.get("DECK_GEN_MODEL") or "gpt-4o"
