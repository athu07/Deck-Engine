# J2W Pre-sales Deck Engine

A locally-run / self-hosted Flask app that assembles a tailored PowerPoint for the
JoulesToWatts sales team. A salesperson enters client context (client, industry,
deck phase, work type(s), function(s), notes, optional research/profile PDFs); the
engine picks the most relevant slides, fills data-driven slides, AI-writes any
missing slide, and produces a downloadable `.pptx`.

## Project layout

```
config.py            Single source of truth for every file path (PROJECT_ROOT-anchored).
wsgi.py / app.py     Entrypoints. Production: `gunicorn wsgi:app`. Dev: `python app.py`.

deckengine/          The web application (Flask app factory + blueprint).
  __init__.py          create_app() — builds the Flask app, registers the blueprint.
  views.py             All HTTP routes (the /build -> /review -> /finalize flow).
  services.py          Business helpers (catalog, dashboard stats, content store, ...).
  constants.py         Form vocabularies, labels, matching knobs.

templates/           Jinja page templates (base _shell.html + one per page).
static/              app.css (design system) + rendered previews.

# the matching + rendering "engine" library (imported by deckengine)
matcher.py relevance.py ai_matcher.py personas.py synonyms.py tagger.py
case_library.py build_library.py editor.py assembler.py slide_generator.py
fill_case_study.py skills.py staging.py research.py meeting_log.py secrets_loader.py

data/                Runtime data inputs (master deck, JSON stores, embeddings,
                     template decks, registry/footprint spreadsheets).
scripts/             One-off build / maintenance scripts (run manually).
eval/                Match-quality scorecard (eval_ericsson.py).
tests/               Smoke tests (pytest).

output/ meetings/ staging/   Runtime-generated (git-ignored; kept via .gitkeep).
archive/             Retired code, backups, superseded assets (kept for reference).
```

## Run it locally

```bash
python -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
python app.py            # http://127.0.0.1:5000
```

Requires `OPENAI_API_KEY` in a `.env` file at the repo root (git-ignored). Without a
valid key the app still runs, but AI features (needs extraction, semantic matching,
AI drafting) degrade to their offline fallbacks.

`debug=False` and there is no auto-reload — after a code change, restart the process.

## Deploy (Docker / Render)

```bash
docker compose up --build          # serves on :5000
```

The image runs `gunicorn wsgi:app`. Set `OPENAI_API_KEY` in the environment (Render:
as an env var; compose reads it from `.env`). `output/`, `meetings/`, and `staging/`
are volume-mounted so generated decks and history persist across restarts.

Optional env overrides (default to the repo root): `DECK_OUTPUT_DIR`,
`DECK_MEETINGS_DIR`, `DECK_STAGING_DIR`.

## Tests

```bash
python -m pytest tests/ -q
```

The smoke test freezes matcher output for fixed contexts (behaviour guard), runs the
full `/build -> /review -> /finalize` flow end-to-end, and checks every page renders.
It is hermetic: no network (AI calls fall back), and it writes only to temp dirs.

## Refreshing the case-study content

Case studies live in `data/case_study_content_store.json`, rebuilt from a master
spreadsheet. To add cases and regenerate everything, run from the repo root:

```bash
python scripts/import_docx_cases.py        # append Word-format cases -> master xlsx
python scripts/build_case_study_store.py   # master xlsx -> content store JSON
python scripts/build_case_embeddings.py    # content store -> semantic embeddings
```

Then restart the app (module caches load the data once at startup).
