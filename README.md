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

The image runs `gunicorn wsgi:app`, and everything the app *writes* — the deck
repository (`meetings/`), the generated `.pptx` files (`output/`), the custom-slide
`staging/`, per-build context, learned templates, client logos, and salesperson-added
industries — is stored on disk. **That disk MUST be persistent, or all of it is lost
on the next restart.**

### Local (Docker)

```bash
docker compose up --build          # serves on :5000
```

`docker-compose.yml` mounts every writable folder to the host, so the data persists
across restarts. `OPENAI_API_KEY` is read from `.env`.

### Render — the important part

A Render web service has an **ephemeral filesystem**: without a persistent disk,
generated decks and history are wiped on every deploy and on Render's routine
restarts, so the Deck repository always looks empty. Fix it with a persistent disk
and point the app's writable folders at it. `render.yaml` in this repo does exactly
that (Docker web service + a 1 GB disk mounted at `/var/data` + the `DECK_*` env vars).

- **New service:** create a Blueprint from `render.yaml`; Render prompts once for
  `OPENAI_API_KEY`.
- **Existing service:** in the dashboard, add a **Disk** (mount path `/var/data`,
  1 GB), then set these environment variables so the app writes onto it:
  `DECK_OUTPUT_DIR=/var/data/output`, `DECK_MEETINGS_DIR=/var/data/meetings`,
  `DECK_STAGING_DIR=/var/data/staging`, `DECK_BUILD_CONTEXT_DIR=/var/data/build_context`,
  `DECK_LEARNED_TEMPLATES_DIR=/var/data/learned_templates`,
  `DECK_CLIENT_LOGOS_DIR=/var/data/client_logos`,
  `DECK_CUSTOM_INDUSTRIES_JSON=/var/data/custom_industries.json`.

A disk needs a **paid** instance (`starter`+), and a disk-attached service can't be
horizontally scaled — both fine for a single sales team. Every path above defaults to
the repo root when its env var is unset, so local runs need no configuration.

**Slide previews on the review page.** The review page shows a rendered picture of
each master-deck slide. Those images are pre-rendered and **committed** under
`static/previews/`, so a host without LibreOffice — a plain Render *web service*, not a
Docker service — still serves them. **When you change the master deck, re-run
`python scripts/prerender_master.py` and commit `static/previews/`.** The Docker build
also runs it as a safety net. (Needs `libreoffice-impress` + `poppler-utils`, which the
Dockerfile installs; the render itself is only ever done at build time, never per request.)

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
