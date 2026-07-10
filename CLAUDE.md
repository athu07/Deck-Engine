# CLAUDE.md — read this before changing anything, and keep the structure

This file tells any AI assistant (Claude Code, Claude Desktop) how this codebase is
organised and **where new code must go**. The repo was reorganised into a clean Flask
package; please keep it that way. Do not put logic back into one big file, and do not
add loose files at the repo root.

## What this is

The **J2W Deck Engine** — a self-hosted Flask app that assembles a tailored PowerPoint
for JoulesToWatts' sales team. A salesperson enters client context (client, industry,
deck phase, work type(s), function(s), notes) and optional deep-research / stakeholder-
profile PDFs; the engine picks the most relevant slides, fills data-driven slides,
AI-writes any missing slide, and produces a downloadable `.pptx`. Owner is **Athithia**
(a non-developer — explain changes simply, confirm before large or irreversible edits).

`BUILD_LOG.txt` is the owner's running changelog — read it for history and **append a
short entry when you make a meaningful change**.

## Project structure (where everything lives)

```
wsgi.py              Production entrypoint (gunicorn wsgi:app).
app.py               Dev entrypoint / shim (python app.py). Both just build the app.

deckengine/          THE application package. All code lives under here.
  __init__.py          create_app() — the Flask app factory; registers the blueprints.
  config.py            SINGLE source of every file path + AI model ids. Nothing else
                       hardcodes a path or filename.
  constants.py         Form vocabularies, labels, matching knobs (INDUSTRIES, PHASES...).
  web/                 The web layer — ONE blueprint per area:
    decks.py             / , /new, /build, /review, /finalize, /deck  (the core flow)
    builder.py           /builder — the Custom Slide Builder (paste content -> branded
                         slide, with a duplicate check and a real rendered preview)
    dashboard.py library.py staging.py templates.py meetings.py
    output.py            file serving (/output, /slide/<id>/download)
    api.py               /create_ai (the AI case-study JSON endpoint)
    view_helpers.py      shell(), the file-busy page, filename slug, dashboard stats
  services/            Business logic (NO Flask imports here):
    matching/            matcher, relevance, ai_matcher, personas, synonyms, tagger,
                         dedupe ("do we already have a slide for this?")
    content/             case_library, content_store, build_library, editor,
                         paste_parser (one pasted document -> its slides)
    rendering/           assembler, slide_generator, fill_case_study, skills, staging,
                         deck_build, preview (one slide -> PNG, via LibreOffice)
    ingest.py            reads uploaded research/profile files (PDF/text)
    meeting_log.py       one JSON per client+phase on each generate
    build_context.py     saves a build's research+profile+transcript by build_id
    infra.py             loads .env / OpenAI key

templates/           Jinja pages: base is _shell.html, one file per page (build.html...).
static/
  app.css              the design system (one stylesheet).
  js/                  deck-tray.js, build.js, library.js, new-form.js (front-end logic).

data/                Runtime data the app READS:
  *.json               library, tagged_library, case_study_content_store, embeddings
  decks/               WORKING_COPY_Master_Deck.pptx (the living master)
  templates/           templates.pptx, skills_templates.pptx, case_study_v2.pptx
  registry/            the .xlsx registry / footprint / case-source spreadsheets

scripts/             One-off build/maintenance scripts, run by hand from the repo root.
eval/                Match-quality scorecard.
tests/               Smoke + fix tests (pytest). Run these after any change.
archive/             Retired code / backups (kept for reference; don't wire it back in).
output/ meetings/ staging/ build_context/   Runtime OUTPUT (git-ignored; auto-created).
```

## Structure rules — follow these for EVERY change

1. **New page or route** → add it as a Blueprint in `deckengine/web/<area>.py` and
   register it in `deckengine/__init__.py::create_app`. Do not add routes to `app.py`.
2. **New matching / scoring logic** → `deckengine/services/matching/`.
   **New content/data access** → `services/content/`.
   **New slide building / rendering** → `services/rendering/`.
   Keep the web layer thin — it should call services, not contain the logic.
3. **File paths**: never write `open("somefile.json")` or hardcode a path. Add the path
   to `deckengine/config.py` and import it (`from deckengine import config`). This is
   what lets the app run from anywhere and on the server.
4. **HTML** → a Jinja file in `templates/` (extend `_shell.html`), rendered with
   `render_template`. **Never** put big HTML strings back into Python.
5. **JavaScript** → `static/js/`. **CSS** → `static/app.css`. Not inline in templates.
6. **Data files** → `data/` (json at `data/`, decks/templates/registry in their
   subfolders); then update the matching path in `config.py`.
7. **One-off scripts** → `scripts/`. Start each with the repo-root bootstrap the others
   use, and import engine code as `from deckengine.services... import ...`.
8. **Only `app.py` and `wsgi.py` live at the repo root.** No other loose `.py` files.
9. **Run the tests after changes:** `python -m pytest tests/ -q` — they must stay green.

## Running / testing / deploying

- **Local (dev):** `python app.py` → http://127.0.0.1:5000. No auto-reload — restart
  after code changes.
- **Dependencies:** `pip install -r requirements.txt` (Flask, python-pptx, openpyxl,
  openai, pypdf, PyMuPDF, python-docx, gunicorn). A local `.venv/` is git-ignored.
- **Tests:** `python -m pytest tests/ -q`. The smoke test drives the real
  `/build → /review → /finalize` flow and pins matcher output; it is hermetic (no
  network) — keep it that way.
- **Deploy:** Docker runs `gunicorn wsgi:app`. `OPENAI_API_KEY` must be set in the
  server env (it is git-ignored and does NOT travel with the code).

## Core architecture (unchanged concepts)

**Content-library model, not positional.** Slides match by content/tags. Every library
slide carries a stable `J2W_ID: CSxx` line in its speaker notes; `read_id()` (in
`services/content/build_library.py`) reads it. The living master is
`data/decks/WORKING_COPY_Master_Deck.pptx` (`assembler.SOURCE`).

**Request flow (one linear path):** `/` → `/build` (pick slides) → `/review` (edit) →
`/finalize` (assemble + download). Deck-in-progress state lives in the browser
(localStorage "deck tray" `j2w_deck`); the slide order is a comma-joined id list that
can contain `CSxx`, `AIP/WFS/MSS` (content-store cases), `SK:`, `FP:`, `NEW:` items.

**The three ways a slide enters a deck:** (1) picked from the library / content store
(`services/matching/matcher.plan` → `services/rendering/assembler` + `skills.build_into`
render content-store cases from `data/templates/case_study_v2.pptx`); (2) AI-drafted to
fill a gap (`services/rendering/slide_generator.draft_case_study` → staged as `NEW:` →
rendered at finalize); (3) data-driven skills slides (`services/rendering/skills.py`,
Workforce + RFI gated).

## AI (OpenAI) — how matching & generation use it

- Key in `.env` (`OPENAI_API_KEY`), git-ignored, loaded by `services/infra.py`.
- Models are config constants: `config.GEN_MODEL` (case-study generation, `gpt-4o`) and
  the cheaper `gpt-4o-mini` for extraction/ranking in `ai_matcher.py`.
- Every AI call is **fail-safe**: if the API is unavailable, matching degrades to the
  algorithmic path and generation to a placeholder. Keep new AI calls fail-safe too.

### The matching + generation design (do not regress these)

`/build` runs `ai_matcher.extract_brief()` (one call) → a structured brief:
`{needs (with domain/use_case), avoid (mismatch flags), expressed_interest, account}`.
It builds a cheap shortlist per need (`relevance.shortlist_cases`, semantics/domain-led);
**selection is algorithmic** (in `web/decks.py`) — the top shortlist candidate that isn't
mismatch-flagged and clears `CAPABILITY_COVER`. `ai_matcher.explain_picks()` then only
*explains* the already-chosen case, so the model cannot mis-pick. The mismatch flags also
demote wrong-domain cases in `relevance.rank_cases(..., avoid=)`. `build_context.save()` persists the deep
research + profile + full transcript by `build_id` so `/create_ai` can reload them and
`slide_generator.draft_case_study` **synthesises** them into a grounded case study.

If you change matching or generation, keep: domain/mismatch-aware selection, the
research+profile+transcript reaching generation, and the fail-safe fallbacks.

## Gotchas

- **Template slides in the master have NO `J2W_ID`** (only a `J2W_TEMPLATE:` notes tag);
  `assembler.build_deck` drops ID-less slides so they don't leak in unfilled.
- **`slide_generator._copy_slide`** copies shape XML + image parts but not charts/OLE;
  templates meant for copying should be text/auto-shape (or extend the copier).
- **Marker filling keeps only the first run's formatting** per paragraph — keep
  `{{MARKER}}`s in their own run/box.
- **`current_salesperson()`** (in `web/view_helpers.py`) is a placeholder — the seam to
  wire real login at deploy.
- Runtime dirs (`output/ meetings/ staging/ build_context/`) are git-ignored and kept
  via a `.gitkeep`; their contents are regenerated, never source.
