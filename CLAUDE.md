# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

The **J2W Deck Engine** — a locally-run Flask app that assembles a tailored PowerPoint for JoulesToWatts' sales team. A salesperson enters client context (client, industry, deck phase, work type(s), function(s), free-text "more information"); the engine picks the most relevant slides from a master deck, fills data-driven slides from spreadsheets, AI-writes any missing slide, and produces a downloadable `.pptx`. Owner is **Athithia** (a non-developer — explain changes simply and confirm before large or irreversible edits).

`BUILD_LOG.txt` is the owner's running notebook and the authoritative history of decisions/changes — read it for context and **keep it updated** when you make meaningful changes. `HANDOFF.md` is an older session handoff (still useful for rationale, but some specifics are now superseded by `BUILD_LOG.txt`).

## Running it

```
py app.py        # Python 3.12 via the `py` launcher; serves http://127.0.0.1:5000
```

- **No auto-reload** (`debug=False`). After *any* code change, kill the process on port 5000 and re-run `py app.py`.
- No `requirements.txt`. Dependencies already installed: `flask`, `python-pptx`, `openpyxl`, `openai`, `pywin32` (only the rolled-back `renderer.py` uses pywin32).
- **No test suite.** Verify changes by exercising the real code: either drive the Flask routes with `app.test_client()` (e.g. `c.post('/build', data=...)`) or call the pipeline modules directly (`matcher.plan(...)`, `skills.candidates(...)`, `assembler.build_deck(...)`). Each module also has a `__main__` demo block.
- AI provider is **OpenAI `gpt-4o-mini`**; key lives in `.env` (`OPENAI_API_KEY`, git-ignored, loaded by `secrets_loader`). AI runs on every `/build` (refine) and every gap generation (cost). The key has been exposed in chat — treat as compromised, rotate before deploy.

## Core architecture

**Content-library model, not positional.** Slides are matched by content/tags, never by deck position. Every slide carries a stable `J2W_ID: CSxx` line in its **speaker notes**. `stamp_ids.py` assigns IDs (insert-safe: only un-ID'd slides get `max+1`; existing IDs never renumber). `read_id()` (in `build_library.py`) reads it everywhere.

**The living master is `WORKING_COPY_Master_Deck.pptx`** — this is what the engine reads/writes (`assembler.SOURCE`). The original `Master_Deck_Case_Study_Portfolio.pptx` is a stale pristine backup; never use it as the source or re-stamp IDs from it (that would renumber everything).

**Two sources of truth drive selection:** the registry `J2W_CaseStudy_Portfolio_Metadata.xlsx` (sheet `Slide Registry`, one row per slide: `include_rule` / `std_group` / `section` / `work_types` / `keywords` / etc.) drives *which* slides; `tagged_library.json` provides slide content/keywords. They align by `CSxx`. The `keywords` column (dot-separated, separator is `·` U+00B7 — your console may mangle it) is the primary matching fuel.

**The request flow (one linear path):**
```
/ or /new (NEW_FORM_BODY)
  -> POST /build      matcher.plan() -> picks + gaps + skills; renders BUILD_BODY (the TOC/Suggested panel)
  -> POST /review     writes AI gap slides as editable Accept/Reject cards (REVIEW_BODY)
  -> POST /finalize   accept->promote, assemble, fill skills slides; renders PREVIEW_BODY (auto-downloads)
  -> /output/<file>   serves the .pptx
```
`/download` + `_maybe_generate()` + the red "verification banner" in `slide_generator` are **vestigial** (left over from a pre-linear-flow design; the live flow never hits them). `FORM_HTML`, `ai_fallback.py`, and `renderer.py` (a rolled-back PowerPoint-COM renderer) are also dead — do not revive without reason.

**All HTML is inline** in `app.py` as big `*_BODY` string constants rendered with `render_template_string`. There are no template files. Deck-in-progress state is carried in the browser via a localStorage "deck tray" (`j2w_deck`) plus hidden form fields; the slide order is a comma-joined id list that can contain `CSxx`, `SK:<area>`, and `FP:<client>` ids.

## The three ways a slide enters a deck

1. **Picked from the library** (`matcher.py` → `assembler.py`). `matcher.plan(context, use_ai=True)` returns `{picks, gaps, suggestions, suggested, ai_used}`: always-in core + per-work-type standard blocks + case studies scored by transcript-keyword hits / industry / function, then refined by `ai_matcher.refine()` (which **must** return plain id strings — it normalizes the model's output to avoid `unhashable dict` crashes). `assembler.build_deck(ids)` keeps only chosen slides in order inside a *copy* of the master, **drops every non-kept slide including ID-less ones**, prunes orphaned media, and saves atomically (`_atomic_save` → temp file + `os.replace`, raising a clear "file is open in PowerPoint" `PermissionError` instead of corrupting).

2. **AI-generated to fill a gap** (`slide_generator.py` + `staging.py`). When a work type has no good case match, `/review` calls `slide_generator.draft(gap, {brief, industry, transcript})` — guided by an editable **brief** (pre-filled by `default_brief()`) and grounded in the **format of similar real slides** (`_similar_slides()`). The draft is staged pending (`staging.add`), shown for Accept/Reject. **Accept = full sign-off**: `staging.promote()` builds the slide into the master, assigns a new `CSxx`, rebuilds `library.json`/`tagged_library.json`, and appends a registry row — so each accept **permanently grows the living master** (back it up before testing this path). Reuse-before-regenerate via `staging.find()`. `/staging` is now a read-only history.

3. **Data-driven skills slides** (`skills.py`, Workforce-only). `skills.candidates(context)` is **gated to a pure-Workforce deck** (`work_types == {WORKFORCE}`); otherwise none. Capability slides match keyword-first (uncapped) plus industry matches capped at `CAP_INDUSTRY=3` (a `STOPWORDS` set keeps generic words like "engineering" from over-matching). Footprint slide is added when the form client matches a Client Footprint row. Data comes from `J2W_Skills_Inventory.xlsx` (only the two *aggregated* sheets — **never** "Consultant Detail"); rows older than 90 days (`last_verified`) get a stale flag. `skills.build_into()` copies the master's `J2W_TEMPLATE: skills` / `J2W_TEMPLATE: company_footprint` slides, fills `{{MARKER}}` tokens, adds a native brand-coloured doughnut into a `{{CHART}}` placeholder if present, and reorders the deck to the final id list.

## Non-obvious gotchas

- **Template slides in the master have NO `J2W_ID`** (only a `J2W_TEMPLATE:` notes tag). They are fill-on-demand templates, not library picks — which is why `assembler.build_deck` must drop ID-less slides (else they leak in unfilled).
- **`slide_generator._copy_slide` copies shape XML but NOT image parts** — pictures in a copied slide become broken-image references. Picture shapes were therefore stripped from the skills templates. Any new template meant for copying should be text/auto-shape only (or `_copy_slide` needs image-part copying).
- **Marker filling preserves only the first run's formatting** per paragraph (`editor.set_text`, `skills.fill_markers`). Keep markers in their own run/box where formatting matters.
- The **meeting log** (`meeting_log.py`) auto-writes one JSON per client+phase (`meetings/J2W_<Client>_<PhaseCode>.json`, newest overwrites) on every generate; `/meetings` searches it. Phase codes: PR/FM/SM/PP.
- The salesperson is captured via `current_salesperson()` in `app.py`, currently a placeholder — the single seam to wire to real login at deploy.
