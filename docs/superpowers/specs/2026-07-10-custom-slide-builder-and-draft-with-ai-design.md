# Design — Custom Slide Builder, Library Auto-Save, Duplicate Detection, Draft with AI

Date: 2026-07-10
Status: Approved
Owner: Athithia (non-technical) · relayed by Krishna

## Context

Two features get work. Neither is greenfield — both extend code that already exists.

**The "Already have the content?" card** (`templates/new_form.html`, `POST /from_content`)
already turns pasted content into a branded slide, auto-detecting one of eight shapes,
and queues several slides before folding them into a generated deck. What it lacks: a
dedicated place to work slide-by-slide, a visual preview of what was built, a check
against the library before building something we already have, and a reliable path
into the library for the case studies it produces.

**The suggested-slides page** (`templates/build.html`) shows a "Draft with AI" button on
every detected gap. The button generates nothing — `prefillAI()` only scrolls to the
manual "Create a slide with AI" form and fills two of its fields.

### What already works (do not rebuild)

- `slide_generator.CONTENT_TEMPLATES` — the eight shapes and their builders.
- `slide_generator.classify_content()` — picks a shape from pasted text.
- `case_library.promote_ai_case()` — writes content-store JSON + a row in
  `Case_Studies_Master_IDed.xlsx` + an embedding. Called from `deck_build.assemble()`.
- `relevance.embed_texts()` + `data/case_embeddings.json` — 177 case vectors.
- `reskin.render_pngs()` — pptx → PNG via LibreOffice + `pdftoppm` (both in the Dockerfile).
- `POST /create_ai` — takes a free-text brief, returns a full structured case study
  grounded in the build's research + profile + full transcript.

### Two latent bugs found during design (in scope, because the builder depends on them)

1. `staging.add()` persists shape fields only for `case_study`, `four_box`, and
   `roadmap_board`. A `box_grid` / `pillar_deepdive` / `scored_list` / `stat_overview` /
   `data_table` record loses `eyebrow`, `blocks`, `rows`, `stats`, `items`, `col_labels`
   on save.
2. `deck_build.assemble()` special-cases only `four_box` and `roadmap_board`. The other
   five shapes fall through to `ai_to_store_record()` — they render as a case study, and
   are promoted into the library as one.

Both must be fixed for the builder to preview all eight shapes.

## Scope

Everything in Feature 1 lives in the "Already have the content?" card and the new page
it opens. It does not touch the "build a full deck from context" form, the matcher, or
the deck flow. Feature 2 touches only the suggested-slides page.

Out of scope: the `/review` page still cannot edit five of the eight shapes. Unchanged,
still a known gap.

---

## Feature 1.1 — Custom Slide Builder (`/builder`)

The card on the new-deck form shrinks to an **"Open Custom Slide Builder"** button, a
read-only list of already-queued slides (read from the same `localStorage`
`j2w_content_queue` key), and the existing hidden `content_slide_ids` field. Fold-into-deck
keeps working exactly as it does today.

### Routes — new blueprint `deckengine/web/builder.py`

| Route | Method | Returns | Purpose |
|---|---|---|---|
| `/builder` | GET | page | The builder |
| `/builder/parse` | POST | JSON | Split the document into slides, match shapes, flag duplicates. **Builds nothing.** |
| `/builder/slide` | POST | JSON | Build ONE slide, stage it, return its preview PNG |
| `/builder/preview/<case_id>` | GET | JSON | Render an existing library case to a PNG |
| `/builder/download` | GET | .pptx | Assemble the queue, promote its case studies |

Registered in `deckengine/__init__.py::create_app`, per CLAUDE.md rule 1.

`POST /from_content` stays as-is (unused by the card, kept for the existing
"review just these slides" path and to avoid breaking a bookmarked flow).

### Page behaviour — paste the whole document, review the split, then build

The MS team hands over a deck's worth of content in one file, marked up by hand:

```
slide 1 - case study
…
slide 2 - four box section
…
slide 3 - road map
…
```

**One textarea.** Work type, industry, and client are set once, above it — slides handed
over together come from the same team.

**Step 1 — parse (`/builder/parse`).** `services/content/paste_parser.py` cuts the text at
each header. Two rules, and **no length limits of any kind**:

- A header is `slide <n>` followed by end-of-line or a separator (`-`, `:`, `.`). The
  separator is the guard: `Slide 2 of the deck covers…` has none, so it is never a
  candidate.
- **The slide numbers validate each other.** Real headers form an increasing chain, so
  the parser keeps the longest increasing subsequence over the numbers. A stray line
  that reads like a header (`Slide 1: as discussed above.`) cannot extend the chain and
  stays as content.

  *Two earlier attempts got this wrong.* A 60-character cap on the header line silently
  swallowed two real slides out of nine in the owner's first document — `Slide 4: Case
  Study 1, Reconciliation automation with agentic AI` is 64 characters. Replacing it
  with a 14-word cap was the same mistake in new clothes. Length was never the signal.

A label either **names a category exactly** (`case study`, `road map`) or it is the
slide's **heading**. There is no fuzzy matching — it once matched the heading
`How we would engage with Voya` to the alias `named list with stats` on the word
`with`, and a preposition picked the template. A heading is kept and **prepended to the
slide's content**, so the generator titles the slide what the author called it.

**The intelligence layer.** Every slide whose category the author did not name goes to
`slide_generator.classify_content_many()`, which reads the real content and maps it onto
the categories we have. **One AI call for the whole document**, with the heading passed
as a hint, never as the answer. The slides are classified together on purpose — a deck's
slides inform each other. No slide reaches the review screen as `auto`.

This cannot reuse `classify_content()`: that function is instructed to answer `(A)` —
`case_study` — whenever unsure, which is right for one unlabelled paste and wrong for a
document. It made all nine slides of the owner's test deck come back case studies. The
batch prompt instead requires all three of *a specific client's situation*, *what was
delivered*, and *the outcome* before `(A)` is allowed.

Duplicates are checked here too, before anything is generated, via
`dedupe.similar_cases_many()` — **one** embedding call for the whole document, not one
per slide. Matches are surfaced only for slides that resolved to `case_study`: the
library holds nothing else, so a roadmap cannot duplicate one.

**Step 2 — the review screen.** One row per slide: its heading, a shape dropdown saying
either *you named this shape* or *we read this from the content* (correct anything we got
wrong), the opening of that slide's content (check the split landed right), and any
duplicate with *Preview it* / *Use `<id>` instead*. Reusing a duplicate means that slide
is never built at all. "Fix the split" returns to the textarea. **Nothing has been spent
at this point** — this screen exists so a bad split, or a misread shape, is caught before
a 40-second build, not after.

**Step 3 — build (`/builder/slide`, once per slide).** The browser fires them four at a
time (`CB_CONCURRENCY`) and fills each placeholder card as its render lands, so the
slides appear in order as they finish rather than all at once behind a spinner. Ten
slides ≈ 40 seconds.

**The result is a scrollable column** of every rendered slide, numbered, each removable —
and each **editable in place**.

### The inline editor

A rendered slide you cannot correct is only half a tool. Every result card has an **Edit**
button.

`services/rendering/slide_schema.py` declares, per `content_type`, which fields each of
the eight shapes has — the same keys `skills._mapping_*` / `_draw_*` and
`fill_case_study.build_mapping` read at render time. One generic editor in `builder.js`
draws all eight from that schema; there is no hand-written form per shape. Field types:
`text`, `textarea`, `strings` (a list), and `objects` (a list of records, which may
themselves hold a list — a roadmap column's items, a pillar block's sub-points).

- `GET /builder/slide/<id>/fields` → the schema and the current values.
- `POST /builder/slide/<id>` → save, then **re-render** the PNG (`force=True`, the cached
  one is stale). What you see afterwards is the real `.pptx`, not a form claiming it
  worked.

Every save runs through the shape's own `_normalize_*` function
(`slide_schema.apply_edits`), so a hand-edit cannot break an invariant the renderer
depends on: empty a case study's capabilities and six come back; blank its title and it
gets one. The editor may only write keys the schema declares — a browser cannot inject a
field, change a slide's `content_type`, or overwrite bookkeeping (`id`, `work_type`,
`status`, `promoted_id`).

The staged record is the single source of truth, so an edit made here follows the slide
into the download, into a generated deck, and into the shared library, without being
threaded through any of those paths. `staging.update_fields()` takes the same lock as
`add()` — same file, same read-modify-write, and slides are built concurrently.

Queue state lives in `localStorage` under the existing `j2w_content_queue` key, so the
new-deck form picks it up unchanged. It survives a reload (that is what makes "Use in a
deck" work), so the sidebar states the queue size and offers to clear it — otherwise a
queue left from an earlier visit would ride silently into the next download.

### Concurrency

`/builder/slide` is called four at a time, so anything it touches must be thread-safe.
`staging.add()` is a read-modify-write on one JSON file whose id was
`"G%03d" % (len(items) + 1)`. Measured, unfixed, with 24 concurrent adds: **4 unique ids
and 1 surviving record.** It is now serialised behind a lock and mints the lowest unused
`G<nnn>` — which also fixes a latent bug where discarding a record freed a number the
next add re-used, silently aliasing two records.

Two exits:
- **Download these slides** → `GET /builder/download` → one `.pptx` of just the queued
  slides, in queue order. No title/closing bookends.
- **Use in a deck** → back to `/new`, queue intact, folds into the generated deck.

### Preview rendering — new `deckengine/services/rendering/preview.py`

```python
slide_png(record, cache_key) -> str | None    # "/static/renders/builder/<key>.png" or None
```

Builds the single slide to a temp one-slide `.pptx`, then renders it with
`reskin.render_pngs()`. Two paths:

- `content_type == "case_study"` → `fill_case_study.fill_row(record, tmp)`.
- anything else → an empty `Presentation` sized from `config.SKILLS_TEMPLATES_PPTX`,
  then `skills.build_into(tmp, [id], [item])` with the item from `deck_build.staged_item()`.

PNGs cache at `static/renders/builder/<staging_id>.png`; a queued slide renders once.

**Fail-safe:** any exception (LibreOffice missing, render timeout) returns `None`. The
page then shows the structured text card it would otherwise show under the image. A
preview failure never blocks building or downloading a slide.

### The shape-handling fix (bugs 1 and 2 above)

- `staging.add()` — persist every shape field: add `eyebrow`, `blocks`, `rows`, `stats`,
  `items`, `col_labels` to the copied-key list.
- `deck_build.py` — new `staged_item(rec, industry)` returns the `skills.build_into()`
  item for any staged record: a `case_study_v2` item carrying a store record for case
  studies, else `{"id", "template": rec["template"], "kind": rec["content_type"],
  "data": rec}`. `assemble()` and `preview.py` both call it. The `four_box` /
  `roadmap_board` branches collapse into it.

---

## Feature 1.2 — Auto-save to the Slide Library

`case_library.promote_ai_case()` is unchanged. It gains one new call site.

**Save on commit, not on build.** A case study is written to the library when you
either download the queue or fold it into a deck — never when it is merely built and
possibly discarded.

- **Download** → `/builder/download` calls `promote_ai_case()` for each queued record
  whose `content_type == "case_study"`.
- **Fold into deck** → `/finalize` already promotes it via `deck_build.assemble()`.

**Idempotency.** `promote_ai_case()` returns the new id; the caller stamps it onto the
staging record as `promoted_id` (new `staging.mark_promoted(stg_id, case_id)`). Both call
sites skip any record that already carries one, so downloading and then building a deck
with the same slide saves it once.

The work-type prefix comes from the record's own `work_type` (required on the builder
form), not the deck's first work type. This fixes the existing best-effort guess for
builder-created slides.

**Only case studies save.** The other seven shapes have no Excel schema — matching what
`deck_build.assemble()` already does for `four_box` and `roadmap_board`.

---

## Feature 1.3 — Duplicate / similar content detection

New service `deckengine/services/matching/dedupe.py`:

```python
similar_cases(text, top_n=3, threshold=0.80) -> [{"id", "title", "work_type", "score"}]
```

Embeds `text` with `relevance.embed_texts()`, cosines against the vectors already in
`data/case_embeddings.json`, returns matches at or above `threshold`, best first.
One embedding call. No new model, no new index.

**Fail-safe:** no API key, no embeddings file, or any exception → `[]`. The build proceeds.

### Behaviour

Runs on paste, **before** the slide is generated (`POST /builder/check`).

- **Below 80%** — nothing shown, build proceeds silently.
- **At or above 80%** — matches render below the paste box: ID, title, similarity
  percentage, a **Preview** button that renders that existing library slide into the
  same preview area (`GET /builder/preview/<case_id>`), and a **Use this one** button
  that drops the existing `AIP`/`WFS`/`MSS` id straight into the queue — no generation,
  no new library record. A **Build a new slide anyway** button proceeds as normal.

Never blocks. The user can always build new.

`THRESHOLD = 0.80` is a module constant in `dedupe.py`, alongside the other tuning
knobs the codebase keeps in `relevance.py` and `constants.py`.

---

## Feature 2 — Draft with AI

### Template (`templates/build.html`)

Inside `#ai-create`, wrap the manual input fields (`ca-topic`, `ca-industry`,
`ca-problem`, `ca-solution`, `ca-results`) and the `ca-genbtn` Generate button in a Jinja
comment `{# ... #}` with a one-line restore note. Keep the card shell, the loader bar
(`ca-loader` / `ca-bar`), and the preview area (`ca-preview`) — they become Draft with
AI's output surface.

Render the card only when `missing` is non-empty. With no gaps, there is nothing to
draft and no manual form, so the card has no purpose.

Retitle the card to "Draft the missing slide" and repoint the gap-section copy.

### Front-end (`static/js/build.js`)

`prefillAI(name, desc, domain, useCase)` → `draftAI(name, desc, domain, useCase, btn)`:

1. If an un-added preview is showing, `confirm()` before replacing it.
2. Assemble the same brief string `createAI()` builds: `"<name>. Problem: <desc>"`,
   plus `Domain:` / `Use case:` when present.
3. `POST /create_ai` with `brief`, `build_id`, `industry`, `client_name`, `recipient` —
   identical to `createAI()`, so the draft is grounded in the real research, profile, and
   full transcript.
4. Render through the existing `caRender()`. Add-to-deck, regenerate, and discard work
   unchanged.
5. On add, mark the originating gap card with a drafted tick.

`caRegen()` re-drafts the last gap rather than reading the (now hidden) form.

`createAI()` stays in the file, unused, ready to restore.

**No route changes.** `/create_ai` already does exactly this.

### One preview at a time

Drafting a second gap while the first sits unadded prompts before replacing it. This
keeps the existing single-global-id preview machinery intact.

---

## Testing

`python -m pytest tests/ -q` must stay green (16 tests), including the golden matcher
pins — nothing here touches matching.

New tests in `tests/test_fixes.py`:

- `dedupe.similar_cases()` returns `[]` with no API key (fail-safe), and ranks an exact
  copy of a stored case above `THRESHOLD` when embeddings are stubbed.
- `staging.add()` round-trips every field of a `box_grid` and a `stat_overview` record.
- `deck_build.staged_item()` maps each of the eight `content_type`s to the right
  `template` / `kind`, and only `case_study` yields a store record.
- `/builder` renders; `/builder/check` returns `{"ok": true, "matches": []}` offline.
- `promote_ai_case()` is called once, not twice, when a record is downloaded and then
  finalized (idempotency via `promoted_id`).

`preview.slide_png()` is not unit-tested against LibreOffice (slow, environment-dependent);
its fail-safe `None` path is.

Manual verification by the owner after implementation.

## Error handling summary

| Failure | Behaviour |
|---|---|
| No `OPENAI_API_KEY` | Duplicate check returns `[]`, build proceeds. Auto-detect falls back to `case_study`. |
| LibreOffice missing / render fails | `slide_png()` returns `None`; the text card shows instead. |
| Excel file locked | `promote_ai_case()` already fails safe; JSON + embedding still save. |
| Empty queue on download | 400 with a friendly message. |
| Non-pptx / unreadable upload | Rejected up front, as `/from_content` does today. |
