# J2W Pre-sales Engine — Handoff Document
**Owner:** Athithia (non-developer — explain changes simply and confirm before large or irreversible edits)
**Last updated:** 26 June 2026
**Authoritative running log:** `BUILD_LOG.txt` (read it first; this file is for orientation)

> Project folder: `C:\Users\E36250417\Downloads\Deck engie`
> (folder name is "Deck engie" — a typo for "engine"; not a mistake to fix)

---

## 1. What this project is

The **J2W Pre-sales Engine** (previously called "Deck Engine") is a locally-run Flask web application for JoulesToWatts' **pre-sales team**. A pre-sales person fills in a short form — client name, industry, deck phase (Pre-read / First Meeting / Second Meeting / Proposal), work type(s) (Workforce / AI Pods / Managed services), optional function tags, and a free-text "Give me more information" box for meeting notes or transcripts. The engine then selects the most relevant slides from a master PowerPoint, fills data-driven capability and footprint slides from a spreadsheet, lets the user create AI-generated structured case study slides on demand, and produces a downloadable `.pptx` tailored for that specific client meeting. The sales team uses the output; the pre-sales team operates the tool. There is no database and no cloud hosting yet — it runs on a laptop with `py app.py`.

---

## 2. Current Status

### What is fully built and working

| Feature | Status |
|---|---|
| Core form (client, industry, phase, work types, functions, transcript) | Done |
| Phase field (Pre-read / First Meeting / Second Meeting / Proposal) | Done (captured, not yet used for selection rules — see Pending) |
| Slide library (138 slides: CS01–CS136 + 2 template slides) | Done |
| Rule-based + keyword + AI matching (`matcher.py`, `ai_matcher.py`) | Done |
| Drag-to-reorder, add, remove slides on Suggested page | Done |
| "You might also include" suggestions panel | Done |
| Linear flow: New → Suggested → Review & Edit → Done/Download | Done |
| AI gap-fill (drafts missing slides, shown for Accept/Reject on Review page) | Done |
| Accept = promotes to master library permanently; Reject = discards | Done |
| Guided brief box for gap-fill (editable pre-filled text per gap) | Done |
| "Create with AI" panel — structured 4-field form (Topic, Problem, Solution, Results) | Done — redesigned 26 Jun 2026 |
| Loading progress bar while AI generates in "Create with AI" | Done — added 26 Jun 2026 |
| Strict case-study format (6 capabilities / 3 results / self-review) | Done |
| Skills slides — capability + company footprint (Workforce-only) | Done |
| Proficiency doughnut chart on capability slides (native PowerPoint chart) | Done |
| Deck phase dropdown + recipient field on form | Done |
| Automatic meeting log (one JSON per client+phase in `meetings/`) | Done |
| "Deck repository" page (`/meetings`) — search all created decks | Done — renamed 26 Jun 2026 |
| Library browse with filters and "+ Add to deck" from library | Done |
| Deck tray (browser localStorage — resume a deck in progress) | Done |
| AI history queue (`/staging`) — read-only log of all AI slides | Done |
| Atomic save (prevents corrupt .pptx if file is open in PowerPoint) | Done |
| Loading spinner on slow steps (AI calls) | Done |
| Done page — manual download only (no auto-download) | Done — fixed 25 Jun 2026 |
| Responsive layout — content centres at all zoom levels | Done — fixed 25 Jun 2026 |
| Tool renamed: "J2W Deck Engine" → "J2W Pre-sales Engine" | Done — 26 Jun 2026 |
| "Past meetings" → "Deck repository" throughout | Done — 26 Jun 2026 |

### In progress / partially done

| Feature | Status |
|---|---|
| Phase-driven slide selection (Pre-read emphasises X, Proposal emphasises Y) | PARKED — waiting for owner to define per-phase rules |
| Expert routing/notification after an AI slide is created | PARKED — needs team login first |

### Not yet started

| Feature | Notes |
|---|---|
| Deploy + team login | The single biggest remaining task. Blocks almost everything below. |
| Real J2W slide template for AI-generated slides | Current `case_study_full` in `templates.pptx` is text-only placeholder. Replace with real J2W design — same markers, no code change needed. |
| Per-field inline editing of "Create with AI" slides on /review | Currently: edit via Regenerate or remove and redo. Review page only edits title/subtitle for library slides. |
| Position AI-written slides into their exact deck section | Currently they slot before the closer slides (CS07/CS08). |
| Visual pixel-perfect slide preview | PPTXjs removed. Needs LibreOffice at deploy time. |

---

## 3. Folder Structure

```
Deck engie/
|
|-- app.py                          The entire Flask app. All HTML is inline string
|                                   constants (NEW_FORM_BODY, BUILD_BODY, REVIEW_BODY,
|                                   PREVIEW_BODY, MEETINGS_BODY, etc.). No template files.
|                                   Key constants at top: WORK_TYPES, WT_LABELS, PHASES,
|                                   INDUSTRIES, FUNCTIONS.
|-- matcher.py                      Slide selection brain: rules + keyword + AI scoring
|-- ai_matcher.py                   OpenAI call that refines keyword picks
|-- assembler.py                    Builds the final .pptx from a list of slide IDs
|-- slide_generator.py              AI-writes gap slides; fills templates; hosts the
|                                   strict case-study prompt + draft_case_study() +
|                                   fill_case_study()
|-- skills.py                       Workforce-only data-driven slides (capability +
|                                   footprint), proficiency doughnut chart
|-- staging.py                      Queues AI-written slides; Accept promotes to master
|-- meeting_log.py                  Writes/reads JSON meeting records in meetings/
|-- build_library.py                Reads slide notes to build library.json
|-- tagger.py                       Rule-based tags to build tagged_library.json
|-- editor.py                       In-deck text replacement (keeps first-run formatting)
|-- stamp_ids.py                    Stamps J2W_ID into each slide's notes (insert-safe)
|-- secrets_loader.py               Loads .env (OPENAI_API_KEY); git-ignored
|-- .env                            API key only. NEVER commit. ROTATE before deploy.
|
|-- WORKING_COPY_Master_Deck.pptx   THE LIVING MASTER. All reads/writes go here.
|                                   138 slides: CS01-CS136 (ID-stamped) +
|                                   2 template slides (J2W_TEMPLATE: skills and
|                                   J2W_TEMPLATE: company_footprint, NO J2W_ID).
|-- WORKING_COPY_Master_Deck.BEFORE_CHART.pptx   Backup before chart layout edit
|-- Master_Deck_Case_Study_Portfolio.pptx         Original pristine backup. NEVER use
|                                                 as source or re-stamp IDs from it.
|-- templates.pptx                  AI-generated slide templates. Two tagged slides:
|                                   - "case_study" (simple: TITLE/KEYWORDS/BULLETS)
|                                   - "case_study_full" (structured: TITLE/SUBHEAD/
|                                     CHALLENGE/SOLUTION/CAPABILITIES/RESULTS)
|                                   Both are text-only placeholder designs.
|                                   Replace with real J2W designs when ready —
|                                   same tag + markers, no code change.
|
|-- J2W_CaseStudy_Portfolio_Metadata.xlsx  Slide registry (sheet: "Slide Registry").
|                                          One row per slide: slide_id, include_rule,
|                                          std_group, section, work_types, keywords, etc.
|                                          Keywords column separator is · (U+00B7).
|-- J2W_CaseStudy_Portfolio_Metadata.BACKUP.xlsx   Registry backup
|-- J2W_Skills_Inventory.xlsx       Skills data. ONLY two aggregated sheets are read:
|                                   "Skills Master (Aggregated)" and
|                                   "Client Footprint (Aggregated)".
|                                   NEVER read "Consultant Detail" — individual names.
|
|-- library.json                    Slide content library (rebuilt by build_library.py)
|-- tagged_library.json             Same + tags (rebuilt by tagger.py after library.json)
|
|-- static/
|   `-- app.css                     Design system. Brand: white #F6F6F6, black #111110,
|                                   teal #2C6E66. Space Grotesk + Inter fonts.
|                                   .form-layout = 2-column grid (collapses at 980px).
|                                   .page has margin:0 auto so content centres at all widths.
|
|-- staging/
|   `-- staging.json                All AI-generated slide records (pending/approved/
|                                   discarded). Grows on every "Create with AI" or gap-fill.
|-- meetings/                       One JSON per client+phase (J2W_<Client>_<Code>.json).
|                                   Newest overwrites for the same client+phase.
|-- output/                         Generated .pptx decks served for download.
|
|-- BUILD_LOG.txt                   Owner's running notebook. Authoritative decision log.
|                                   UPDATE THIS when making meaningful changes.
|-- CLAUDE.md                       Instructions for Claude Code (this tool).
|-- HANDOFF.md                      This file.
|
|-- ai_fallback.py                  DEAD CODE. Old static re-tagger. Do not revive.
`-- renderer.py                     DEAD CODE. Rolled-back PowerPoint COM renderer.
                                    Left on disk in case we revisit it later.
```

---

## 4. How the Engine Works End to End

### The linear request flow (one path only)

```
GET /  or  /new
  Renders NEW_FORM_BODY: the context form.
  Fields: client name, industry, deck phase (required), work types (required,
  at least one), functions (optional, "Any function" clears selection),
  recipient (optional), "Give me more information" (transcript / MOM / notes).

POST /build
  matcher.plan(context) runs:
    1. Core always-in slides (intro, why J2W)
    2. Per-work-type standard blocks from the registry
    3. Case studies scored by: transcript keyword hits, then industry, then function
    4. ai_matcher.refine() (OpenAI) — filters false positives, adds optional slides
    5. Gap detection: work types with no strong case match flagged "needs to be created"
    6. skills.candidates(context): if pure-Workforce deck, adds capability/footprint slides
  Renders BUILD_BODY: the "Suggested slides" panel.
    - Drag-to-reorder list
    - Gaps shown as amber chips with an editable brief box
    - "You might also include" suggestions
    - "Create with AI" card (always visible — structured form, see below)

POST /review
  For each gap work type: calls slide_generator.draft() -> stages the content.
  Shows per-slide editable cards:
    - Library slides: edit title/subtitle, [CLIENT] token fill
    - AI-written gaps: edit title/keywords/bullets, Accept/Reject choice (default Accept)

POST /finalize
  For each AI card: Reject -> staging.discard(); Accept -> staging.promote()
    (promotes to master deck + library + registry, gets real CS id like CS137)
  Assembles the deck: assembler.build_deck(ids) -> copy of master, keep only chosen slides
  "Create with AI" slides (NEW:<id>) built from templates.pptx case_study_full template
  skills.build_into() fills and slots capability/footprint slides
  editor replaces [CLIENT] tokens with the real client name
  Atomic save to output/<filename>.pptx
  meeting_log.write() saves the session record to meetings/
  Renders PREVIEW_BODY: "Your deck is ready" page.
    Download button is manual — user must click it. No auto-download.

GET /output/<filename>?dl=1
  Serves the .pptx as a file attachment (browser downloads it).
```

### Other routes

| Route | What it does |
|---|---|
| `/meetings` | "Deck repository" — search all previously generated decks by industry / work type / phase |
| `/library` | Browse all 138 slides with filters; "+ Add to deck" adds any slide to current deck |
| `/deck` | Resume an in-progress deck from browser localStorage |
| `/staging` | Read-only AI history — log of all AI-written slides (accepted / rejected / pending) |
| `/create_ai` | POST endpoint (AJAX); called by the "Create with AI" panel |
| `/generate` | POST endpoint (AJAX); called by per-gap "Generate with AI" buttons |
| `/slide/<id>/download` | Download a single slide as a 1-slide .pptx from the library |

### "Create with AI" — how the structured form works

The panel is always visible on the Suggested page (BUILD_BODY). It replaced an earlier free-text textarea with a structured 4-field form:

| Field | Required | Pre-filled |
|---|---|---|
| Topic / Use case | Yes | No |
| Industry | No | Yes — from the deck's industry |
| Problem / Challenge | Yes | No |
| Solution | No | No |
| Results | No | No |

On Generate:
1. JS assembles a single brief string from the fields: `"<topic>. Problem: <problem>. Solution: <solution>. Results: <results>."`
2. A teal progress bar animates below the button (advances to ~88%, then snaps to 100% on completion).
3. POST /create_ai sends the brief + industry to `slide_generator.draft_case_study()` via OpenAI.
4. Returns structured JSON: title / subhead / challenge / solution / capabilities[6] / results[3] / self-review verdict.
5. Shown inline as a preview card — Regenerate / Add to deck / Discard.
6. If added: rides in the deck order as `NEW:<staging_id>`.
7. At /finalize: built from the `case_study_full` template.
8. NOT promoted to master library — "this deck only" (owner's explicit choice).

The human supplies the real facts; the AI just writes them up in J2W's strict format and self-checks.

### Gap-fill vs "Create with AI"

| | Gap-fill | Create with AI |
|---|---|---|
| Triggered by | Missing case study for a work type | Pre-sales wants a custom slide |
| Input | Pre-filled brief (editable one-liner) | 4-field structured form |
| Format | Simple (title/keywords/bullets) | Full case study (challenge/solution/caps/results) |
| After Accept | Promoted to master library forever | This deck only (not in library) |

---

## 5. Skills Slides — Detailed Rules

**Gate:** Skills slides appear ONLY in a **pure Workforce deck** — "Workforce" selected AND no other work type. If Workforce + Managed services is selected, no skills slides.

**Capability slides** (tagged `J2W_TEMPLATE: skills` in the master, around slide position 14):
- **Keyword match** (uncapped): a skill area name from "Skills Master (Aggregated)" appears in the meeting notes/transcript. Every keyword hit gets a slide.
- **Industry match** (capped at 3): the deck's industry appears in that skill area's `industries_served` column. Max 3 slides from this path.
- `STOPWORDS` in `skills.py` prevents generic words ("engineering", "solutions", "management", "platforms", "services", "modernization", "legacy") from matching every capability slide.
- Each capability slide gets a **proficiency doughnut chart** (Expert/Intermediate/Junior counts, brand-coloured teal palette, native PowerPoint chart via python-pptx). The `{{CHART}}` marker in the template is replaced with the real chart at the same position.
- Staleness: `last_verified` > 90 days → yellow warning in the Suggested panel.

**Company footprint slide** (tagged `J2W_TEMPLATE: company_footprint`, around slide position 13):
- Trigger: form client name exactly matches a row in "Client Footprint (Aggregated)".
- One slide per deck maximum. No chart — data is single aggregates, not chart-shaped.

**Both template slides have NO `J2W_ID`** — only a `J2W_TEMPLATE:` tag in their notes. This is intentional. `assembler.build_deck()` drops all slides without a J2W_ID, preventing the templates from leaking unfilled into every deck.

**Skills data column shape** (for when real data replaces mock data):
- "Skills Master (Aggregated)": `skill_area`, `total_consultants`, `expert`, `intermediate`, `junior`, `avg_ramp_up_weeks`, `available_now`, `industries_served`, `example_clients`, `last_verified`
- "Client Footprint (Aggregated)": `client_name`, `total_deployed`, `active_engagements`, `skill_areas_count`, `relationship_duration`, `skills_deployed`, `divisions_served`, `expansion_areas`, `snapshot_date`, `last_verified`
- `industries_served` is plain text. `TECH_IT` in the deck is normalised to `tech` before matching.

---

## 6. Key Decisions and Why

1. **Content-library, not positional.** Slides are matched by content/tags (stable CS-id in notes), never by position. You can rearrange the master deck without breaking anything.

2. **`keywords` column is the primary matching fuel.** The middle-dot-separated keywords column in the registry drives most case-study scoring. Separator is `·` (U+00B7) — editors may show it as a period; the code handles it.

3. **AI at assembly time, not tagging time.** An experiment with static AI re-tagging did not beat the rule-based tagger. AI is useful for reading a specific transcript to score specific cases, not for static metadata.

4. **AI-written slides are reviewed before use.** Gap-fill and "Create with AI" slides go through Accept/Reject on /review. Accepting permanently promotes to master (new CS id, library row, registry row). Rejecting discards. `/staging` is now a read-only history log.

5. **Accept = permanent growth of the master.** Each Accept adds a slide to `WORKING_COPY_Master_Deck.pptx`. Back it up before testing the Accept path. The original (`Master_Deck_Case_Study_Portfolio.pptx`) is a pristine backup — never use it as the source.

6. **Meeting/deck log must be behind team login.** Stores client names and context. Never a public URL.

7. **Skills data is aggregated-only.** `J2W_Skills_Inventory.xlsx` has a "Consultant Detail" sheet with individual names. That sheet is NEVER read.

8. **Function grouping of skills needs team review** before the skills slides go live with real data.

9. **Salesperson is a placeholder.** `current_salesperson()` in `app.py` returns `"[NOT LOGGED IN - wire to login at deploy]"`. Wire the real logged-in user there at deploy time — one place, everything else reads from it.

10. **OpenAI key was exposed in chat once.** Treat as compromised — REVOKE in the OpenAI dashboard and put a fresh one in `.env` before deploy.

11. **Primary users are pre-sales team.** Sales team use the output. Pre-sales operate the tool (they know which case studies are relevant, they interpret transcripts, they approve AI slides). Tool name updated to "J2W Pre-sales Engine" to reflect this.

---

## 7. What's Next (in order)

### Priority 1 — Deploy + team login
Unblocks almost everything else:
- Choose hosting (internal server, cloud VM, Docker — `Dockerfile` + `docker-compose.yml` already in the folder)
- Wire `current_salesperson()` to the real logged-in user
- Put the whole app behind auth — deck repository and transcript input must never be public
- Rotate the OpenAI API key into the new environment's `.env`
- Expert routing: after an AI slide is created, notify the relevant solution lead (needs team identity from login)

### Priority 2 — Phase-driven slide selection
Phase (Pre-read / First Meeting / Second Meeting / Proposal) is captured but not yet used to change which slides are picked. Owner must define per-phase rules (e.g. "Pre-read = intro + overview only; Proposal = include pricing and ROI"). **Do NOT guess sales logic — wait for owner's input.**

### Priority 3 — Replace mock skills data with real organised data
`J2W_Skills_Inventory.xlsx` has placeholder/dummy aggregated data. Real data needs to go into the two aggregated sheets in the same column shape.

### Priority 4 — Real J2W case-study slide template
`case_study_full` in `templates.pptx` is a text-only placeholder. When the design team builds the real PowerPoint template, drop it into `templates.pptx` as a slide tagged `J2W_TEMPLATE: case_study_full` with the same markers. No code change needed.

### Priority 5 (nice to have)
- Per-field inline editing of "Create with AI" slides on /review page
- Visual (pixel-perfect) slide preview — needs LibreOffice at deploy
- Position AI-generated slides into their exact deck section (currently before closers CS07/CS08)

---

## 8. Fragile / Non-obvious Things — Do Not Touch Without Reading This

### {{MARKERS}} must be in single text runs
`editor.set_text()` and `skills.fill_markers()` replace `{{MARKER}}` tokens by rewriting the **first run** of a paragraph and blanking the rest. If a marker is split across two runs in PowerPoint (e.g. `{{TI` then `TLE}}`), it will NOT match. Always keep each marker in one clean text run in the template.

### Template slides have NO J2W_ID — intentional
`J2W_TEMPLATE: skills` and `J2W_TEMPLATE: company_footprint` in the master, and both slides in `templates.pptx`, intentionally have no `J2W_ID`. `assembler.build_deck()` drops ALL slides without a J2W_ID — that is how templates are prevented from leaking unfilled. **Do NOT stamp template slides with `stamp_ids.py`.**

### `stamp_ids.py` is insert-safe — existing IDs never renumber
Only un-stamped slides get a new ID (`max + 1`). Safe to run after adding new slides. **Never run it against `Master_Deck_Case_Study_Portfolio.pptx`** — that would renumber everything and break all cross-references.

### `_copy_slide` does NOT copy image parts
`slide_generator._copy_slide()` deep-copies shape XML but not image part references. Picture shapes in a template become broken-image boxes in the output. This is why ALL picture shapes were removed from both skills template slides. Any new template for copying must be text/auto-shape only.

### lxml elements are "falsy" when childless
An lxml XML element with no children evaluates as `False` in Python. Always use `is None` / `is not None` when testing lxml elements — `if elem:` will wrongly skip a valid but childless element. This caused a reorder bug in `skills.build_into()` (now fixed).

### Deck tray lives in localStorage, not the server
Browser `localStorage` key `j2w_deck` holds the in-progress deck. Clearing browser storage loses the deck. No server-side session exists.

### DIVIDER slides are always excluded
7 divider slides (CS09, CS14, CS17, CS20, CS21, CS35, CS45) are section headers. `matcher.py` always excludes them from every deck.

### Leader slides CS61/CS62 are never auto-picked
Surfaced as suggestions only. Salesperson can add them manually.

### Meeting log overwrites for the same client+phase
Second deck for "Acme Bank / First Meeting" replaces the first JSON. Only the latest is kept. Owner's explicit choice.

### The "Don't Break" list
- `WORKING_COPY_Master_Deck.pptx` — back up before any Accept-path testing
- `staging/staging.json` — append-only history; don't truncate it
- The `·` separator in the keywords column — do not replace with regular periods
- `assembler._atomic_save()` — do not revert to direct `prs.save()` (corrupts files if PowerPoint has them open)
- `PIN_TO_END` slides: CS07 ("Next steps") and CS08 ("Let's win together") always go last; AI slides insert before them
- "Create with AI" JS reads `ca-topic`, `ca-problem`, `ca-solution`, `ca-results`, `ca-industry` — these IDs must match the HTML form field IDs in `BUILD_BODY`

---

## 9. How to Run Locally

```
py app.py
```
Serves at **http://127.0.0.1:5000**. Python 3.12 via the `py` launcher.

**No auto-reload.** After ANY code change: kill the process, run `py app.py` again, then refresh the browser.

Kill the server: press **Ctrl+C** in the terminal, or Task Manager → find Python → End task.

`.env` file must exist in the project folder:
```
OPENAI_API_KEY=sk-...
```

Installed packages (no reinstall needed for local dev): `flask`, `python-pptx`, `openpyxl`, `openai`.

### Quick smoke-test after a change (no browser needed)
```python
from app import app
c = app.test_client()
r = c.post('/build', data={
    'client_name': 'Test', 'industry': 'BANKING',
    'phase': 'First Meeting', 'work_types': 'WORKFORCE', 'functions': 'HR'
})
print(r.status_code)  # should be 200
```
