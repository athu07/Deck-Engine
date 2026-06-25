# J2W Deck Engine — Handoff

> Written for a fresh session with zero prior context. Read this top to bottom
> before touching anything. Project folder: `C:\Users\E36250417\Downloads\Deck engie`
> (note the folder name is **"Deck engie"** — a typo for "engine", not a mistake to fix).

---

## 1. What this project is

The **J2W Deck Engine** is a web tool for JoulesToWatts' (J2W) sales/tech team. A
salesperson fills in a client's context (client name, industry, work type(s),
function(s), and optional free-text "more information" — a meeting transcript,
minutes/MOM, or client research). The engine then **assembles a tailored
PowerPoint** by picking the most relevant slides out of a 133-slide master deck,
**auto-writing any slide that's missing** (via AI), letting the user reorder /
add / remove / edit, **previewing** the deck in the browser, and **downloading**
the `.pptx`. It's a Flask app run locally; the user/owner is **Athithia**.

---

## 2. Current status (honest)

### Built and working
- **Stable slide IDs** — every slide carries a `J2W_ID: CSxx` line in its speaker
  notes. Insert-safe stamping.
- **Content library** — `library.json` / `tagged_library.json` (133 records:
  id + text + keywords + auto-tags + pointer).
- **Matcher** (`matcher.py`) — context -> slide IDs using the registry's rules +
  keyword scoring + **AI refinement** + **gap detection** + **suggested extras**.
- **Assembler** (`assembler.py`) — slide IDs -> a new `.pptx` (formatting
  preserved; unused media pruned so files aren't bloated).
- **Full web UI** (`app.py` + `static/app.css`) — a high-end platform, brand
  colours white `#F6F6F6` / black `#111110` / teal `#2C6E66`, dark icon sidebar.
  Pages: **New deck** (landing), **Slide library**, **Templates**, **Review
  queue**, **Dashboard**, plus the build / review / preview flow.
- **Missing-slide generation** (`slide_generator.py` + `templates.pptx`) — AI
  writes the slide content, Python pours it into a template; happens
  automatically when a gap is flagged.
- **Review + preview** — text editing of title/subtitle (`/review`), then an
  in-browser **PPTXjs** preview before download.
- **Self-learning loop** (`staging.py`) — AI-written slides are saved as
  "pending", **reused** on future matching slots, and **promoted into the master
  library only after a human approves** them in the Review queue.
- **Content loaded** — 64 real MS case-study slides were added (CS69–CS132), and
  one AI-generated slide was approved/promoted during testing (CS133).

### Partial / has caveats
- **Templates** — the only template (`case_study`) is a **text-only PLACEHOLDER**.
  The real J2W-designed template still needs to be made and swapped in. Image-rich
  templates need extra work (see gotchas — `_copy_slide` only copies text shapes).
- **Industry/function tags** on the 64 new slides are `AUTO` (machine-guessed),
  refinable. Matching still works because keywords (verbatim) dominate scoring.
- **Preview = PPTXjs** (in-browser, best-effort, **view-only**). A pixel-perfect
  "MS Office quality" preview using PowerPoint COM was built and then **rolled
  back at the owner's request** (see §6 and §7).
- **Generated slides land at the END** of the deck, not slotted into their section.

### Not started
- **Deploy** — getting it onto a URL the team uses without running Python.
- **Team login / auth** — and the requirement that the "more information" content
  (transcripts/MOM/research) must sit **behind team login** (it's sensitive).
- Inline **body-text** editing (only title/subtitle today).
- Cross-industry "relevant anyway" AI suggestions (suggested extras are currently
  same-industry / same-function only).
- Deeper file-size reduction (pruning unused layouts).

---

## 3. Folder structure (where things live)

### Python modules (the engine)
| File | What it does |
|---|---|
| `app.py` | The whole Flask web app — every page, route, and HTML template (inline `render_template_string`). Start here. |
| `static/app.css` | The design system (colours, fonts, sidebar, cards, all components). |
| `stamp_ids.py` | Writes `J2W_ID: CSxx` into each slide's notes. **Insert-safe**: only un-ID'd slides get the next free number; existing IDs are never renumbered. |
| `build_library.py` | Opens a deck -> `library.json` (one record per slide: id, title, subtitle, body, full_text, keywords, source pointer). `read_id()` lives here. |
| `tagger.py` | Rule-based auto-tagging: `kind / work_type / industry / function` + a per-tag confidence (AUTO/HUMAN) + the keyword list. Keyword maps `WORK_TYPE`, `INDUSTRY`, `FUNCTION`. |
| `matcher.py` | **The brain.** `plan(context, use_ai=...)` -> `{picks, gaps, suggestions, suggested, ai_used}`. Reads the registry for selection rules. Holds `EXCLUDE`, `LEADER_IDS`, `PIN_TO_END`. |
| `ai_matcher.py` | OpenAI call (`gpt-4o-mini`) that refines case-study picks from the transcript and chooses relevant OPTIONAL slides. Called by `matcher.plan` when `use_ai=True`. |
| `assembler.py` | `build_deck(slide_ids, out)` -> a new `.pptx` (keeps only chosen slides in order, drops the rest, prunes their media). `SOURCE = WORKING_COPY_Master_Deck.pptx`. |
| `editor.py` | Title/subtitle editable-field extraction, `set_text` (keeps first run's formatting), `apply_edits`, `replace_tokens` (fills `[CLIENT]`). |
| `slide_generator.py` | AI writes slide content (`draft`) + builds a slide from a template (`append_generated`, `_copy_slide`, `_fill`). Manages `templates.pptx` (tagged template slides with `{{TITLE}}/{{KEYWORDS}}/{{BULLETS}}` markers). |
| `staging.py` | The self-learning loop: `find` (reuse), `add` (save pending), `promote` (approve -> into master+library+registry), `discard`. Stores content in `staging/staging.json`. |
| `secrets_loader.py` | Loads `.env` into the environment (`load_env`, `get_key`). |
| `renderer.py` | **UNUSED / rolled back.** PowerPoint-COM slide-to-PNG renderer. Left on disk in case we revisit; **not imported anywhere**. Do not re-enable without reading §7. |
| `ai_fallback.py` | Old experiment (rules-vs-AI tagging). Unused; ignore. |

### Data / deck / config files
| File / folder | What it is |
|---|---|
| `Master_Deck_Case_Study_Portfolio.pptx` | The **ORIGINAL** master deck. **Has no IDs and is stale** (the new 64 slides + CS133 are not in it). Keep as a pristine backup; **do NOT use it as the source.** |
| `WORKING_COPY_Master_Deck.pptx` | **THE LIVING MASTER** (133 slides, IDs stamped in notes). This is what the engine reads and what new/approved slides are added to. `assembler.SOURCE` points here. |
| `J2W_CaseStudy_Portfolio_Metadata.xlsx` | **The registry.** Sheet `Slide Registry` (one row per slide: slide_id, section, kind, std_group, include_rule, work_types, primary_industry, primary_function, keywords, confidence, title) drives selection. Also `READ ME`, `Tag Dictionary`, `Transcript Parsing` sheets. 133 rows. |
| `J2W_CaseStudy_Portfolio_Metadata.BACKUP.xlsx` | Backup of the registry before the 64 slides were added. |
| `library.json` / `tagged_library.json` | 133 records (raw / tagged). The content library. |
| `templates.pptx` | Generation templates. Currently one: `case_study` (text-only placeholder), tagged `J2W_TEMPLATE: case_study` in its notes, with `{{TITLE}} {{KEYWORDS}} {{BULLETS}}` markers. |
| `staging/staging.json` | Pending / approved AI-written slides (the learning queue). |
| `output/*.pptx` | Generated decks (e.g. `Tailored_Deck_<client>.pptx`, `Slide_CSxx.pptx`). |
| `static/renders/` | Leftover slide PNGs from the rolled-back renderer. Harmless; can be deleted. |
| `.env` | Holds `OPENAI_API_KEY`. **See §7 — the key is exposed and should be rotated.** |
| `.gitignore` | Ignores `.env`. |
| `BUILD_LOG.txt` | The running build log (owner's notebook). Architecture backlog + decisions. |
| `j2w_simplified_architecture.svg` | The owner's architecture diagram (5-step flow). |

---

## 4. How the engine works, end to end

1. **New deck form** (`/`, landing). Inputs: client name, industry (select),
   **function(s)** (required, multi-select chips), **work type(s)**
   (Workforce / AI Pods / Managed services chips), and **"Give me more
   information"** (free text — transcript / MOM / research). No AI toggle —
   **AI runs every time.**

2. **Match** (`POST /build` -> `matcher.plan(context, use_ai=True)`):
   - Always include the CORE slides; for each selected work type, include its
     standard block.
   - Score every case study: **+3 per transcript keyword hit** (primary signal),
     **+2** industry match, **+1** industry-in-keywords, **+1** function match.
     Keep the top ~3 per work type.
   - **AI refinement** (`ai_matcher.refine`): given the transcript + candidates,
     pick the genuinely relevant cases (drops keyword false positives, e.g.
     "iOS" meaning a bank app, not the EdTech case) and choose relevant OPTIONAL
     slides. Falls back to keyword picks if AI returns nothing/errors.
   - **Gaps**: if a needed work-type slot has no decent match (best score < 2),
     emit a "needs to be created" gap.
   - **Suggested extras**: next-best related slides (same industry / function /
     matches notes), shown as an add-able strip.
   - **Excluded**: the 7 reference-only dividers; **leaders never auto-picked.**
   - Order: natural deck order, with `CS07` (Next Steps) and `CS08` (Let's win
     together) pinned to the end.

3. **Suggested slides page** (`BUILD_BODY`): draggable cards (reorder, ↑/↓,
   remove), an Add-slide picker, the "You might also include" strip, the gaps
   ("Written for you" — auto-generated on download via hidden `gen` fields), an
   AI badge, and a sticky rail with **Download .pptx** / **Review & edit text**.

4. **Review & edit** (`POST /review` -> `REVIEW_BODY`): per-slide editable
   **title + subtitle** with the body text shown for context, `[CLIENT]`
   pre-filled. (Currently text-only; a visual render was rolled back.)

5. **Assemble + generate** (`/download` or `/finalize`):
   - `assembler.build_deck(ids)` builds the deck in the chosen order.
   - `editor.replace_tokens` fills `[CLIENT]`.
   - `_maybe_generate(path)` handles the gaps: for each, **reuse** a previously
     generated slide for that work-type+industry (`staging.find`) if one exists
     (no AI call), else **draft** a new one (`slide_generator.draft`, AI) and
     **stage it as pending** (`staging.add`); then append to the deck.

6. **Preview** (`PREVIEW_BODY`): in-browser **PPTXjs** render of the built file
   (served from `/output/<name>`), with "Thank you, Athithia — your deck is
   ready" and a **Download .pptx** button. An error banner shows if a slide
   won't render (the file is still complete).

7. **Self-learning** (`/staging`): newly AI-written slides appear in the
   **Review queue** as "pending". **Approve** -> `staging.promote` builds the
   slide into `WORKING_COPY_Master_Deck.pptx` (stamps the next `CSxx`), rebuilds
   `library.json` + `tagged_library.json`, and appends a registry row -> the
   slide becomes a normal, matchable library slide. **Discard** drops it.
   Pending slides **never** affect matching until approved.

Other routes: `/dashboard` (stats), `/library` (browse/filter all slides +
per-slide download via `/slide/<id>/download`), `/templates` (markers + how-to),
`/output/<file>` (serve/download a built deck).

---

## 5. Key decisions (do not undo without good reason)

1. **Content library, NOT a positional index.** Slides are matched by
   content/tags, not by position. Deck order is irrelevant. Every slide is a
   standalone record keyed by a **stable `J2W_ID`** in its notes.
2. **IDs live in slide NOTES** so they travel with the slide. Stamping is
   **insert-safe** (new slides get max+1; existing IDs never change).
3. **The `keywords` column is the matching fuel.** It's the human-authored
   dot-separated tag string copied from each slide (separator is `·`, U+00B7).
   Trust it over inferred industry/function.
4. **The registry drives selection** (`include_rule` / `std_group` / `section`);
   the auto-tagged library provides content + a fallback. Deck and registry align
   by `CSxx` id.
5. **AI-written slides are QUARANTINED** in `staging` until a human approves them
   (prevents the library drifting/bloating by matching its own unreviewed output).
   Reuse-before-regenerate makes it cheaper/better over time.
6. **The 7 dividers are reference-only** and excluded from every deck:
   `EXCLUDE = {CS09, CS14, CS17, CS20, CS21, CS35, CS45}` in `matcher.py`.
7. **Leader/people slides (CS61, CS62) are never auto-picked** — always surfaced
   as a suggestion (`LEADER_IDS` in `matcher.py`).
8. **AI provider is OpenAI** (`gpt-4o-mini`), not Claude — that's the key the
   owner had. Key in `.env` as `OPENAI_API_KEY`.
9. **"Give me more information" (transcript / MOM / client research) is
   sensitive** and, when this is deployed, **must sit behind team login.** No auth
   exists yet (single-user local tool) — this is a deploy-time requirement, not
   yet built.
10. **Preview is PPTXjs (no Docker).** A PowerPoint-COM pixel-perfect renderer was
    built and **deliberately rolled back** by the owner — see §7 before reviving it.
11. **Work on `WORKING_COPY_Master_Deck.pptx`, never the original.** The working
    copy is now the living master (has the IDs + the 64 new slides + CS133).

---

## 6. What's next (in order)

The session ended right after **rolling back** the PowerPoint-COM visual preview
to the PPTXjs preview. There is **no half-finished work in flight** — the app is
in a clean, working state. The next tasks, in priority order:

1. **Real J2W template for generated slides.** Replace the text-only placeholder
   in `templates.pptx`: design a real J2W slide in PowerPoint, add the marker
   tokens `{{TITLE}} {{KEYWORDS}} {{BULLETS}}`, tag its notes `J2W_TEMPLATE:
   case_study`, drop it into `templates.pptx`. **Caveat:** if the template has
   images/logos, `slide_generator._copy_slide` must be upgraded to copy image
   parts (today it only copies text shapes).
2. **Deploy + team login.** Host so the team uses a URL without Python, and put
   the app (especially the "more information" inputs) **behind team login**
   (decision #9). Decide the hosting target; render/preview on the server.
3. **Decide the preview approach for deploy.** PPTXjs (current) is view-only and
   best-effort. The PowerPoint-COM renderer (`renderer.py`) gives pixel-perfect
   output but needs a machine with PowerPoint and careful file-locking (see §7).
   For a server deploy, LibreOffice headless is the usual answer.
4. **Refine tags on the 64 new slides** (industry/function are AUTO) if exact
   matching matters — optional.
5. **Optional polish:** position generated slides into their section (not end);
   inline body-text editing; deeper file-size reduction.

---

## 7. Fragile / non-obvious / do-not-touch

- **The OpenAI key in `.env` is EXPOSED.** It was pasted into chat earlier, so
  treat it as compromised. The owner said keep it for now and decommission/replace
  it once development is done. **AI runs on every `/build`** (cost) — keep that in
  mind. `secrets_loader.py` strips a leading space in the `.env` value.
- **`renderer.py` is rolled back and must NOT be re-imported casually.** Why it
  was pulled: PowerPoint COM (a) could close the user's **own open PowerPoint**
  via `app.Quit()`, and (b) left the output `.pptx` **locked**, causing
  `PermissionError` 500s on the next build. If reviving: never call `app.Quit()`,
  always `pres.Close()` in a `finally`, render with retry-on-`PermissionError`,
  and degrade gracefully. For a server, prefer LibreOffice headless instead.
- **Never re-stamp IDs from the ORIGINAL master.** `Master_Deck_Case_Study_
  Portfolio.pptx` has no IDs; stamping it fresh would assign IDs by position and
  **renumber everything**, breaking the stable-ID contract. Always edit
  `WORKING_COPY_Master_Deck.pptx` (which already has IDs) and run `stamp_ids.py`
  there — it only fills in un-ID'd slides.
- **Adding slides to the master** (the supported flow): add them to
  `WORKING_COPY_Master_Deck.pptx` in PowerPoint (give each a `·`-separated
  keyword line as its subtitle) -> `py stamp_ids.py WORKING_COPY_Master_Deck.pptx`
  -> `py build_library.py` -> `py tagger.py` -> add registry rows (see how the 64
  were added / how `staging.promote` appends a row). Back up the registry first.
- **python-pptx can't truly duplicate a slide.** `slide_generator._copy_slide`
  deep-copies text shapes into a blank-layout slide (fine for text; **images do
  not transfer**). A freshly created notes slide has **no placeholder**, so
  `staging._stamp_id` borrows a notes placeholder from an existing slide before
  writing the ID — keep that pattern if you touch promotion.
- **No auto-reload.** The Flask dev server runs with `debug=False`. After any code
  change you must restart it: it runs on **port 5000**; stop the old process
  (e.g. kill whatever's on port 5000) and run `py app.py` again.
- **`gpt-4o-mini` is mildly non-deterministic** — it occasionally returns no
  cases for a terse transcript; the keyword fallback in `matcher.plan` covers
  this, so don't "fix" the fallback away.
- **The registry and `tagged_library.json` are read fresh from disk on each
  request** — promoting a slide updates them and is picked up without a restart
  (but a code change still needs a restart).
- **Folder name is `Deck engie`** (typo). Paths and the running server assume it.
- Leftover files that are safe to delete: `static/renders/*`, `err.log` (if
  present), `output/*` (regenerated on demand).

---

## 8. How to run it

```
cd "C:\Users\E36250417\Downloads\Deck engie"
py app.py
```
Then open **http://127.0.0.1:5000**. Python is invoked via the **`py`** launcher
(Python 3.12). Key packages already installed: `flask`, `python-pptx`,
`openpyxl`, `openai`, `pywin32` (only used by the rolled-back renderer).

`BUILD_LOG.txt` has the owner's running notes and the architecture backlog — read
it alongside this file for the owner's own framing of each box/step.
