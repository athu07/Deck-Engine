# J2W Pre-sales Deck Engine — Handoff Document

**Owner:** Athithia (non-developer — explain changes simply, confirm before large or irreversible edits).
**Last updated:** 2 July 2026.
**Authoritative running log:** `BUILD_LOG.txt` — the owner's chronological notebook of every decision/change. Read it after this file. This document is for orientation; `BUILD_LOG.txt` is the source of truth for *why* things are the way they are.

> Written so someone who has never seen this project can pick it up with zero loss. If something here disagrees with the code, the code wins — but tell the owner.

---

## 1. What this project is

The **J2W Deck Engine** is a locally-run **Flask** web app that assembles a **tailored PowerPoint** for JoulesToWatts' (J2W) sales team. A salesperson enters client context (client, industry, deck phase, work type(s), function(s), free-text notes) — and optionally uploads a **deep-research brief** and the **stakeholder's profile** (LinkedIn/bio). The engine picks the most relevant case studies and standard slides, fills data-driven slides from spreadsheets, AI-writes any slide that's missing, and produces a downloadable `.pptx`. Users are J2W salespeople prepping for a specific meeting with a specific person; the goal is a deck that makes the buyer feel "this vendor understands my job and my company."

Run it:
```
py app.py        # Python 3.12 via the `py` launcher; serves http://127.0.0.1:5000
```
- **No auto-reload** (`debug=False`). After *any* code change, kill the process on port 5000 and re-run `py app.py`.
- No `requirements.txt`. Installed: `flask`, `python-pptx`, `openpyxl`, `openai`, `pywin32`, `pypdf`, `PyMuPDF (fitz)`.
- **No test suite.** Verify by driving routes with `app.test_client()` or calling modules directly (`matcher.plan(...)`, `relevance.rank_cases(...)`, `skills.build_into(...)`), or run the scorecard `py eval_ericsson.py`.
- **AI provider: OpenAI** — `gpt-4o-mini` (extraction, "why it fits", draft slides) + `text-embedding-3-small` (semantic matching). Key in `.env` (`OPENAI_API_KEY`, git-ignored, loaded by `secrets_loader`). **The key was exposed in chat — treat as compromised, rotate before any deploy.**

---

## 2. Current status

### The 5-step request flow (one linear path, all live)
```
/ or /new  (NEW_FORM_BODY)                 (1) enter client context
   -> POST /build   matcher.plan()          (2) build — engine plans the deck
   -> POST /review                          (3) review AI gap slides (accept/reject)
   -> POST /finalize                        (4) finalize & assemble the .pptx
   -> /output/<file>                        (5) download .pptx (auto-downloads)
```
All five steps work. Each step's HTML is an inline string constant in `app.py` (`*_BODY`) rendered with `render_template_string` — there are **no template files**.

### The three ways a slide is produced (all working)
1. **Picked from the library** — case studies come from the **content store** (`case_study_content_store.json`, 160 cases, ids `AIP/WFS/MSS`), ranked by `relevance.py` and built on demand from the shared branded template `case_study_v2.pptx`. Core/standard/structural slides come from the master deck (`assembler.py`, `CSxx` ids).
2. **AI-generated for a gap** — when a needed capability has no matching case, `slide_generator.draft_case_study()` writes a client-specific case study; it's held in `staging.py` and shown as an editable Accept/Reject card, then rendered into the **same branded template** at finalize.
3. **Data-driven skills slide** — `skills.py`, **Workforce-only** (see §4). Filled from `J2W_Skills_Inventory.xlsx` into the redesigned `skills_templates.pptx`.

### Built and working (this is a working v1 of the core loop)
- **Matching engine (rebuilt this session — the big change).** `matcher.plan()` + `relevance.py`:
  - Case candidates are **hard-filtered to the selected work types** (an MS deck shows only `MSS` cases; no AI-Pod leakage).
  - Ranked by **meaning** (per-ask OpenAI embeddings), **full-text + title** lexical overlap, and **direct skill→title term match** (the reliable signal for B2B terms like GCC / procurement / contracts).
  - **Industry is preferred**, not forced: same-industry cases always eligible; a cross-industry case only survives if its content match is strong (`CROSS_INDUSTRY_MIN`).
  - **Deep research + stakeholder profile drive the deck.** Each named need/skill → our best in-work-type case: **covered → leads the deck** ("research match"); **not covered → flagged "not in our library → generate."**
  - Deck is kept **tight** (sized to the number of matched needs; no off-function padding), de-duplicated (MMR), and each pick carries an **honest, person-specific "why it fits."**
- **Inputs & UI.** Optional **deep-research** upload and **stakeholder-profile** upload (PDF/text, parsed by `research.py`). On the build page: "Why this deck matches", "Not in our library — worth building", and a **search box** in the add-slide picker.
- **Case studies.** 160 cases; titles cleaned (no redundant domain tails); **no em-dashes anywhere** (hyphens only); **client names anonymized** on every slide via `anon_client()` (`CLIENT: Leading <Domain> <descriptor>`), never a real company.
- **Draft-with-AI.** Client-specific case study in the strict format (Title, Client|Domain|Function, Challenge ≤100w, Solution ≤100w, 6 capabilities, 3 results), rendered in the branded `case_study_v2` template.
- **Skills slides** — the 3 templates redesigned to match the case-study branding (red+teal bar, cards); markers + charts intact.
- **Add-slide picker** now excludes the 105 legacy master case slides, so **every case study renders in the new Python template regardless of how it's added** (auto-pick, suggestion, or library search).
- **Meeting log** (`meeting_log.py`) auto-writes one JSON per client+phase on generate.

### In progress / partially done
- **Persona / profile matching** — largely done via the profile upload (extract function/skills → direct match). `personas.py` (buyer-role detection) still exists and contributes a light boost, but the **profile-driven path is now primary**. A dedicated PERSONA *tag on every slide* + a persona picker in the form is not yet built.
- **Gap handling** — done: research/transcript-first, misses flagged as "not in our library," and **Draft-with-AI** is the "+ generate" action (it pre-fills the creator). The one nuance still open: an AI-drafted slide currently looks like an anonymized real case study — if the scenario is invented it should be **visibly labelled "illustrative approach,"** never presented as a delivered project (see §7).
- **Tag strengthening** — ongoing; matching now leans on semantic + direct-term more than raw tags, but richer tags still help.

### Not started
- Finalize/lock the master deck template design.
- Login / team-auth page (the salesperson identity is a placeholder — `current_salesperson()` in `app.py`).
- Full research *step* between "enter context" and "build" (today research is an **upload**, not an automated web-research stage — see §6).
- Persona as a first-class form field + per-slide tag.
- Anything beyond decks (proposals/RFQ/RFP/retros).

---

## 3. Folder structure (key files)

**App & flow**
- `app.py` — the Flask app. All HTML is inline (`NEW_FORM_BODY`, `BUILD_BODY`, `REVIEW_BODY`, `PREVIEW_BODY`). Routes: `/`, `/new`, `/build`, `/review`, `/finalize`, `/create_ai`, `/output/<f>`, `/library`, `/meetings`. Holds the research/profile upload wiring, the "why matched"/"missing" computation, `_ai_to_store_record()`, `_legacy_case_ids()`, `COVERAGE_THRESHOLD`.
- `matcher.py` — `plan(context, use_ai, priority_ids)`: always-in core + per-work-type standard blocks (from the registry, `CSxx`) + **case selection via `relevance`** (work-type gated, research-led, tight, de-duped). `_account_functions()`, `_is_cxo()`, `_case_reason()`.
- `relevance.py` **(new)** — the scoring core. `rank_cases()` (semantic + lexical + title-boost + industry/function, research-weighted over the mail), `best_cases()` (direct skill→title term match for per-need mapping), `coverage()` / `lexically_covered()` (work-type-aware "do we have this?"), `embed_texts()`, `_split_asks()`, `ask_count()`, `max_similarity()` (MMR dedup). Tuning knobs live here: `W_SEMANTIC, W_LEXICAL, W_TITLE, W_INDUSTRY, W_FUNCTION, MAIL_WEIGHT, CROSS_INDUSTRY_MIN`.

**Case-study library (content store)**
- `case_study_content_store.json` — 160 cases (`AIP`=AI Pods 37, `WFS`=Workforce 43, `MSS`=MS Solution 80). Each: title, domain, industry code, function, challenge, solution, 6 capabilities ({title,body}), 3 results, keywords.
- `case_library.py` — serves store records in the matcher's row shape (`candidate_rows`, `all_rows`, `all_cases`, `title_map`, `_search_text`).
- `case_embeddings.json` **(new)** — one embedding vector per case. **Rebuild whenever case text changes:** `py build_case_embeddings.py`.
- `build_case_study_store.py` — builds the store + engine-side Excel mirror from the owner's `Case_Studies_Master_IDed.xlsx` (source is in `Downloads`). Normalizes em/en dashes to `-`; cleans titles (`_clean_title`).
- `Case_Studies_Master_IDed.xlsx` — the case-study source spreadsheet (IDs in col 1).
- `case_study_v2.pptx` — the **branded case-study template** (red+teal split bar, Challenge/Solution cards, 6 capability cards, 3 results). Generated by `create_case_study_template.py` (**has a shebang → run with `py -3`**).
- `fill_case_study.py` — fills `case_study_v2.pptx` from one store record. `anon_client()` (name-free CLIENT descriptor), `build_mapping()`, `split_result()`, `split_capability()`.

**AI**
- `ai_matcher.py` — `extract_accelerators(notes)` (needs from research/notes), `extract_profile(profile)` (function/skills/current-role from a bio), `explain_fit(person_ctx, recipient, picks)` (per-case "why it fits THEM"). Legacy `refine()`/`extract_asks()` remain but are **unused**.
- `research.py` **(new)** — extract text from an uploaded PDF/text file (pypdf → PyMuPDF fallback). Used for both research and profile uploads.
- `slide_generator.py` — `draft_case_study(brief, context)` ("Create with AI"), `_copy_slide()`, placeholder templates. (`draft()` legacy gap path and the red "verification banner" are vestigial.)
- `staging.py` — pending AI slides (`add/find/promote/discard/all_items`). `/review` shows them; accept = promote.
- `eval_ericsson.py` **(new)** — automated scorecard against a real hand-mapped example. Run `py eval_ericsson.py`. **Note: it calls `matcher.plan` directly and bypasses the app's need-extraction, so it now under-reports — the real `/build` path is representative.**

**Data-driven skills**
- `skills.py` — `candidates(context)` (Workforce-only gate), `build_into(deck, order, cands)` (renders skills + store-case slides into the deck and reorders). Reads only the **aggregated** sheets of `J2W_Skills_Inventory.xlsx`.
- `skills_templates.pptx` — the 3 skills template slides (redesigned to match branding). Generated by `create_skills_templates.py` (**no shebang → run with `py`**).
- `J2W_Skills_Inventory.xlsx` — skills data. **Only the two aggregated sheets are used. The "Consultant Detail" sheet has individual names and must NEVER be read.**

**Master deck & registry**
- `WORKING_COPY_Master_Deck.pptx` — **the living master** the engine reads/writes (`assembler.SOURCE`). Core/standard/structural `CSxx` slides live here.
- `Master_Deck_Case_Study_Portfolio.pptx` — stale pristine backup; **never** use as source or re-stamp IDs from it.
- `J2W_CaseStudy_Portfolio_Metadata.xlsx` — the **registry** (sheet `Slide Registry`): one row per master slide with `include_rule` / `kind` / `std_group` / `section` / `work_types` / `keywords`. Drives *which* master slides go in; `kind == CASE_STUDY` rows are the 105 legacy case slides (now hidden from the add-picker).
- `assembler.py` — `build_deck(ids)` builds the `CSxx` slides from a copy of the master (drops non-kept incl. ID-less template slides, prunes media, atomic save).
- `tagged_library.json` / `library.json` — master-slide content/keywords (`CSxx`). `build_library.py` (`read_id`, `build`), `stamp_ids.py` (insert-safe ID assignment).

**Support**
- `personas.py`, `tagger.py` (FUNCTION/INDUSTRY taxonomy + `_score`), `synonyms.py`, `editor.py` (marker/text edits), `meeting_log.py`, `secrets_loader.py`.
- `backups/pre_match_rebuild/` — pre-rebuild copies of `matcher.py`, `case_library.py`, and the store.

---

## 4. The skills-slide rules (do not weaken without asking)

- **Both skills slides appear ONLY in a pure Workforce deck** — Workforce selected and nothing else (`skills.candidates` is gated to `work_types == {WORKFORCE}`). Any other mix → no skills slides.
- **Capability slide auto-includes** when a skill area matches the context/transcript (keyword-first, uncapped; industry matches capped at `CAP_INDUSTRY=3`; a `STOPWORDS` set stops generic words like "engineering" over-matching).
- **Footprint slide** only for an **existing partner** whose client name matches a Client Footprint row (has footprint data).
- Both **auto-add but are removable** in the suggestions panel (they ride in the order as `SK:` / `FP:` ids).
- **Stale data (>90 days, `last_verified`) flags a warning** on the slide.
- **Markers must stay in single text runs** — marker filling preserves only the first run's formatting per paragraph (`editor.set_text` / `skills.fill_markers`). Keep each `{{MARKER}}` in its own run/box.
- Template slides are tagged in **speaker notes**. Current code uses three kinds: `J2W_TEMPLATE: skill_deepdive`, `J2W_TEMPLATE: industry_strength`, `J2W_TEMPLATE: company_footprint` (the "capability" slide is `skill_deepdive`/`industry_strength`; the footprint slide is `company_footprint`). `find_template()` matches on these note tags — don't rename them without updating `skills.py`.

---

## 5. Key decisions and why

- **Content-library design, not positional.** Slides are matched by content/tags, never deck position. Every master slide carries a stable `J2W_ID: CSxx` in its speaker notes; store cases use `AIP/WFS/MSS`.
- **Keywords/tags were the original matching fuel** — still contribute, but the engine now leans more on **semantic meaning (embeddings) + direct skill→title term matching**, which proved far more reliable for B2B functional terms than tags or pure embeddings alone.
- **AI-written slides are quarantined until approved** — drafted → `staging` → shown as Accept/Reject → only then built into the deck. Nothing AI-written reaches a deck silently.
- **Meeting log stays behind team login at deploy** — it stores real client names (`meetings/J2W_<Client>_<Phase>.json`).
- **Skills data is aggregated-only, never individual names** — only the two aggregated sheets of the inventory are read; the "Consultant Detail" sheet is off-limits.
- **Canonical skills list still needs team review** — ~433 raw skill strings collapsed to ~238 distinct; the groupings (especially **SAP-by-module** and **generic role titles**) still need a human pass before they're trusted for matching.
- **Confidentiality: no real client/company names on any slide** — only "J2W"/"JoulesToWatts." Enforced structurally by `anon_client()`; the CLIENT line is always a generic descriptor.

---

## 6. Pending tasks to finish v1 (in order)

1. **Finalize the master template** — lock the visual design of the core/structural master slides.
2. **Scrub client names from all case studies** — *structurally done* (every slide anonymized via `anon_client`), but do a **content audit** of challenge/solution/results text for any lingering real names or identifying details.
3. **Strengthen tags** — richer, cleaner `keywords` per case (helps every matching path). Fold in the canonical-skills review from §5.
4. **PERSONA as a first-class tag + matcher** — this is *partly delivered* via the profile upload (function/skills → direct match). To finish: add a **persona field/picker** in the form (e.g., QA/BA, CFO, Head of Testing) and a **persona tag on every slide**, so the engine mixes/matches by *who we're meeting* even without a profile upload.
5. **Gap handling** — *largely delivered*: transcript/research-first; if missing, flag "asked but not present," suggest the nearest, and offer a **"Draft with AI"** button that pre-fills the creator. Remaining: the "illustrative approach" labelling from §7.
6. **Login page** — replace the `current_salesperson()` placeholder with real team auth; put the meeting log behind it.

---

## 7. The newest direction (owner's current top pain point): company research + better case-study matching

**The problem:** the case-study library may simply not contain a story the client needs.

**The plan, and where it stands now:**
- **(a) Better matching so near-fit cases surface for non-obvious matches — DONE (extends the tags/persona work).** The rebuilt engine reads full case text + meaning + direct skill terms, gated by work type, industry-preferred. Real proof: on the Merck/Karthik test it went from an all-wrong batch to surfacing his actual function (procurement, GCC, contracts) and correctly flagging real estate/facilities as missing.
- **(b) Flag gaps honestly — DONE.** "Not in our library — worth building" lists exactly the needs we can't prove, work-type-aware. It does **not** stretch a weak case to look like a match.
- **(c) A research step between "enter context" and "build" — PARTIALLY DONE, as an upload.** Today the salesperson uploads a **deep-research brief** (from ChatGPT/Claude) and the **stakeholder profile**; the engine reads them and feeds richer signal into the matcher. What's **not** built: an **automated web-research stage** inside the tool that researches the company itself.
- **The illustrative-approach slide — CAUTION, not fully solved.** Draft-with-AI can generate a research-informed slide for a gap, but it currently renders like an anonymized real case study. Per the owner's rule it **must be clearly distinct from a real case study and never presented as a delivered project.** Next session should add a visible "Illustrative approach — not a delivered engagement" label/band to AI-drafted slides whose scenario is invented, and keep the human-review gate (nothing research-derived goes silently into a client deck — the salesperson must verify it).

---

## 8. Future roadmap

- **Claude API deep research** in the gap area (automated, cited company research feeding the matcher). Note: Anthropic has **no embeddings endpoint** — embeddings would stay on OpenAI `text-embedding-3-small` (or Voyage); only the reasoning/research calls would move to Claude.
- **Profile input → research** (auto-research the named stakeholder + company).
- **Email / Outlook integration** to auto-fetch meeting transcripts/threads instead of pasting.
- **Expand beyond decks** — proposals, RFQ, RFP responses, retros.

---

## 9. Fragile / non-obvious — gotchas and "do not touch"

- **No auto-reload.** Kill port 5000 and re-run `py app.py` after any code change, or you're testing stale code.
- **`create_case_study_template.py` has a shebang → run with `py -3`** (bare `py` hits the Windows Store alias). **`create_skills_templates.py` has no shebang → run with `py`.**
- **`_copy_slide` copies shape XML but NOT image parts.** Any template meant for copying (case_study_v2, skills templates) must be **text/auto-shape only** — pictures become broken references. This is why the skills templates have no images.
- **Marker filling keeps only the first run's formatting** per paragraph — keep each `{{MARKER}}` in its own run/box.
- **`case_embeddings.json` must be rebuilt when case text changes** — `py build_case_embeddings.py`. If you edit the store and forget, meaning-match silently drifts.
- **Every `/build` makes ~4–5 OpenAI calls** (transcript embed, extract accelerators, extract profile, 2 coverage embeds, explain_fit). Needs `OPENAI_API_KEY`; each AI step **fails safe** (falls back to lexical/no-list) if the key is missing. Cost ≈ $0.002/build.
- **`eval_ericsson.py` under-reports now** — it bypasses the app's need-extraction. Judge quality from the real `/build` page, not the scorecard number.
- **Legacy master case slides (`CSxx`, `kind == CASE_STUDY`, ~105 of them) are stale duplicates** of the store cases (old template). They are filtered out of the add-slide picker (`app._legacy_case_ids()`); every case study now renders in `case_study_v2`. **Do not** re-add them to the picker or re-stamp IDs from `Master_Deck_Case_Study_Portfolio.pptx`.
- **`WORKING_COPY_Master_Deck.pptx` is the living master.** Back it up before testing anything that writes to it (the old AI-accept/promote path grew the master).
- **Skills inventory "Consultant Detail" sheet has individual names — NEVER read it.**
- **OpenAI key was exposed in chat — rotate before deploy.** Key lives only in `.env` (git-ignored); never hardcode it.
- **Vestigial / dead — do not revive without reason:** `/download` route, `_maybe_generate()`, the red "verification banner" in `slide_generator`, `slide_generator.draft()` (legacy gap path), `renderer.py`, `ai_fallback.py`, `FORM_HTML`, and `ai_matcher.refine()`/`extract_asks()`.
- **Tuning knobs** (if matching feels off): `relevance.MAIL_WEIGHT` (research vs mail), `W_TITLE`, `CROSS_INDUSTRY_MIN`, `DEDUP_GATE/DEDUP_WEIGHT`, `app.COVERAGE_THRESHOLD` (covered vs missing), `matcher.pick_cap` (deck size). Extraction is slightly non-deterministic run-to-run — expect minor churn on borderline picks.

---

## 10. How to resume

1. `py app.py` → open http://127.0.0.1:5000.
2. Try a real meeting: fill the form, attach a **research PDF** and/or **stakeholder profile**, and watch the "Why this deck matches" + "Not in our library" sections.
3. If you change case text: `py build_case_study_store.py` then `py build_case_embeddings.py`.
4. Read `BUILD_LOG.txt` (bottom entries) for the exact reasoning behind the current matching behaviour.

**Immediate next steps (owner's priorities):** (1) add the "illustrative approach" label to AI-drafted gap slides (§7); (2) persona as a form field + per-slide tag (§6.4); (3) login page + put the meeting log behind it (§6.6).
