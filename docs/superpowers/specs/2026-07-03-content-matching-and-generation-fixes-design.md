# Design — Content Matching & Case-Study Generation Fixes

Date: 2026-07-03
Status: Approved (pending spec review)
Owner: Athithia (non-technical) · relayed by Krishna

## Context

Two core issues were reported from the user's perspective:

**Issue 1 — Matching surfaces term-matches, not domain best-fit.** Given a client
profile + deep-research brief + transcript, the deck-leading content picks are
chosen by `relevance.best_cases()`, which matches on the *bare capability name* and
scores literal keyword overlap far above semantics/domain (term boost up to +1.05
vs cosine 0–1; industry only +0.02). The brief's synergy mapping and **mismatch
flags** (e.g. "computer vision for electrical products vs. quality web") are never
extracted or used. Result: a case that merely contains a term wins over the
right-domain case, even though the correct case exists in the library.

**Issue 2 — AI case study is shallow.** The generator `slide_generator.draft_case_study`
receives only a thin brief + industry/recipient + a **transcript truncated to 2000
chars**. The **deep-research brief and client profile are never passed at all** —
they are read at `/build`, used for matching, then discarded. With almost no facts,
the prompt ("the human supplies the facts; infer metrics from benchmarks") makes the
model invent a generic, definition-level case.

Audit of what reaches generation today (the table to fix — every row must become
"yes/used"):

| Input | Today | Target |
|---|---|---|
| Typed brief | yes | yes |
| Industry / recipient / function | yes | yes |
| Transcript | **partial** (truncated 2000 chars) | **full**, synthesized |
| Deep research brief | **no** | **yes**, synthesized |
| Client profile | **no** | **yes**, synthesized |
| Gap-detection output | **partial** (name only) | **yes** — fills the specific gap with researched content |

Both trace to one root gap: **the deep-research brief and profile are treated as
throwaway matching fuel** — reduced to bare names, used once, discarded — so neither
ranking (can't see domain/mismatch) nor generation (never receives them) can use the
brief's actual intelligence.

## Goals / success criteria

1. Deck-leading content is selected by **domain + semantic + expressed-interest**
   fit, and the brief's **mismatch flags actively exclude** wrong-domain cases.
2. Every generation input (full transcript, deep research, profile, the specific
   gap) is **passed and meaningfully used** — no truncation, no discarding.
3. Generated case studies are **specific, synthesized, professional / company-
   standard** — grounded in the client's real context, not generic definitions.
4. Each pick is **explainable** (domain / prior-interest / role) so mismatches are
   easy to catch.
5. **No existing flow breaks**; every AI dependency **fails safe** (and, unlike
   today, degraded states are surfaced rather than silent).

Approach approved by user: **B — LLM judgment over a cheap algorithmic shortlist,
with the rebalanced algorithm as the offline fallback**; generation upgraded to a
stronger model.

## Design

### A. Shared foundation — structured brief understanding

New AI function `ai_matcher.extract_brief(research, profile, transcript)` → one call
returning a structured "matching brief":

```
{
  "needs":              [{"name","description","domain","use_case"}],
  "avoid":              [{"capability","reason"}],   # the mismatch flags, as signals
  "expressed_interest": ["..."],                     # accelerators discussed in the transcript
  "account":            {"industry","role","company_context"}
}
```

- **Replaces** the current `extract_accelerators` + `extract_profile` calls on the
  `/build` path (net calls stay ~flat).
- **Fail-safe:** on error/empty, fall back to today's name-only extraction
  (`extract_accelerators`/`extract_profile`) so `/build` still works offline.
- Prompt instructs the model to read the deep-research brief's synergy mapping and
  mismatch flags and emit them structurally; keep `avoid` conservative (only explicit
  or strongly-implied misfits).
- **Reference priority** (per the reported expected behavior) is honored here:
  profile-only → needs come from the profile (role/company context); when a deep-
  research brief is present it is **primary**, and the transcript contributes
  `expressed_interest` (prior sales context) rather than driving the needs.

### B. Issue 1 — domain-aware, mismatch-aware selection

Flow inside `deckengine/services/web/decks.py::build()` (replacing the current
`best_cases`-only priority step):

1. **Structured brief** via `extract_brief` (§A).
2. **Cheap shortlist per need** — `relevance.best_cases` reworked:
   - Query vector embeds `name + domain + use_case` (not the bare 2-word name).
   - **Rebalance weights:** literal-term boost `0.35 → ~0.12` (and title-only
     `0.25 → ~0.10`); **industry `0.02 → ~0.15`**; function boost raised. Semantics +
     domain now lead; keyword overlap is a minor tiebreak.
   - Work-type gated (`allowed_ids`). Return top ~8 candidates per need (not just 1).
3. **LLM re-rank (one call)** `ai_matcher.rank_shortlist(brief, shortlist_by_need)`:
   - Input: structured brief (needs w/ domain/use-case, **avoid flags**,
     **expressed_interest**, account) + shortlist candidates (id, title, domain,
     industry, function, blurb).
   - Output per need: best-fit case id(s) that genuinely fit, **excluding any that
     violate an avoid flag**, each with `reason` + `signal ∈ {domain,interest,role}`.
     A need with no real fit → returns none → becomes a true gap.
   - Provides the `reason`+`signal` for the **priority (need-matched) picks**,
     replacing `explain_fit` for those. The remaining fill cases (chosen by
     `rank_cases`, not tied to a specific need) keep their existing deterministic
     `matcher` reasons; `explain_fit` is dropped from the hot path (net calls stay flat).
   - **Fail-safe:** on error, use the algorithmic shortlist's top pick per need with
     a generic reason.
4. **Avoid flags applied to the whole ranking**, not only priorities: pass `avoid`
   into `matcher.plan` → `relevance.rank_cases`, which zeroes/demotes cases whose
   domain+capability match an avoid flag — so a flagged mismatch can't enter via the
   fill step either.
5. **Expressed interest weighted** even when a brief exists (fixes the current
   `if not research_needs and not profile_needs` fallback-only bug): expressed-
   interest items are added to the needs considered for priority picks.
6. **Explainability surfaced:** the re-rank's `reason`+`signal` feed the "Why this
   deck matches" rationale (`rationale[].why` / `.fit`); optionally a small
   "skipped X — flagged as mismatch" note.

`covered` needs → `priority_ids` (lead the deck, as today). Uncovered → `missing`
("Not in our library"). `matcher.plan(..., priority_ids, research=lead_research,
avoid=...)` otherwise unchanged.

### C. Issue 2 — real context into generation

1. **Persist build context** — new module `deckengine/services/build_context.py`:
   `save(build_id, ctx)` / `load(build_id)` writing `build_context/<build_id>.json`
   = `{research, profile, transcript, industry, recipient, functions, client_name}`.
   - File-backed (works across gunicorn's 2 workers; an in-memory dict would not).
   - New writable dir `config.BUILD_CONTEXT_DIR` (env-overridable like the others,
     git-ignored, kept via `.gitkeep`).
   - Light pruning: on `save`, delete entries older than 7 days (bounded growth).
   - Written at the end of `/build` once `research_text`/`profile_text`/`transcript`
     are known, keyed by the `build_id` already generated there.
2. **`/create_ai` reloads by `build_id`** — `build.js::createAI()` sends `build_id`
   (already in `SERVER_CTX`/`BUILD_ID`); the server loads the full context.
   - **Fail-safe:** if the file is missing (old build / restart / pruned), degrade to
     today's browser-sent transcript + industry.
3. **`draft_case_study` reworked** (`slide_generator.py`):
   - New params: `research`, `profile`, full `transcript`, and the specific `gap`
     (need name + description, from `prefillAI`).
   - Prompt changed from *"the human supplies the facts; infer from benchmarks"* to
     *"synthesize the provided deep-research, profile, and transcript into a case
     grounded in THIS client and THIS gap; use benchmarks only to fill genuine metric
     gaps, and mark those as illustrative."* Keep the strict format (6 caps / 3
     results, anonymized, no em-dashes) and the branded `case_study_v2` render.
   - Bounded token caps: research ~6k, profile ~4k, transcript ~6k, brief ~1.5k.
   - **Fail-safe:** unchanged — on API error, return the normalized placeholder.
4. **Model upgrade:** this generation call uses `gpt-4o` (config constant
   `GEN_MODEL`), applied to **both** create paths (gap-card + manual panel).

### Module / interface summary

- `ai_matcher.extract_brief(research, profile, transcript) -> dict` (new)
- `ai_matcher.rank_shortlist(brief, shortlist_by_need) -> dict` (new; absorbs
  `explain_fit`)
- `relevance.best_cases(...)` — reworked query vector + weights; returns top-N
- `relevance.rank_cases(..., avoid=...)` — new avoid demotion
- `matcher.plan(..., avoid=...)` — threads avoid through
- `deckengine/services/build_context.py` — `save` / `load` (new)
- `slide_generator.draft_case_study(brief, context)` — richer context + prompt +
  `GEN_MODEL`
- `deckengine/services/web/decks.py::build()` / `api.py::create_ai()` — wire the above
- `config.py` — `BUILD_CONTEXT_DIR`, `GEN_MODEL`

### Error handling / fail-safes (explicit)

- Every new AI call wrapped; on failure falls back to the deterministic path
  (algorithmic shortlist / name-only extraction / today's generation).
- Build-context load failure → graceful degrade to browser transcript.
- Where a degrade materially changes output (brief couldn't be parsed; context
  missing), surface a small non-blocking notice in the UI rather than silently
  degrading — per the lesson from the earlier silent-PDF-failure bug.

### Testing & quality benchmark

- Smoke suite stays the guard. The **golden matcher IDs will change** (intended) —
  rebaseline **only after** confirming the target scenario improved.
- New tests (hermetic, offline/mocked AI): avoid-flag exclusion (a flagged case is
  not picked), `build_context` save/load round-trip, `draft_case_study` receives
  research/profile (assert they reach the prompt builder).
- **Quality benchmark** (user's ask): pin one "good vs current" target case study and
  one target matching scenario (ideally the computer-vision-electrical-vs-web
  example) to measure fixes against before shipping.
- Manual AI-on verification via the real `/build → Draft with AI` flow with a live key.

## Out of scope

- No UI/visual redesign of the deck templates (branding/layout unchanged; "professional"
  is achieved via content depth + correct selection).
- No change to work-type gating, core `rank_cases` weighting beyond the avoid signal,
  or the assembly/rendering pipeline.
- No auth / salesperson wiring.

## Rollout / verification order

1. Foundation (`extract_brief`) + fallbacks, tests green.
2. Issue 1 ranking rework + avoid threading; verify target scenario; rebaseline golden.
3. Issue 2 build-context persistence + generation rework + model upgrade; verify depth.
4. Full smoke + gunicorn + one live browser build end-to-end.
