# Design — UI Redesign, Pass 1: the shell and the tokens

Date: 2026-07-10
Status: Approved (pass 1 scope), pending spec review
Owner: Athithia (non-technical) · relayed by Krishna

## The problem

The app looks sparse and unfinished on a wide screen. Three decisions cause it, and the
white space is a symptom of all three rather than a problem in itself.

**1. Every page is built as a landing page.** Thirteen templates open with the same block:
an `eyebrow`, a 46px `h1.display`, and a `.lede` capped at `max-width: 560px`. That is
roughly 200px of vertical space before the first control, on a tool used dozens of times a
week, and the 560px cap against a 1180px page is the ragged right edge visible on
Templates, AI history and Deck repository.

**2. The container fights the screen.** `.page { max-width: 1180px }` plus an 84px sidebar
leaves ~290px of dead gutter on each side of a 1844px display. Widening it alone makes the
emptiness worse — the same three cards, stretched.

**3. The chrome is decorative.** In `_shell.html` the search button, the adjust button, the
"AT / J2W" avatars and the Settings nav item have no behaviour. Four affordances that
promise something and do nothing. This is most of why the app "feels like things are
missing": it is showing controls that are not there.

Underneath those: the sidebar is 84px of unlabelled icons with Dashboard — the home page —
exiled below a spacer to the bottom; and the right rail on five pages is a static
"How it works" poster occupying the most valuable column on screen.

## The idea

The engine draws every slide in **Oswald** (headings) and **Raleway** (body), with a red
accent bar beside a bold condensed title and a two-tone red/teal bar along the top edge.
The app is set in **Inter** and **Space Grotesk**, and has never once used that bar.

The tool and the artifact it produces do not share a voice.

Pass 1 makes them share one. The interface adopts the deck's typefaces and its single
structural mark. Everything else gets quieter and denser, not louder.

**The signature:** the two-tone bar, at the master deck's real **67:33 red-to-teal ratio**
(measured 2026-07-07: red 8.89in : teal 4.44in on a 13.33in slide — *not* 50/50, and not
the style guide's stated teal-left). It appears exactly twice: as the top edge of the app
shell, and as the active-item indicator in the sidebar. One accessory, used everywhere,
quietly. It is not repeated on cards, buttons, or headers.

## Scope

Pass 1 is **the shell and the tokens**. It touches `templates/_shell.html` (33 lines) and
`static/app.css` (461 lines), plus the removal of the hero block from each page template.
No page is rewritten; no route, service, or test changes.

Out of scope, deliberately, for later passes: the Library table, stateful right rails,
empty states, the command palette.

---

## 1. The app shell

Three zones replace the centred column.

```
┌────┬───────────────────────────────────────────┬──────────────┐
│    │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓░░░░░░░░░░░│  ← brand bar │
│ nav│ Library / All slides       232   [ + New ]│              │
│    ├───────────────────────────────────────────┤   context    │
│ 240│                                           │   rail       │
│ px │  work area — fluid, 1600px cap            │   320px      │
│    │                                           │   collapsible│
└────┴───────────────────────────────────────────┴──────────────┘
```

**The header bar replaces the hero block.** 48px: breadcrumb, page title inline, count if
the page has one, primary action pinned right. Reclaims ~180px of vertical on every page.
The `eyebrow` / `h1.display` / `lede` trio is deleted from all thirteen templates; the lede
copy that carries real information moves into the page body or the rail, and the copy that
only restates the title is cut.

**The container becomes fluid:** `min(1600px, 100% - 64px)`, with 32px minimum gutters.

**The right rail** keeps its current 320px but its content becomes page-specific in a later
pass. In pass 1 it simply stops being `position: static` inside `.form-layout` and becomes a
real shell column that can collapse. The five pages that already have a `.side-card` keep
their content; the pages without one (Library, Templates, AI history, Deck repository) get
an empty, collapsed rail — no filler.

## 2. The sidebar

84px → **240px expanded, 72px collapsed**, with the state stored in `localStorage`.
Labels beside every icon. Four groups, matching how the work actually runs:

| Group | Items |
|---|---|
| — | **New deck** (primary button, pinned top) |
| Home | Dashboard |
| Create | New deck, Custom Slide Builder |
| Assets | Slide library, Templates |
| Records | AI history, Deck repository |

Dashboard moves to the top, where a home page belongs. The active item is marked by the
two-tone bar as a 3px left edge — red over teal, same 67:33 split.

**Settings is removed.** It is a link to `#`. It comes back when there is a settings page.

## 3. Tokens

### Type

Both families are already on Google Fonts, so this is a swap in the existing `@import`,
not a new dependency.

| Role | Now | Pass 1 | Why |
|---|---|---|---|
| Display / headings | Space Grotesk | **Oswald** | the deck's own title face |
| Body / UI | Inter | **Raleway** | the deck's own body face |
| Data / IDs | Inter | **IBM Plex Mono** | `CS01`, `MSS042`, `AIP007` appear on every card and every row; a proportional face makes them unscannable |

Scale, tightened for a working tool:

```
--fs-display  32px   (dashboard greeting only; was 46px on every page)
--fs-h1       20px   (page title in the header bar)
--fs-h2       16px   (card titles)
--fs-body     14px   (was 15px)
--fs-small    12.5px
--fs-mono     12.5px
```

### Density

| Token | Now | Pass 1 |
|---|---|---|
| `--r-lg` (cards) | 24px | **12px** |
| `--r-md` | 16px | **10px** |
| `--r-sm` | 12px | **8px** |
| `--r-pill` | 999px | 999px (buttons only) |
| card padding | 22px 24px | **16px 18px** |
| `.grid` gap | 18px | **12px** |
| `--side` | 84px | **240px / 72px** |

24px radii and 22px padding on a card holding three lines of text is what makes the Library
read as furniture rather than data.

### Colour

**Unchanged.** The palette is fixed and correct. Pass 1 adds only the two brand values the
slides already use, so the shell can draw the bar with the real deck colours rather than
the UI's teal:

```
--brand-red   #D62839    (master deck red)
--brand-teal  #2A9D8F    (master deck teal)
```

Note these differ from the UI's `--teal: #2C6E66`. That is deliberate: `--teal` stays the
interaction colour; `--brand-*` are the artifact's colours and are used only for the bar.

## 4. Dead controls

Deleted from `_shell.html`: the search button, the adjust button, the avatar pair, the
Settings nav item.

The search button is not merely removed — it is a promise the app should keep. A global
command palette over 232 slides, 18 templates and every generated deck is the highest-value
addition available, and it is **pass 4**. Removing the fake control now is honest; leaving
it is not.

---

## What this does not fix

Stated plainly, so nobody expects it to:

- **The Library still browses when it should scan.** 232 slides in three-across cards. It
  needs a table, and that is pass 2.
- **The right rails still say "How it works" forever.** Pass 3.
- **Empty states are still dead ends.** Pass 3.
- **There is still no search.** Pass 4.

Pass 1 fixes the hero block, the container, the sidebar, the type, the density, and the
lying chrome. That is most of the visible damage, because the hero block and the centred
container cause most of it.

## Risks

**Raleway at 14px in dense UI.** It is a display-leaning humanist sans and is less legible
in tight tables than Inter. If the Library table (pass 2) reads poorly, the fallback is to
keep Oswald headings and the brand bar, and revert body copy to Inter — the "unify headings
only" option. This is a real risk and the reason type is in pass 1 rather than later: it
needs to be seen on real pages early, while reverting is still cheap.

**Thirteen templates lose their hero block.** Mechanical, but it touches every page. The
smoke test drives `/build → /review → /finalize` and asserts every route returns 200; it
does not assert on markup, so it will not catch a mangled template. Each page must be
loaded and looked at.

**Two font families added, two removed.** Net zero requests. Oswald and Raleway are already
what the `.pptx` output specifies, so a designer opening a generated deck beside the app
will now see one system.

## Verification

- Every route returns 200 and renders: `/`, `/builder`, `/library`, `/templates`,
  `/staging`, `/meetings`, `/dashboard`, plus `/build → /review` with a real deck.
- `python -m pytest tests/ -q` stays green (48 tests). Nothing here touches services.
- Screenshot each of the seven pages at 1844px and at 1280px, before and after.
- Confirm the header bar reclaims ≥150px of vertical on Templates and Deck repository.
- Confirm no page has a horizontal scrollbar at 1280px.
- Keyboard: the sidebar collapse toggle is focusable, the active item is announced.
- `prefers-reduced-motion` respected on the sidebar transition.
