# -*- coding: utf-8 -*-
"""
paste_parser.py  --  cut one pasted document into its slides.

The MS team hands over a whole deck's worth of content in one document, marked up
the way a human would:

    slide 1 - case study            <- a CATEGORY: they know the shape
    <the case study>

    slide 2: How we think before we build     <- a HEADING: they don't
    <three principles>

    slide 3
    <content, no label at all>

This module does ONE job: split the text into slides. It does not decide what shape a
slide is unless the salesperson named the category outright, and it never infers a
shape from a heading -- /builder/parse does that, with one AI call over the real
content, which is far better evidence than a title.

Two rules that matter:

  * A header is `slide <n>` followed by end-of-line or a separator (`-`, `:`, `.`).
    The separator is the guard: "Slide 2 of the deck covers..." has none, so it can
    never split the document.

  * The SLIDE NUMBERS validate each other. Real headers form an increasing sequence;
    a stray line that looks like one does not fit the chain. We keep the longest
    increasing run and drop the rest. Earlier versions instead capped the header at
    60 characters, then at 14 words -- both were arbitrary limits on content nobody
    controls, and the 60-char cap silently swallowed two real slides out of nine in
    the owner's first real document ("Slide 4: Case Study 1, Reconciliation automation
    with agentic AI" is 64 characters). Length was never the signal. Structure is.
"""

import re

# A header line: optional "slide"/"sl"/"#", a number, then EITHER end-of-line or a
# separator followed by whatever the salesperson wrote. No length limit -- a slide
# title can be as long as it likes.
_HEADER = re.compile(
    r"^\s*(?:slide|sl)\s*#?\s*(\d+)\s*(?:[-–—:.)\]]+\s*(?P<label>.*?))?\s*$",
    re.IGNORECASE)

AUTO = "auto"          # "the label didn't name a category" -- the caller reads the content

# The ONLY labels that skip the AI. An exact, unambiguous naming of a category we have.
# Anything else -- a heading, an odd phrasing, nothing at all -- goes to the intelligence
# layer, which reads the slide's actual content. Fuzzy-matching labels was a mistake: the
# heading "How we would engage with Voya" once matched "named list with stats" on the
# single word "with", and the slide got the wrong template because of a preposition.
_ALIASES = {
    "case_study": ["case study", "casestudy", "case", "customer story", "client story",
                   "success story", "case study slide"],
    "four_box": ["four box", "four box section", "for box section", "4 box", "fourbox",
                 "four way", "four way breakdown", "quadrant", "four sections"],
    "roadmap_board": ["roadmap", "road map", "roadmap board", "phased roadmap", "phases",
                      "board", "kanban", "timeline", "phase plan"],
    "box_grid": ["box grid", "grid", "boxes", "sections", "pillars"],
    "pillar_deepdive": ["capability deep dive", "deep dive", "deepdive", "pillar",
                        "pillar deep dive", "pillar deepdive", "capability"],
    "scored_list": ["scored list", "named list", "list with stats",
                    "named list with stats"],
    "stat_overview": ["stat overview", "stats overview", "stats", "statistics",
                      "headline stats", "headline stats overview", "metrics", "kpis", "kpi"],
    "data_table": ["data table", "table", "matrix"],

    # ── the ten style-guide shapes (owner's designs, 2026-07-10) ──────────────
    "pain_point_list": ["pain point list", "pain points", "problem list", "problems",
                        "the problem", "pain point", "challenges"],
    "platform_overview": ["platform overview", "platform", "product overview",
                          "what is", "overview"],
    "before_after_split": ["before after", "before and after", "before after split",
                           "before/after", "workflow comparison", "transformation"],
    "comparison_split": ["comparison split", "capabilities and comparison",
                         "split comparison", "capability comparison"],
    "pillar_grid": ["pillar grid", "pillars grid", "four pillars", "capability pillars",
                    "numbered pillars", "feature grid"],
    "option_columns": ["option columns", "options", "three column comparison",
                       "architecture comparison", "option comparison"],
    "agent_architecture": ["agent architecture", "agents", "architecture",
                           "component architecture", "system architecture"],
    "governance_list": ["governance list", "governance", "governance layer",
                        "timeline list", "layers"],
    "guardrail_columns": ["guardrail columns", "guardrails", "governance columns",
                          "themed columns", "framework"],
    "opportunity_cards": ["opportunity cards", "opportunities", "use cases",
                          "opportunity outcome", "opportunities and outcomes"],
}


def _norm(text):
    """Lowercase, punctuation stripped, single-spaced -- so 'Four-Box Section.' and
    'four box section' compare equal."""
    return " ".join(re.sub(r"[^a-z0-9]+", " ", (text or "").lower()).split())


def match_template(label):
    """A label -> a CONTENT_TEMPLATES key ONLY when it names a category exactly, else
    AUTO. AUTO is not a failure: it means "ask the content", which the caller does."""
    norm = _norm(label)
    if not norm:
        return AUTO
    for key, aliases in _ALIASES.items():
        if norm == key.replace("_", " ") or norm in (_norm(a) for a in aliases):
            return key
    return AUTO


def _candidates(lines):
    """Every line that LOOKS like a header: [(line_index, number, label)]."""
    out = []
    for i, line in enumerate(lines):
        m = _HEADER.match(line)
        if m:
            out.append((i, int(m.group(1)), (m.group("label") or "").strip()))
    return out


def _longest_increasing(cands):
    """Keep the largest set of candidates whose slide NUMBERS increase in document order.

    This is what makes the parser robust without capping anything: a real deck numbers
    its slides 1, 2, 3...; a stray line that merely reads like a header ("Slide 2: as
    discussed above.") can't extend that chain and is dropped. Classic O(n^2) longest
    increasing subsequence -- n is the number of slides, so it's free.
    """
    if not cands:
        return []
    n = len(cands)
    best = [1] * n              # length of the best chain ending at i
    prev = [-1] * n
    for i in range(n):
        for j in range(i):
            if cands[j][1] < cands[i][1] and best[j] + 1 > best[i]:
                best[i], prev[i] = best[j] + 1, j
    end = max(range(n), key=lambda i: best[i])
    chain = []
    while end != -1:
        chain.append(cands[end])
        end = prev[end]
    return list(reversed(chain))


def parse(text):
    """Cut `text` into slides. Returns a list of:

        {"number", "label", "heading", "template", "matched", "content", "preview"}

    `template` is a CONTENT_TEMPLATES key when the label named a category outright,
    otherwise AUTO -- the caller resolves those from the content. `matched` says whether
    the salesperson's own label decided the shape.

    A label that isn't a category is the slide's HEADING: it's kept in `heading` and
    prepended to `content`, so the generator titles the slide what the salesperson
    called it instead of inventing a title.

    With no headers at all, the whole paste is ONE slide -- pasting a single slide must
    keep working exactly as it did.
    """
    text = (text or "").strip()
    if not text:
        return []

    lines = text.splitlines()
    cuts = _longest_increasing(_candidates(lines))
    if not cuts:                    # no headers -> one slide, shape read from the content
        return [_slide(1, "", AUTO, False, text)]

    slides = []
    for pos, (line_i, number, label) in enumerate(cuts):
        end = cuts[pos + 1][0] if pos + 1 < len(cuts) else len(lines)
        content = "\n".join(lines[line_i + 1:end]).strip()
        if not content:
            continue                # a header with nothing under it isn't a slide
        template = match_template(label)
        if template == AUTO and label:
            # a heading, not a category -- keep it WITH the content so the slide is
            # titled what the salesperson called it
            slides.append(_slide(number, label, AUTO, False,
                                 label + "\n" + content, heading=label))
        else:
            slides.append(_slide(number, label, template, template != AUTO, content))
    # text before the first header is preamble, deliberately ignored -- a cover note, not
    # a slide. If every header turned out empty, treat the whole paste as one slide.
    return slides or [_slide(1, "", AUTO, False, text)]


def _slide(number, label, template, matched, content, heading=""):
    return {"number": number, "label": label, "heading": heading, "template": template,
            "matched": matched, "content": content,
            "preview": _first_line(content)}


def _first_line(content, limit=110):
    """The opening of the slide's content, so the split can be eyeballed."""
    flat = " ".join(content.split())
    return flat[:limit] + ("…" if len(flat) > limit else "")
