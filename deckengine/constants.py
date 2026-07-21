# -*- coding: utf-8 -*-
"""
constants.py  --  form vocabularies, labels, and matching knobs used by the views.

Pure data (no Flask, no I/O beyond deriving lists from tagger). Kept in one place
so the web layer and templates share a single source for dropdown options etc.
"""

from deckengine import config
from deckengine.services.matching import tagger

# A need is "already covered" if a case matches it this semantically (cosine),
# OR our cases already use its words (lexical). We only flag a TRUE gap — clearly
# absent on BOTH — because a false gap (rebuild what we have) is the worst outcome,
# and broad capability areas naturally score ~0.5 against any one specific case.
COVERAGE_THRESHOLD = 0.50

# A need is "covered" by a case when the top capability match clears this cosine bar
# (or the need's word is literally in the case title). Below it, the need is an honest
# gap ("Not in our library"). Lower than COVERAGE_THRESHOLD because the shortlist query
# is the bare capability, so genuine matches land a touch lower.
CAPABILITY_COVER = 0.42

# generic management labels that aren't a case-study TOPIC — never pick or flag them
_GENERIC_NEEDS = {
    "project management", "program management", "budget management", "risk management",
    "stakeholder management", "change management", "general management", "people management",
    "team management", "operations management", "performance management", "cost management",
    "vendor management", "process improvement", "data analytics", "cost optimization",
}

OUTPUT_DIR = config.OUTPUT_DIR
CONTENT_STORE = config.CONTENT_STORE_JSON

# INDUSTRIES = the built-in matching TAXONOMY (codes tagger/relevance score against —
# an industry-boost, cross-industry threshold, and the skills-slide Excel mapping all
# key off these exact 8 codes). Never append a salesperson's free-typed industry here.
INDUSTRIES = list(tagger.INDUSTRY.keys())

# The New-deck form's "Other…" option value. The salesperson's real industry is then
# typed into the companion `industry_other` field, so ANY endpoint reading the form
# must resolve this sentinel before use — never treat it as an industry name (it
# reached the AI research prompt verbatim once; see web/view_helpers.resolve_industry).
INDUSTRY_OTHER = "__OTHER__"

FUNCTIONS = list(tagger.FUNCTION.keys())
WORK_TYPES = ["WORKFORCE", "AI_POD", "MS"]
OWNER = "Athithia"

# Friendly labels for the work-type codes (used in dropdowns / tables).
WT_LABELS = {"WORKFORCE": "Workforce", "AI_POD": "AI Pods", "MS": "Managed services"}

# Deck phase — fixed list, pick exactly one, in this order.
PHASES = [
    "Intro",           # sent to the client before a meeting, as preparation
    "First Meeting",   # introductory — industry-specific, tuned to the stakeholder
    "Second Meeting",  # focused on the specific thing the client showed interest in
    "Proposal",        # formal proposal stage
]


def all_industries():
    """The built-in taxonomy PLUS any salesperson-added "Other" industries, for the
    dropdown only. Called fresh per-request (not cached) so a newly-added custom
    industry shows up immediately, no restart needed. NOTE: a custom industry is a
    display label only — it does not get the industry-boost matching weight the 8
    built-in codes get (see industries.py)."""
    from deckengine.services.content import industries as _custom
    return INDUSTRIES + [i for i in _custom.load() if i not in INDUSTRIES]
