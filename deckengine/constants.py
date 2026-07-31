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

# generic management labels that aren't a case-study TOPIC — never pick or flag them.
# A NEED with one of these names matches dozens of unrelated cases equally (no
# discrimination), so the wrong one wins on a tiebreak — the ARKO/Deepak failure
# (2026-07-27): a profile's shallow "Top Skills" tags ("Finance Analytics", "Process
# Optimization", "Data Governance"...) got extracted as needs and matched a market-data
# case / PE due-diligence case over the close/reconciliation/AP cases that actually fit.
# The real fix is in ai_matcher.extract_brief (extract the concrete workflow underneath a
# label, not the label); this stoplist is the deterministic backstop for umbrella labels
# that still slip through. Kept to pure management/transformation umbrellas — NOT terms
# that can be a genuine capability in some account (e.g. "data governance" platform work).
_GENERIC_NEEDS = {
    "project management", "program management", "budget management", "risk management",
    "stakeholder management", "change management", "general management", "people management",
    "team management", "operations management", "performance management", "cost management",
    "vendor management", "process improvement", "data analytics", "cost optimization",
    # umbrella labels added 2026-07-27 (ARKO): never a case-study-shaped capability on their own
    "process optimization", "resource management", "finance transformation",
    "finance transformation leadership", "digital transformation",
    "digital finance transformation", "business process reengineering",
    "agile methodologies", "transformation roadmapping", "value realization",
    "finance process design", "finance technology strategy", "finance systems implementation",
    "program leadership", "operational excellence", "business process management",
    "process automation", "data governance", "agile project management",
    "cross-functional leadership", "cross-functional team leadership", "data governance frameworks",
    "agile project management in finance",
    # HR / talent umbrella labels added 2026-07-29 (NTT DATA / Tirumala): broad
    # People-&-Culture PROGRAM areas, never a single case-study-shaped workforce
    # capability on their own -- the real, matchable need is the specific recruitment
    # workflow underneath (RPO, hire-train-deploy, high-volume hiring, technical
    # recruiting, workforce planning). Bare "talent acquisition" is deliberately NOT
    # here -- it IS the core capability; only these umbrella FORMS are filtered.
    "talent acquisition strategy", "talent management", "talent strategy",
    "employee engagement", "employee engagement programs", "learning and development",
    "learning and development initiatives", "learning and development strategy",
    "hr transformation", "human resources transformation", "people and culture",
    "employer branding", "organizational development", "corporate communications",
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
