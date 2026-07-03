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

INDUSTRIES = list(tagger.INDUSTRY.keys())
FUNCTIONS = list(tagger.FUNCTION.keys())
WORK_TYPES = ["WORKFORCE", "AI_POD", "MS"]
OWNER = "Athithia"

# Friendly labels for the work-type codes (used in dropdowns / tables).
WT_LABELS = {"WORKFORCE": "Workforce", "AI_POD": "AI Pods", "MS": "Managed services"}

# Deck phase — fixed list, pick exactly one, in this order.
PHASES = [
    "Pre-read",        # sent to the client before a meeting, as preparation
    "First Meeting",   # introductory — industry-specific, tuned to the stakeholder
    "Second Meeting",  # focused on the specific thing the client showed interest in
    "Proposal",        # formal proposal stage
]
