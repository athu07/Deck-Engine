# -*- coding: utf-8 -*-
"""
personas.py  --  Persona (buyer-role) tagging + matching.

A deck for a QA Lead should look different from one for a CFO. This module:

  1. Defines the 8 personas J2W sells to (PERSONAS).
  2. detect()      -- reads WHO we're meeting (the form's "recipient" field) and
                      the "more information" notes/transcript, and returns the
                      persona(s) in play.
  3. score_boost() -- given the detected persona(s) and a slide's registry row,
                      returns a relevance bonus (so persona-relevant case studies
                      rank higher).
  4. tag_slide()   -- which persona(s) a slide naturally speaks to, derived from
                      its function / work_type / keywords. Stored as a tag so it's
                      visible and editable, and auto-applies to new slides.

Persona relevance is DERIVED from a slide's existing function/work_type/keywords
— so there is NO need to hand-tag all 105 slides. Add a slide and it just works.

HOW TO EXTEND (safe for a non-developer):
  * To recognise more job titles, add phrases to a persona's "aliases".
  * To change which slides a persona cares about, edit its "functions" /
    "work_types" / "topics". Keep everything lowercase except the FUNCTION /
    WORK_TYPE codes (those are UPPERCASE to match the registry).
"""

import re

import synonyms

# ─────────────────────────────────────────────────────────────────────────────
# The 8 personas. Keyed by a stable CODE; "label" is what the UI shows.
#   aliases    : phrases that, if found in recipient/notes, name this persona
#   functions  : registry primary_function codes this persona cares about
#   work_types : registry work_type codes this persona leans toward
#   topics     : extra concept words that signal relevance (synonym-expanded)
# ─────────────────────────────────────────────────────────────────────────────
PERSONAS = {
    "ENGINEERING_HEAD": {
        "label": "Engineering / Delivery Head",
        "aliases": [
            "head of engineering", "engineering head", "vp engineering",
            "vp of engineering", "director of engineering", "engineering director",
            "engineering manager", "delivery head", "head of delivery",
            "delivery manager", "head of software", "software engineering head",
        ],
        "functions": ["DEVOPS_CLOUD", "ENGINEERING_DESIGN", "GCC_SETUP",
                      "QUALITY_ENG_TESTING"],
        "work_types": ["AI_POD", "WORKFORCE"],
        "topics": ["delivery", "architecture", "engineering", "platform",
                   "scalability", "velocity", "gcc", "ai pod", "devops",
                   "ci/cd", "team augmentation"],
    },
    "CIO_CTO": {
        "label": "CIO / CTO",
        "aliases": [
            "cio", "cto", "chief information officer", "chief technology officer",
            "chief technical officer", "technology head", "head of technology",
            "head of it", "vp technology", "vp of technology", "it head",
        ],
        "functions": ["DATA_AI_PLATFORM", "DEVOPS_CLOUD"],
        "work_types": ["AI_POD", "MS"],
        "topics": ["strategy", "modernization", "cloud migration", "security",
                   "platform", "transformation", "roadmap", "total cost",
                   "aiops", "observability", "generative ai"],
    },
    "QA_LEAD": {
        "label": "QA / Test Lead",
        "aliases": [
            "qa lead", "test lead", "head of qa", "quality lead", "qa manager",
            "test manager", "head of testing", "qa head", "quality head",
            "quality assurance", "sdet", "test architect",
        ],
        "functions": ["QUALITY_ENG_TESTING"],
        "work_types": ["AI_POD"],
        "topics": ["test automation", "performance testing", "pentesting",
                   "quality", "defect", "coverage", "mobile qa", "regression"],
    },
    "DATA_AI_LEAD": {
        "label": "Data / AI Lead",
        "aliases": [
            "head of data", "head of ai", "chief data officer", "cdo",
            "data science", "data scientist", "ml lead", "ai lead",
            "head of analytics", "data engineering lead", "head of machine learning",
            "ai head", "analytics head",
        ],
        "functions": ["DATA_AI_PLATFORM"],
        "work_types": ["AI_POD", "MS"],
        "topics": ["machine learning", "generative ai", "rag", "agentic", "nlp",
                   "forecasting", "data platform", "analytics", "computer vision",
                   "model"],
    },
    "FINANCE_HEAD": {
        "label": "Finance / CFO",
        "aliases": [
            "cfo", "chief financial officer", "finance head", "head of finance",
            "finance director", "controller", "vp finance", "vp of finance",
            "financial controller", "head of fp&a",
        ],
        "functions": ["FINANCE_OPS"],
        "work_types": ["MS"],
        "topics": ["roi", "cost savings", "invoice", "ap/ar", "efficiency",
                   "finance", "equity research", "budget"],
    },
    "TALENT_HR_HEAD": {
        "label": "HR / Talent Head",
        "aliases": [
            "chro", "head of hr", "head of talent", "talent acquisition",
            "hr head", "recruitment head", "head of recruiting", "people officer",
            "hr director", "ta head", "head of people", "talent head",
            "head of ta",
        ],
        "functions": ["TALENT_ACQUISITION"],
        "work_types": ["WORKFORCE"],
        "topics": ["hiring", "talent acquisition", "rpo", "c2h", "attrition",
                   "talent pipeline", "staffing", "workforce", "recruiting"],
    },
    "PROCUREMENT_OPS": {
        "label": "Procurement / Supply Chain",
        "aliases": [
            "head of procurement", "procurement head", "procurement", "cpo",
            "chief procurement officer", "supply chain", "supply chain head",
            "coo", "chief operating officer", "operations head", "head of operations",
            "ops head", "head of supply chain",
        ],
        "functions": ["SUPPLY_CHAIN_OPS"],
        "work_types": ["MS"],
        "topics": ["supply chain", "procure to pay", "demand planning",
                   "inventory", "operations", "logistics", "material planning"],
    },
    "CEO_FOUNDER": {
        "label": "CEO / Founder",
        "aliases": [
            "ceo", "founder", "co-founder", "cofounder", "chief executive",
            "managing director", "president", "owner", "promoter",
        ],
        # CEO/founder cares about the story & outcomes, not a single function.
        "functions": [],
        "work_types": [],
        "topics": ["transformation", "outcome", "partnership", "growth",
                   "business value", "strategy", "scale", "roi"],
    },
}

# Personas that prefer high-level / narrative slides over deep technical proof.
_EXEC_PERSONAS = {"CEO_FOUNDER", "CIO_CTO", "FINANCE_HEAD"}


def _find_alias(text, aliases):
    """Return the first alias that appears as a whole phrase in text, else None.
    Short alias codes (cio/cto/cfo/coo/cdo/cpo/chro/ceo/ta) need word boundaries
    so they don't match inside other words."""
    for a in aliases:
        if re.search(r"\b" + re.escape(a) + r"\b", text):
            return a
    return None


def detect(recipient, transcript):
    """Return an ordered list of detected persona CODES.

    The 'recipient' (who we're meeting) is the strongest signal and comes first;
    personas only named in the notes/transcript come after. De-duplicated."""
    recipient = (recipient or "").lower()
    transcript = (transcript or "").lower()

    primary, secondary = [], []
    for code, p in PERSONAS.items():
        in_recipient = _find_alias(recipient, p["aliases"])
        in_notes = _find_alias(transcript, p["aliases"])
        if in_recipient:
            primary.append(code)
        elif in_notes:
            secondary.append(code)

    seen, ordered = set(), []
    for code in primary + secondary:
        if code not in seen:
            seen.add(code)
            ordered.append(code)
    return ordered


def labels(persona_codes):
    """Human-readable labels for a list of persona codes."""
    return [PERSONAS[c]["label"] for c in persona_codes if c in PERSONAS]


def _row_function(row):
    return (row.get("primary_function") or "").upper().strip()


def _row_work_types(row):
    raw = row.get("work_types") or ""
    return {w.strip().upper() for w in raw.replace("|", ",").split(",") if w.strip()}


def _row_text(row):
    """Lowercased keyword + title text for topic matching."""
    return ((row.get("keywords") or "") + " " + (row.get("title") or "")).lower()


def score_boost(persona_codes, row, max_boost=4):
    """Relevance bonus for a slide given the detected persona(s).

    +2  slide's primary_function is one this persona cares about
    +1  slide's work_type overlaps this persona's preferred work_types
    +1  a persona topic word (synonym-expanded) appears in the slide's keywords
    Capped at max_boost so persona never dominates a strong transcript match.
    Returns (boost:int, why:list[str])."""
    if not persona_codes:
        return 0, []

    fn = _row_function(row)
    wts = _row_work_types(row)
    text = _row_text(row)

    boost, why = 0, []
    for code in persona_codes:
        p = PERSONAS.get(code)
        if not p:
            continue
        local = 0
        if fn and fn in p["functions"]:
            local += 2
        if wts & set(p["work_types"]):
            local += 1
        if any(synonyms.hits_in(t, text) for t in p["topics"]):
            local += 1
        if local:
            boost += local
            why.append(p["label"])
    return min(boost, max_boost), why


def tag_slide(row):
    """Which persona CODES this slide naturally speaks to (derived).
    Used by tagger.py to persist a 'persona' tag. A slide qualifies for a
    persona if its function matches OR a persona topic word is in its keywords."""
    fn = _row_function(row)
    wts = _row_work_types(row)
    text = _row_text(row)

    out = []
    for code, p in PERSONAS.items():
        if code == "CEO_FOUNDER":
            continue   # CEO is audience-level, not slide-level; skip auto-tag
        hit = False
        if fn and fn in p["functions"]:
            hit = True
        elif any(synonyms.hits_in(t, text) for t in p["topics"]) and (
                not p["work_types"] or (wts & set(p["work_types"]))):
            hit = True
        if hit:
            out.append(code)
    return out


def prefers_overview(persona_codes):
    """True if the audience leans executive (prefer narrative over deep proof)."""
    return any(c in _EXEC_PERSONAS for c in persona_codes)


if __name__ == "__main__":
    # Quick self-check
    samples = [
        ("Head of QA", "we struggle with flaky regression and need test automation"),
        ("CFO", "main goal is cost reduction and faster invoice processing"),
        ("VP Engineering", "stand up an AI pod with strong CI/CD"),
        ("", "meeting the CHRO about RPO for 300 hires"),
        ("CTO", ""),
    ]
    for rec, tr in samples:
        codes = detect(rec, tr)
        print(f"recipient={rec!r:18} notes={tr[:40]!r:42} -> {labels(codes)}")
