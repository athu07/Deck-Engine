# -*- coding: utf-8 -*-
"""
tagger.py  --  Box 0, Step 2 (+ Step 3): rule-based AUTO-tagging.

Takes ONE slide record (from build_library.py) and assigns its tags using
simple keyword/structure rules -- NO AI. Each tag carries a confidence flag:
  AUTO  = the machine assigned it (trust, but verify)
  HUMAN = a person has confirmed it (set later, never overwritten by AUTO)

Because tagging is driven only by the slide's own content, the SAME function
works on a brand-new slide dropped in later (Step 3) -- no manual indexing.
"""

import re

import personas  # derive which buyer-role(s) a slide speaks to

MIDDOT = "·"  # the "keyword string" separator used on the slides

# --- keyword maps: lowercase needle -> the tag it votes for -----------------

WORK_TYPE = {
    "WORKFORCE": ["workforce", "sourcing engine", "talent", "hiring", "recruit",
                  "rpo", "c2h", "greenfield", "brownfield", "odc", "bot model",
                  "coe build", "attrition", "consultant", "staffing", "pofu",
                  "talent pipeline", "contract-to-hire"],
    "AI_POD": ["ai pod", "ai-first pod", "pod model", "research wing", "devsecops",
               "ci/cd", "cd automation", "qualizen", "spcr", "test automation",
               "pentesting", "loadrunner", "performance testing", "mobile qa",
               "utaf", "gaming console", "ott"],
    "MS": ["managed service", "aiops", "agentic", "observability",
           "predictive maintenance", "forecasting", "procure-to-pay",
           "demand planning", "nx open", "cae", "cfd", "geometric deep learning",
           "rag", "equity research", "mainframe", "invoice", "carbon", "esg",
           "battery on cloud", "load balancing"],
}

INDUSTRY = {
    "HEALTHCARE": ["healthcare", "medtech", "clinical", "patient", "diagnostic",
                   "mammography", "ultrasound", "pharma", "hipaa", "fda", " ema ",
                   "health insurance", "hospital", "medical"],
    "BFSI": ["bank", "fintech", "finance", "insurance", "equity", "investment",
             "kyc", "aml", "open banking", "mainframe", "venture capital", " vc ",
             "private equity", "asset manager", "invoice", "ap/ar", "credit risk"],
    "AUTOMOTIVE": ["automotive", "oem", "adas", "vehicle", " ev ", "battery",
                   "sheet metal", "stamping", "aerodynamic", "warranty"],
    "TELECOM": ["telecom", "ott", "5g"],
    "MANUFACTURING": ["manufactur", "procurement", "procure-to-pay", "mrp",
                      "inventory", "nx open", "forming", "steel", "spend analytics",
                      "material planning"],
    "ENERGY": ["renewable", "solar", "grid", "carbon", "esg", "sustainability",
               "utilities", "o&m", "emission", "energy"],
    "TECH_IT": ["aiops", "observability", "it operations", "iso 27001",
                "infrastructure", "data center"],
    "AVIATION": ["aviation", "airline"],
}

FUNCTION = {
    "DEVOPS_CLOUD": ["devops", "devsecops", "ci/cd", "cd automation",
                     "cloud migration", "aws", "argo cd", "iac", "pipeline",
                     "github actions", "hashicorp"],
    "GCC_SETUP": ["gcc", "greenfield", "brownfield", "odc", "bot model",
                  "coe build", "capability center", "innovation hub"],
    "QUALITY_ENG_TESTING": ["testing", "test automation", "loadrunner",
                            "performance testing", "pentesting", "utaf",
                            "mobile qa", " qa ", "quality engineering"],
    "DATA_AI_PLATFORM": ["machine learning", " ml ", "rag", "nlp", "predictive",
                         "agentic", "forecasting", "deep learning", "diagnostics",
                         "multi-agent", "observability", "aiops"],
    "TALENT_ACQUISITION": ["talent acquisition", "hiring", "recruit", "rpo",
                           "c2h", "sourcing", "attrition", "talent pipeline",
                           "ca hiring"],
    "SUPPLY_CHAIN_OPS": ["supply chain", "procure", "demand planning", "inventory",
                         "load balancing", "demand response", "mrp",
                         "material planning"],
    "ENGINEERING_DESIGN": ["nx open", "catvba", "cae", "forming simulation", "cfd",
                           "omniverse", "engineering design", "stamping",
                           "aerodynamic", " cad "],
    "FINANCE_OPS": ["invoice", "ap/ar", "cash application", "equity research",
                    "deal sourcing", "mainframe", "3-way match", "sell-side"],
}


def _score(text, keyword_map):
    """Return (best_tag, votes_dict). best_tag is None if nothing matched."""
    votes = {}
    for tag, needles in keyword_map.items():
        hits = sum(1 for n in needles if n in text)
        if hits:
            votes[tag] = hits
    if not votes:
        return None, votes
    best = max(votes, key=votes.get)
    return best, votes


def _parse_keywords(subtitle):
    """If the subtitle is a middle-dot keyword string, return its parts."""
    if subtitle.count(MIDDOT) >= 2:
        return [p.strip() for p in subtitle.split(MIDDOT) if p.strip()]
    return []


def _detect_kind(title, subtitle, keywords):
    t = title.strip()
    if re.fullmatch(r"\d{1,2}", t):                 # "01" / "02" / "03"
        return "DIVIDER"
    if keywords:                                    # middle-dot keyword string
        return "CASE_STUDY"
    if "j2w delivery" in t.lower() or t.upper() == "J2W":  # section / brand dividers
        return "DIVIDER"
    return "STANDARD"


def tag_record(rec):
    """Add a 'tags' block (value + confidence per dimension) to one record."""
    title = rec.get("title", "")
    subtitle = rec.get("subtitle", "")
    # pad with spaces so " ml "/" qa "/" vc " style needles can match word-ish
    text = " " + (rec.get("full_text", "") + " " + title + " " + subtitle).lower() + " "

    keywords = _parse_keywords(subtitle)
    kind = _detect_kind(title, subtitle, keywords)

    work_type, _ = _score(text, WORK_TYPE)
    industry, _ = _score(text, INDUSTRY)
    function, _ = _score(text, FUNCTION)

    # CORE standards (cover, who-we-are, etc.) have no single work type/industry.
    if kind in ("STANDARD", "DIVIDER") and not keywords:
        # keep work_type if the slide clearly names one block, else leave blank
        if work_type is None:
            pass

    def cell(value):
        return {"value": value, "confidence": "AUTO"}

    # Persona is derived from the function/work_type/keywords this slide carries,
    # so it stays correct for brand-new slides without manual tagging.
    persona_codes = personas.tag_slide({
        "primary_function": function or "",
        "work_types": work_type or "",
        "keywords": MIDDOT.join(keywords) if keywords else "",
        "title": title,
    })

    rec["keywords"] = keywords
    rec["tags"] = {
        "kind": cell(kind),
        "work_type": cell(work_type),
        "industry": cell(industry),
        "function": cell(function),
        "persona": {"value": persona_codes, "confidence": "AUTO"},
    }
    return rec


def tag_library(records):
    return [tag_record(r) for r in records]


if __name__ == "__main__":
    import json
    import sys
    src = sys.argv[1] if len(sys.argv) > 1 else "library.json"
    out = sys.argv[2] if len(sys.argv) > 2 else "tagged_library.json"
    recs = json.load(open(src, encoding="utf-8"))
    recs = tag_library(recs)
    json.dump(recs, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"Tagged {len(recs)} records -> {out}")
