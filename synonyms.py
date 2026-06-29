# -*- coding: utf-8 -*-
"""
synonyms.py  --  Matching-fuel booster.

Problem: a slide is tagged "test automation" but the meeting notes say
"automated QA" or "we need help with our test suite". Plain word matching
misses it, so the right slide never surfaces.

Solution: equivalence GROUPS. Every term in a group is treated as the same
concept. When the matcher checks whether a slide's keyword appears in the
transcript, it also checks every synonym in that keyword's group.

This is NON-DESTRUCTIVE: it does not change the curated keyword strings in the
registry. It is a runtime expansion layer, so it ALSO helps brand-new slides
automatically — no manual re-indexing.

HOW TO EXTEND (safe for a non-developer):
  Find the group that fits, add your word to that list (lowercase). To add a
  brand-new concept, append a new [ "...", "..." ] list. Keep words lowercase.
  Words shorter than 3 characters are ignored during matching (avoids "ml"/"qa"
  matching random letters) UNLESS they appear inside a longer phrase here.
"""

import re

# Each inner list = one concept. All terms in it are considered equivalent.
SYNONYM_GROUPS = [
    # ---- Engineering / DevOps / Cloud ----
    ["ci/cd", "cicd", "ci cd", "continuous integration", "continuous delivery",
     "continuous deployment", "cd automation", "build pipeline",
     "deployment pipeline", "release pipeline", "delivery pipeline", "pipeline"],
    ["devops", "devsecops", "dev ops", "dev sec ops", "sre",
     "site reliability", "platform engineering"],
    ["cloud migration", "cloud modernization", "cloud adoption", "lift and shift",
     "re-platforming", "cloud native", "migrate to cloud"],
    ["infrastructure as code", "iac", "terraform", "hashicorp", "argo cd",
     "github actions", "gitops"],
    ["aws", "amazon web services", "azure", "gcp", "google cloud", "cloud"],

    # ---- Quality Engineering / Testing ----
    ["test automation", "automated testing", "automation testing", "qa automation",
     "quality engineering", "test engineering", "automated qa", "test suite",
     "regression automation", "qe"],
    ["performance testing", "load testing", "loadrunner", "jmeter",
     "stress testing", "scalability testing"],
    ["pentesting", "penetration testing", "security testing", "vapt",
     "vulnerability assessment", "ethical hacking"],
    ["mobile qa", "mobile testing", "app testing", "device testing"],

    # ---- Data / AI / ML ----
    ["machine learning", "ml", "deep learning", "neural network", "ml model",
     "predictive model", "model training"],
    ["generative ai", "genai", "gen ai", "llm", "large language model",
     "gpt", "foundation model", "prompt engineering"],
    ["rag", "retrieval augmented generation", "retrieval-augmented",
     "vector search", "vector database", "semantic search"],
    ["agentic", "ai agent", "ai agents", "multi-agent", "multi agent",
     "autonomous agent", "agentic ai"],
    ["nlp", "natural language processing", "text analytics", "language model"],
    ["forecasting", "demand forecasting", "predictive analytics",
     "predictive maintenance", "anomaly detection"],
    ["data platform", "data lake", "data warehouse", "lakehouse",
     "data pipeline", "data engineering", "etl"],
    ["analytics", "business intelligence", "bi", "dashboards", "reporting",
     "insights"],
    ["computer vision", "image recognition", "diagnostics", "ai diagnostics"],

    # ---- GCC / Capability / Delivery models ----
    ["gcc", "global capability center", "global capability centre",
     "capability center", "capability centre", "coe", "center of excellence",
     "centre of excellence", "odc", "offshore development center",
     "innovation hub", "delivery center"],
    ["greenfield", "brownfield", "bot model", "build operate transfer",
     "set up from scratch", "stand up a team"],
    ["ai pod", "ai-first pod", "pod model", "delivery pod", "squad model",
     "dedicated pod"],

    # ---- Talent / Workforce / Hiring ----
    ["talent acquisition", "hiring", "recruiting", "recruitment", "sourcing",
     "talent sourcing", "talent pipeline", "staffing", "headcount"],
    ["rpo", "recruitment process outsourcing", "managed hiring"],
    ["c2h", "contract to hire", "contract-to-hire", "contingent",
     "contingent workforce", "contract staffing"],
    ["attrition", "retention", "churn", "backfill"],
    ["workforce", "consultants", "deployed talent", "resource augmentation",
     "staff augmentation", "team augmentation"],

    # ---- Finance Ops ----
    ["invoice", "invoicing", "accounts payable", "accounts receivable",
     "ap/ar", "ap ar", "3-way match", "three way match", "cash application"],
    ["equity research", "deal sourcing", "sell-side", "buy-side",
     "investment research"],
    ["roi", "return on investment", "cost savings", "cost reduction",
     "efficiency", "tco", "total cost of ownership"],

    # ---- Supply chain / Ops ----
    ["supply chain", "procure to pay", "procure-to-pay", "p2p",
     "procurement", "sourcing optimization"],
    ["demand planning", "demand response", "material planning", "mrp",
     "inventory optimization", "inventory management"],

    # ---- Engineering design / CAE ----
    ["nx open", "cad", "cae", "cfd", "forming simulation", "stamping",
     "aerodynamic", "geometric deep learning", "simulation", "omniverse"],

    # ---- Managed services / Ops platforms ----
    ["managed services", "managed service", "aiops", "observability",
     "monitoring", "it operations", "noc", "l1 l2 l3 support"],

    # ---- Security / Compliance ----
    ["iso 27001", "soc 2", "compliance", "regulatory", "hipaa", "fda",
     "gdpr", "audit", "governance"],

    # ---- Industry shorthands (help industry surface in free text) ----
    ["healthcare", "medtech", "medical", "clinical", "pharma", "life sciences",
     "hospital", "patient care"],
    ["bfsi", "banking", "financial services", "fintech", "finance",
     "insurance", "capital markets"],
    ["automotive", "oem", "adas", "vehicle", "ev", "electric vehicle", "mobility"],
    ["telecom", "telecommunications", "ott", "5g", "network"],
    ["manufacturing", "industrial", "factory", "plant", "shop floor"],
    ["energy", "renewable", "solar", "grid", "utilities", "sustainability",
     "carbon", "esg", "emissions"],
]


def _build_index(groups):
    """term -> frozenset(all terms in its group). A term may appear in several
    groups; its expansion is the union of every group it belongs to."""
    idx = {}
    for grp in groups:
        gset = set(g.strip().lower() for g in grp if g.strip())
        for term in gset:
            idx.setdefault(term, set()).update(gset)
    return {k: frozenset(v) for k, v in idx.items()}


_INDEX = _build_index(SYNONYM_GROUPS)


def known_terms():
    """All surface forms known to the synonym index (for scanning a phrase)."""
    return _INDEX.keys()


def expand(term):
    """Return the set of equivalent surface forms for a term (always includes
    the term itself, lowercased). Unknown terms expand to just themselves."""
    t = (term or "").strip().lower()
    if not t:
        return frozenset()
    forms = set(_INDEX.get(t, ()))
    forms.add(t)
    return frozenset(forms)


def expand_many(terms):
    """Union-expand a list of terms."""
    out = set()
    for t in terms:
        out |= expand(t)
    return out


def hits_in(term, text):
    """True if `term` OR any of its synonyms appears as a whole word in `text`.
    `text` should already be lowercase."""
    if not text:
        return False
    for form in expand(term):
        if len(form) < 3:
            continue
        if re.search(r"\b" + re.escape(form) + r"\b", text):
            return True
    return False


if __name__ == "__main__":
    # Quick self-check
    tests = [
        ("test automation", "we want help with automated qa for our mobile app"),
        ("ci/cd", "the team needs a solid deployment pipeline"),
        ("rpo", "looking at recruitment process outsourcing for 200 hires"),
        ("gcc", "stand up a global capability centre in india"),
        ("machine learning", "exploring genai and llm use cases"),  # different concept
    ]
    for tag, txt in tests:
        print(f"{tag!r:20} in {txt!r:55} -> {hits_in(tag, txt)}")
    print(f"\nTotal indexed terms: {len(_INDEX)}")
