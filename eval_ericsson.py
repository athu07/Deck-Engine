# -*- coding: utf-8 -*-
"""
eval_ericsson.py  --  Automated scorecard for case-study matching quality.

Uses the real Ericsson (telecom manufacturing) example the owner curated by
hand: an email thread pasted as the meeting notes, and the 19 case studies a
human decided were the right proof points. We measure how many of those 19 the
engine surfaces on its own.

This is the pass/fail metric for the matching rebuild. Run it after every change:
    py eval_ericsson.py

It prints, for a few realistic form set-ups, how many of the 19 land in PICKS
and in PICKS+SUGGESTED. Higher is better. No files are written.
"""

import io
import sys

import matcher

# The 19 cases a human mapped to the email's solutions (the "gold" answer key).
GOLD = [
    "MSS075", "MSS003", "MSS013", "MSS076", "MSS063", "WFS028", "MSS065",
    "MSS066", "MSS067", "MSS001", "MSS064", "AIP003", "AIP015", "AIP016",
    "AIP017", "MSS002", "MSS016", "MSS032", "WFS026",
]

# A faithful stand-in for the pasted email thread — same asks, same language
# (including the salesperson's product codenames, which match no case on words).
TRANSCRIPT = """
Account: Ericsson — global telecom equipment manufacturing, plus the IT org for
the factories. This is a "what to say" thread mapping our solutions to his world.

PlanForge AI — agentic production planning. Planners burn 6-10 hours every supply
disruption stitching ERP, MRP and supplier data; we want to cut signal-to-plan lag
from days toward real time, and prove it scales to 60-80 global plants with no added
headcount.

MatrixOps — margin intelligence on the factory floor. A GenAI layer over existing
MPC/PID control that explains setpoint drift and prescribes fixes: throughput,
yield and energy improvements. Live OEE recovery across MES, QMS, WMS and ERP.

DataCrystal — a unified data control plane across multi-cloud (AWS, Azure, GCP):
security baselines, data residency and unified observability at enterprise scale.

Rosetta AI — SAP warehousing and supply-chain module enhancement with predictive
analytics and real-time dashboards to reduce lead times and clear bottlenecks.

AgentShield — certify AI agents before deployment: approval policies, human-in-the-
loop controls and full audit trails; reduce ungoverned agent spend.

TrustLayer AI — LLMOps and GenAI observability: prompt-level tracing, RAG quality
scoring, cost attribution, plus governance guardrails so every model interaction is
monitored and policy-bounded. Compliance and risk.

Orchestrate AI — vendor and partner ecosystem: procurement cycle compression,
structured RFQs, automated approval routing, and agent-to-agent orchestration.

TestMatic AI — NPI validation and product engineering: ADAS video analytics,
S/4HANA migration, hardware-in-the-loop, domain-certified test.

Telecom-specific proof: managed testing for an OTT platform launch (faster releases,
less manual effort), partner integration testing for a telecom MVNO ecosystem
(hundreds of business flows across partner APIs), and network provisioning engine
migration testing for hundreds of thousands of partners with zero revenue leakage.

Also in the manufacturing/supply-chain thread: demand-driven material planning,
IT-OT convergence security (he owns IT for global manufacturing), intelligent
inventory control, and a data & analytics team build for a manufacturing group.
"""

CORE = {"CS01", "CS02", "CS03", "CS04", "CS05", "CS06", "CS07", "CS08"}


def _cases(ids):
    return [s for s in ids if s[:3] in ("AIP", "WFS", "MSS")]


def run(label, industry, work_types, use_ai=False):
    ctx = {
        "industry": industry,
        "work_types": work_types,
        "recipient": "CIO / Head of IT",
        "transcript": TRANSCRIPT,
    }
    r = matcher.plan(ctx, use_ai=use_ai)
    picks = _cases([p["slide_id"] for p in r["picks"] if p["slide_id"] not in CORE])
    sugg = _cases([s["slide_id"] for s in r["suggested"]])
    gold = set(GOLD)
    in_picks = gold & set(picks)
    in_both = gold & (set(picks) | set(sugg))
    print(f"  {label}")
    print(f"     industry={industry!r:12} work_types={work_types}  ai={use_ai}")
    print(f"     case picks: {len(picks)}   |   gold in PICKS: {len(in_picks)}/19"
          f"   |   gold in PICKS+SUGGESTED: {len(in_both)}/19")
    missed = sorted(gold - in_both)
    if missed:
        print(f"     still missed: {missed}")
    print()
    return len(in_picks), len(in_both)


def main():
    print("=" * 74)
    print("ERICSSON MATCHING SCORECARD  (target: capture the 19 hand-picked cases)")
    print("=" * 74)
    print("--- lexical only (offline, no embeddings) ---")
    run("A) under-selected work type (telecom, AI Pods only)", "TELECOM", ["AI_POD"])
    run("B) all work types (telecom)", "TELECOM", ["AI_POD", "MS", "WORKFORCE"])
    print("--- with semantic meaning-match ON (ai=True) ---")
    run("A) under-selected work type (telecom, AI Pods only)", "TELECOM", ["AI_POD"], use_ai=True)
    run("B) all work types (telecom)", "TELECOM", ["AI_POD", "MS", "WORKFORCE"], use_ai=True)
    run("C) all work types (manufacturing)", "MANUFACTURING", ["AI_POD", "MS", "WORKFORCE"], use_ai=True)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    main()
