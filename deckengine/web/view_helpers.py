# -*- coding: utf-8 -*-
"""
view_helpers.py  --  shared helpers for the web blueprints.

`shell()` wraps a page body in the site chrome; the rest are small view-support
helpers (page-safe filename, the file-busy error page, the current salesperson
stub, and the two catalog/stat helpers the pages need).
"""

import glob
import json
import os
import re
from collections import Counter

from flask import render_template

from deckengine import config
from deckengine import constants
from deckengine.services.content import industries as custom_industries
from deckengine.services.matching import matcher
from deckengine.services.rendering import slide_generator


def shell(body, active="home", crumb="<b>Home</b> / Overview", title="J2W Pre-sales Engine", tabs=None):
    return render_template("_shell.html", body=body, active=active, crumb=crumb, title=title, tabs=tabs)


def file_busy_page(err):
    """Friendly page when the output .pptx can't be written (open/locked/syncing)."""
    body = ("<div class='card' style='border-left:5px solid #c0392b;background:#fdecea;"
            "color:#8a2a1e'><h2 class='sec-title' style='color:#8a2a1e'>Couldn't save the deck</h2>"
            "<p style='margin:8px 0 0'>%s</p>"
            "<p style='margin:10px 0 0;font-size:13px'>Tip: if a previous version of this "
            "deck is open in PowerPoint, close it, then go <a href='javascript:history.back()'>back</a> "
            "and click Download again.</p></div>" % err)
    return shell(body, active="new", crumb="<b>New deck</b> / Preview")


def safe_filename(name):
    return re.sub(r"[^A-Za-z0-9_-]+", "_", name).strip("_") or "Client"


def resolve_industry(form):
    """The industry the salesperson ACTUALLY meant, from a New-deck form payload.

    Picking "Other…" makes the select's value the literal `__OTHER__` sentinel; the
    real, typed industry rides in the companion `industry_other` field. Every endpoint
    that reads the form's industry must go through here, because the sentinel is NOT
    an industry name -- "Research this account" used to POST it straight through to
    deep_research.strategic_brief(), so the brief was researched for an account in the
    "(__OTHER__)" industry and the whole deck was then matched and AI-written from that
    junk brief (owner-reported, 2026-07-17: a typed "Other" industry was ignored by the
    generated content). Server-side so a stale cached new-form.js, a no-JS submit, or a
    future caller can never reintroduce it.
    """
    industry = (form.get("industry") or "").strip()
    if industry == constants.INDUSTRY_OTHER:
        industry = (form.get("industry_other") or "").strip()
    return industry


def remember_custom_industry(industry):
    """Persist a salesperson-typed industry that isn't in the built-in taxonomy, so it
    is in the dropdown for every build after this one (owner's spec: an industry typed
    once must be pickable the second time). No-op for a built-in code or an empty name.
    Shared by /build and /research_account -- whichever moment she first commits the
    industry is the moment it should be remembered, not only a completed deck."""
    if industry and industry.lower() not in {i.lower() for i in constants.all_industries()}:
        custom_industries.add(industry)


def current_salesperson():
    """Who generated the deck. No login exists yet, so return a clearly-marked
    placeholder. At deploy time, wire this to the logged-in user."""
    return "[NOT LOGGED IN - wire to login at deploy]"


_LEGACY_CASE_IDS = None


def legacy_case_ids():
    """The 105 legacy CASE_STUDY slides in the master deck (old template). They're
    superseded by the content-store cases (which render from the new branded
    template), so they must NOT appear in the add-slide picker — otherwise a manual
    add pulls the OLD-template version instead of the new branded one."""
    global _LEGACY_CASE_IDS
    if _LEGACY_CASE_IDS is None:
        try:
            _LEGACY_CASE_IDS = {r["slide_id"] for r in matcher.load_registry()
                                if (r.get("kind") or "").strip().upper() == "CASE_STUDY"}
        except Exception:
            _LEGACY_CASE_IDS = set()
    return _LEGACY_CASE_IDS


def dash_stats():
    """Aggregate library / deck stats for the dashboard page."""
    try:
        recs = json.load(open(config.TAGGED_LIBRARY_JSON, encoding="utf-8"))
    except Exception:
        recs = []
    total = len(recs)
    cases = sum(1 for r in recs if r.get("tags", {}).get("kind", {}).get("value") == "CASE_STUDY")
    wt, ind = Counter(), Counter()
    for r in recs:
        t = r.get("tags", {})
        if t.get("work_type", {}).get("value"):
            wt[t["work_type"]["value"]] += 1
        if t.get("industry", {}).get("value"):
            ind[t["industry"]["value"]] += 1
    try:
        templates = list(slide_generator.list_templates().keys())
    except Exception:
        templates = []
    decks = []
    for p in sorted(glob.glob(os.path.join(constants.OUTPUT_DIR, "*.pptx")),
                    key=os.path.getmtime, reverse=True)[:6]:
        nm = os.path.basename(p)
        decks.append({"name": nm.replace("Tailored_Deck_", "").replace(".pptx", "").replace("_", " "),
                      "size": "%.1f MB" % (os.path.getsize(p) / 1048576)})
    wt_items = [("Workforce", wt.get("WORKFORCE", 0)), ("AI Pods", wt.get("AI_POD", 0)),
                ("Managed services", wt.get("MS", 0))]
    wt_max = max([c for _, c in wt_items] + [1])
    ind_items = ind.most_common(5)
    ind_max = max([c for _, c in ind_items] + [1])
    return dict(total=total, cases=cases, templates=templates, decks=decks,
                wt_items=wt_items, wt_max=wt_max, ind_items=ind_items, ind_max=ind_max)
