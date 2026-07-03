# -*- coding: utf-8 -*-
"""
deck_build.py  --  assemble the final tailored deck (the core of /finalize).

Given the ordered id list and the client context, this builds the master `CSxx`
slides, applies inline edits + client-token replacement, and renders the content-
store case studies / AI-drafted slides into the deck. Kept out of the route so the
web layer only orchestrates (promote AI slides, log the meeting, render the page).
"""

import re

from deckengine.services.rendering import assembler
from deckengine.services.rendering import skills
from deckengine.services.rendering import staging
from deckengine.services.content import editor
from deckengine.services.content.content_store import content_store


def ai_to_store_record(content, industry_code):
    """Reshape an AI-drafted case study into the content-store record shape so it
    renders from the branded case_study_v2 template, identical to library cases."""
    sub = content.get("subhead", "") or ""
    fm = re.search(r"Function:\s*([^|]+)", sub)
    dm = re.search(r"Domain:\s*([^|]+)", sub)
    # the account's own industry is authoritative for the domain (the AI's subhead
    # sometimes drops the function in the Domain slot)
    domain = (industry_code or "").replace("_", " ").title() or (dm.group(1).strip() if dm else "")
    return {
        "id": content.get("id", ""),
        "title": content.get("title", "Proposed Case Study"),
        "domain": domain,
        "industry": (industry_code or "").upper(),
        "function": (fm.group(1).strip() if fm else ""),
        "challenge": content.get("challenge", ""),
        "solution": content.get("solution", ""),
        "capabilities": content.get("capabilities", []),   # "Name: line" -> split_capability
        "results": content.get("results", []),
    }


def assemble(final_ids, path, *, client, industry, work_types, transcript, edits, case_edits=None):
    """Build the deck at `path` from `final_ids`, applying edits/tokens and rendering
    skills slides, content-store cases, and any AI-drafted (NEW:) cases. `case_edits`
    maps a case id -> the fields the user changed on the review slide-view (title,
    domain, function, challenge, solution, capabilities, results); they override the
    stored record so the built slide reflects the edits."""
    case_edits = case_edits or {}
    # "Create with AI" slides ride as NEW:<staging_id>; build them from the staged
    # case-study content into THIS deck (not promoted to the master library).
    create_items = []
    for oid in final_ids:
        if oid.startswith("NEW:"):
            rec = staging.get(oid[4:])
            if rec:
                record = ai_to_store_record(rec, industry)
                if case_edits.get(oid):
                    record.update(case_edits[oid])     # user edits win
                create_items.append({"id": oid, "template": "case_study_v2", "record": record})
    # Content-store case studies (AIP/WFS/MSS ids) -> rendered fresh from the shared
    # case_study_v2 template, anonymised + dash-clean, into THIS deck.
    store_recs = content_store()
    store_items = []
    for oid in final_ids:
        if oid in store_recs:
            record = dict(store_recs[oid])
            if case_edits.get(oid):
                record.update(case_edits[oid])         # user edits win
            store_items.append({"id": oid, "template": "case_study_v2", "record": record})
    # Skills slides ride along in final_ids (SK:/FP: ids); re-derive their data here.
    skills_cands = skills.candidates({"work_types": work_types, "industry": industry,
                                      "transcript": transcript, "client_name": client})

    assembler.build_deck(final_ids, out=path)     # builds the CS slides (skills/NEW ids ignored)
    if edits:
        editor.apply_edits(path, edits)
    if client:
        editor.replace_tokens(path, {"[CLIENT]": client, "[Client]": client, "[client]": client})
    skills.build_into(path, final_ids, skills_cands + create_items + store_items)   # fill + slot extras
