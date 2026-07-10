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
from deckengine.services.rendering import client_context
from deckengine.services.rendering import client_logo
from deckengine.services.content import editor
from deckengine.services.content import case_library
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


def assemble(final_ids, path, *, client, industry, work_types, transcript, edits,
             case_edits=None, phase=""):
    """Build the deck at `path` from `final_ids`, applying edits/tokens and rendering
    skills slides, content-store cases, and any AI-drafted (NEW:) cases. `case_edits`
    maps a case id -> the fields the user changed on the review slide-view (title,
    domain, function, challenge, solution, capabilities, results); they override the
    stored record so the built slide reflects the edits. `phase` drives the Client
    Context / Tailored Approach pair (Workforce, First/Second stage only)."""
    case_edits = case_edits or {}
    # "Create with AI" slides ride as NEW:<staging_id>; build them into THIS deck,
    # AND auto-save each one into the shared content store (owner's choice:
    # automatic on every accept, no extra step) so it's searchable/reusable on
    # future projects. Best-effort by design (case_library.promote_ai_case never
    # raises) -- a save failure never blocks the deck itself.
    first_wt = (work_types or [None])[0]   # best-effort id-prefix for a mixed deck
    create_items = []
    for oid in final_ids:
        if oid.startswith("NEW:"):
            rec = staging.get(oid[4:])
            if not rec:
                continue
            if rec.get("content_type") == "four_box":
                # a 4-way breakdown, not a case study (owner's spec, 2026-07-08) --
                # goes through skills.build_into's generic marker-fill path
                # instead of the case-study-specific one; no content-store save,
                # since case_library's schema is case-study-shaped, not this.
                create_items.append({"id": oid, "template": "four_box", "kind": "four_box",
                                     "data": {"title": rec.get("title", ""),
                                             "subhead": rec.get("subhead", ""),
                                             "boxes": rec.get("boxes", [])}})
                continue
            if rec.get("content_type") == "roadmap_board":
                # a variable-column phased roadmap/board, not a case study --
                # drawn programmatically (skills.build_into's roadmap_board
                # branch), same no-content-store-save reasoning as four_box.
                create_items.append({"id": oid, "template": "roadmap_board", "kind": "roadmap_board",
                                     "data": {"title": rec.get("title", ""),
                                             "subhead": rec.get("subhead", ""),
                                             "intro": rec.get("intro", ""),
                                             "columns": rec.get("columns", []),
                                             "legend": rec.get("legend", []),
                                             "footer_title": rec.get("footer_title", ""),
                                             "footer_body": rec.get("footer_body", "")}})
                continue
            record = ai_to_store_record(rec, industry)
            if case_edits.get(oid):
                record.update(case_edits[oid])     # user edits win
            create_items.append({"id": oid, "template": "case_study_v2", "record": record})
            try:
                case_library.promote_ai_case(record, first_wt, industry)
            except Exception:
                pass    # never let a library-save failure block the deck
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
    # Client Context / Tailored Approach (CS65/CS66) -- real master-deck ids, but
    # gated + AI-filled here rather than by the registry (see client_context.py).
    cc_cands = client_context.candidates({"work_types": work_types, "phase": phase,
                                          "transcript": transcript})

    assembler.build_deck(final_ids, out=path)     # builds the CS slides (skills/NEW ids ignored)
    if edits:
        editor.apply_edits(path, edits)
    if cc_cands:
        client_context.fill_into(path, final_ids, cc_cands)
    if client:
        # bracket-style tokens (legacy convention) + curly-style, single AND double
        # brace (the new 39-slide master's slide 26 mixes "{{CLIENT}}", "{{client}}"
        # and a single-brace "{client}") -- same run-level replace either way
        editor.replace_tokens(path, {
            "[CLIENT]": client, "[Client]": client, "[client]": client,
            "{{CLIENT}}": client, "{{Client}}": client, "{{client}}": client,
            "{CLIENT}": client, "{Client}": client, "{client}": client,
        })
        client_logo.stamp_into(path, client)   # no-op if no logo was uploaded for this client
    skills.build_into(path, final_ids, skills_cands + create_items + store_items)   # fill + slot extras
