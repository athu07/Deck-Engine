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


def staged_item(oid, rec, industry, case_edit=None):
    """One staged record (a pasted or AI-drafted slide) -> the item shape
    `skills.build_into` renders. `oid` is the deck order id ("NEW:<staging_id>").

    Every non-case-study shape goes down the same generic path: the staged record
    IS the data dict `skills._mapping_*` / `_draw_*` read. Before this existed only
    four_box and roadmap_board were handled, so a pasted box_grid / pillar_deepdive /
    scored_list / stat_overview / data_table fell through to the case-study branch --
    it rendered as a case study AND was promoted into the library as one.

    Case studies also carry a content-store `record` (the branded case_study_v2
    template fills from it); `case_edit` overrides it with the user's /review edits.
    """
    content_type = rec.get("content_type", "case_study")
    if content_type != "case_study":
        return {"id": oid, "template": rec.get("template", content_type),
                "kind": content_type, "data": rec}
    record = ai_to_store_record(rec, industry)
    if case_edit:
        record.update(case_edit)                       # user edits win
    return {"id": oid, "template": "case_study_v2", "kind": "case_study", "record": record}


def promote_case(stg_id, rec, record, fallback_work_type, industry):
    """Auto-save an accepted case study into the shared library (content-store JSON +
    the source Excel + an embedding), so it's searchable/reusable on future projects.

    Idempotent: `promote_ai_case` mints a fresh id every call, so a slide that is both
    downloaded from the Custom Slide Builder AND folded into a deck would otherwise be
    saved twice. The new id is stamped onto the staging record; a record that already
    carries one is skipped.

    The record's OWN work type wins (the builder requires one per slide); the deck's
    first work type is only a fallback for a mixed-work-type deck. Best-effort by
    design -- a library-save failure never blocks the deck.
    """
    if rec.get("promoted_id"):
        return rec["promoted_id"]
    work_type = (rec.get("work_type") or "").strip() or fallback_work_type
    try:
        new_id = case_library.promote_ai_case(record, work_type, industry)
    except Exception:
        return None
    if new_id:
        staging.mark_promoted(stg_id, new_id)
    return new_id


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
            stg_id = oid[4:]
            rec = staging.get(stg_id)
            if not rec:
                continue
            item = staged_item(oid, rec, industry, case_edits.get(oid))
            create_items.append(item)
            # Only case studies reach the library -- case_library's schema is
            # case-study-shaped, so the other seven shapes have nowhere to go.
            if item["kind"] == "case_study":
                promote_case(stg_id, rec, item["record"], first_wt, industry)
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
