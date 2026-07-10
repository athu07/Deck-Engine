# -*- coding: utf-8 -*-
"""api.py  --  JSON endpoints: 'Create a slide with AI' (/create_ai), 'Research
this account' (/research_account), and 'Find logo' (/logo_preview)."""

import base64

from flask import Blueprint, request

from deckengine.services.rendering import slide_generator, staging, client_logo
from deckengine.services import build_context, deep_research, ingest

bp = Blueprint("api", __name__)


@bp.route("/create_ai", methods=["POST"])
def create_ai():
    """The 'Create with AI' button: write a full structured CASE STUDY from a
    free-text brief (strict format + self-review). Stages it and returns the content
    as JSON so the page shows it inline. Added to the deck as a NEW:<id> order item;
    built into THIS deck at finalize (not promoted to the master library).

    The rich build context (deep research + profile + FULL transcript) is reloaded
    from build_context by build_id, so the case study is grounded in the real client
    context; falls back to the browser-sent fields if that context is unavailable."""
    brief = request.form.get("brief", "").strip()
    if not brief:
        return {"ok": False, "error": "Please describe the slide you want."}, 400
    bc = build_context.load(request.form.get("build_id", ""))
    industry = bc.get("industry") or request.form.get("industry", "")
    client = bc.get("client_name") or request.form.get("client_name", "")
    recipient = bc.get("recipient") or request.form.get("recipient", "")
    functions = bc.get("functions")
    functions = ", ".join(functions) if isinstance(functions, list) else (functions or request.form.get("functions", ""))
    transcript = bc.get("transcript") or request.form.get("context", "")   # FULL transcript from the store
    content = slide_generator.draft_case_study(brief, {
        "industry": industry,
        "recipient": recipient,
        "function": functions,
        "notes": transcript,
        "research": bc.get("research", ""),   # deep-research brief — now reaches generation
        "profile": bc.get("profile", ""),     # stakeholder profile — now reaches generation
    })
    content["kind"] = "user_created"
    rec = staging.add(content, "", industry, client)
    return {"ok": True, "id": "NEW:" + rec["id"],
            "title": content["title"], "subhead": content["subhead"],
            "challenge": content["challenge"], "solution": content["solution"],
            "capabilities": content["capabilities"], "results": content["results"],
            "review": content["review"]}


@bp.route("/research_account", methods=["POST"])
def research_account():
    """'Research this account' button on the new-deck form: a live, executive-
    grade strategic account brief (deep_research.strategic_brief -- the
    owner's own consulting-style prompt), ADDITIVE to the existing manual
    research/profile file upload -- returned as JSON for the salesperson to
    review/edit before /build ever sees it (never silently applied). Needs a
    company name (client_name); the stakeholder (recipient) is optional and
    sharpens the brief when given. If a research/profile file is ALREADY
    attached on the same form, its content is read here too (owner's spec,
    2026-07-09) and passed in as grounding -- the live research builds on
    what's already on file for this account, not just the typed name."""
    client_name = request.form.get("client_name", "").strip()
    recipient = request.form.get("recipient", "").strip()
    industry = request.form.get("industry", "").strip()
    if not client_name:
        return {"ok": False, "error": "Enter a client name first."}, 400
    research_text = ingest.extract_text(request.files.get("research_file"))
    profile_text = ingest.extract_text(request.files.get("profile_file"))
    brief = deep_research.strategic_brief(client_name, recipient, industry,
                                          profile_text=profile_text, research_text=research_text)
    if not brief:
        return {"ok": False, "error": "Couldn't find anything reliable for that "
                                       "account — try pasting notes or a research file instead."}
    return {"ok": True, "brief": brief}


@bp.route("/logo_preview", methods=["POST"])
def logo_preview():
    """'Find logo' button: best-effort auto-fetch (client_logo.fetch_via_search),
    returned as a data URI so the page can show it BEFORE it's used anywhere --
    auto-fetch is real but not reliable (see client_logo.py's module docstring
    for what was actually tested), so this is a preview-and-confirm step, not a
    silent apply. If confirmed, the SAME data URI rides through to /build in a
    hidden field, so the deck gets exactly what was previewed, not a re-rolled
    search result."""
    client_name = request.form.get("client_name", "").strip()
    domain = request.form.get("domain", "").strip()
    if not client_name:
        return {"ok": False, "error": "Enter a client name first."}, 400
    data = client_logo.fetch_via_search(client_name, domain)
    if not data:
        return {"ok": False, "error": "Couldn't find a usable logo automatically — "
                                       "try uploading the file instead."}
    return {"ok": True, "data_uri": "data:image/png;base64," + base64.b64encode(data).decode("ascii")}
