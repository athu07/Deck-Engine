# -*- coding: utf-8 -*-
"""api.py  --  JSON endpoint for the 'Create a slide with AI' button (/create_ai)."""

from flask import Blueprint, request

from deckengine.services.rendering import slide_generator, staging
from deckengine.services import build_context

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
