# -*- coding: utf-8 -*-
"""api.py  --  JSON endpoint for the 'Create a slide with AI' button (/create_ai)."""

from flask import Blueprint, request

from deckengine.services.rendering import slide_generator, staging

bp = Blueprint("api", __name__)


@bp.route("/create_ai", methods=["POST"])
def create_ai():
    """The 'Create with AI' button: write a full structured CASE STUDY from a
    free-text brief (strict format + self-review). Stages it and returns the content
    as JSON so the page shows it inline. Added to the deck as a NEW:<id> order item;
    built into THIS deck at finalize (not promoted to the master library)."""
    brief = request.form.get("brief", "").strip()
    if not brief:
        return {"ok": False, "error": "Please describe the slide you want."}, 400
    industry = request.form.get("industry", "")
    client = request.form.get("client_name", "")
    content = slide_generator.draft_case_study(brief, {
        "industry": industry,
        "recipient": request.form.get("recipient", ""),
        "function": request.form.get("functions", ""),
        "notes": request.form.get("context", ""),
    })
    content["kind"] = "user_created"
    rec = staging.add(content, "", industry, client)
    return {"ok": True, "id": "NEW:" + rec["id"],
            "title": content["title"], "subhead": content["subhead"],
            "challenge": content["challenge"], "solution": content["solution"],
            "capabilities": content["capabilities"], "results": content["results"],
            "review": content["review"]}
