# -*- coding: utf-8 -*-
"""
builder.py  --  the Custom Slide Builder (/builder).

Separates CONTENT SOURCING (the salesperson supplies it) from SLIDE DESIGN (we apply
the branded template). The MS team hands over a whole deck's worth of content in one
document; this is where it becomes real, branded slides without going anywhere near
the matcher or the AI case-study writer.

Three steps, and the middle one is the whole point:

    POST /builder/parse   split the document into slides, match each label to a
                          template, flag library duplicates. FREE: pure text plus one
                          embedding call. Nothing is generated.
    (the review screen)   the salesperson fixes a shape, drops a duplicate, or goes
                          back and fixes the text -- before anything is paid for.
    POST /builder/slide   build ONE slide. The browser fires these four at a time and
                          fills each card as its render lands, so ten slides is one
                          wait with visible progress, not a spinner.

Getting the split wrong and discovering it after a 40-second build is the failure the
review screen exists to prevent.

Queue state lives in the browser (localStorage `j2w_content_queue`) -- the same key
the new-deck form reads -- so the two exits both work with no server-side session:
    * Download these slides   -> GET /builder/download?ids=...
    * Use in a deck           -> back to /new, queue folds into the generated deck

Case studies reach the shared library ON COMMIT (download or deck), never merely on
build -- a slide you build and then delete never pollutes the library. See
deck_build.promote_case for the idempotency that lets you do BOTH without saving twice.

NOTE: /builder/slide is called CONCURRENTLY (four at a time), so anything it touches
must be thread-safe -- staging.add() is serialised behind a lock for exactly this.
"""

import os

from flask import Blueprint, render_template, request, jsonify, send_file, abort
from pptx import Presentation

from deckengine import config
from deckengine import constants
from deckengine.constants import WT_LABELS
from deckengine.services import ingest
from deckengine.services.content import case_library, paste_parser
from deckengine.services.matching import dedupe
from deckengine.services.rendering import (deck_build, preview, skills, slide_generator,
                                           slide_schema, staging)
from .view_helpers import shell, file_busy_page, safe_filename

bp = Blueprint("builder", __name__)

_WORK_TYPES = ("WORKFORCE", "AI_POD", "MS")


def _pasted_content(form, files):
    """The pasted text plus any attached document, as one string."""
    content = (form.get("content", "") or "").strip()
    ftext = ingest.extract_text(files.get("content_file"))
    if ftext:
        content = (content + "\n\n" + ftext).strip()
    return content


@bp.route("/builder")
def builder_page():
    body = render_template("builder.html",
                           content_templates=slide_generator.CONTENT_TEMPLATES,
                           work_types=_WORK_TYPES, wt_labels=WT_LABELS,
                           industries=constants.all_industries(),
                           threshold_pct=int(dedupe.THRESHOLD * 100))
    return shell(body, active="builder", crumb="<b>Custom Slide Builder</b>")


@bp.route("/builder/parse", methods=["POST"])
def builder_parse():
    """Cut the pasted document into slides and report the split -- WITHOUT building
    anything (owner's spec, 2026-07-10). Pure text plus one embedding call, so the
    salesperson sees the split, the shape matched to each label, and any library
    duplicate BEFORE a single AI call or slide render is paid for. Getting the split
    wrong and finding out after a 40-second build is the failure this prevents."""
    content = _pasted_content(request.form, request.files)
    if not content:
        return jsonify({"ok": False, "error": "Paste the content or attach a document first."})
    work_type = (request.form.get("work_type", "") or "").strip().upper()
    if work_type not in _WORK_TYPES:
        return jsonify({"ok": False, "error": "Choose a work type — it's needed to save "
                                              "these into the shared library."})
    slides = paste_parser.parse(content)
    if not slides:
        return jsonify({"ok": False, "error": "Couldn't find any slide content in that."})

    # THE INTELLIGENCE LAYER (owner's spec, 2026-07-10). The salesperson often doesn't
    # know what shape a slide is -- they write a heading ("How we think before we build")
    # or nothing. Where they DID name a category, honour it untouched. Where they didn't,
    # read the content and map it onto the categories we actually have. ONE call for the
    # whole document, so ten unlabelled slides cost one round-trip, and the model sees
    # the slides side by side (a deck's slides inform each other's shapes).
    unknown = [i for i, s in enumerate(slides) if s["template"] == paste_parser.AUTO]
    if unknown:
        read = slide_generator.classify_content_many(
            [{"heading": slides[i]["heading"], "content": slides[i]["content"]}
             for i in unknown])
        for i, key in zip(unknown, read):
            slides[i]["template"] = key
            slides[i]["inferred"] = True      # the review screen says so, and offers a change

    # one embedding call for the whole document, not one per slide
    matches = dedupe.similar_cases_many([s["content"] for s in slides],
                                        allowed_work_types=[work_type])
    for slide, hits in zip(slides, matches):
        # The library holds ONLY case studies, so a match can only ever be a real
        # duplicate of a case-study-shaped slide -- and by now every slide HAS a shape.
        if slide["template"] != "case_study":
            hits = []
        for m in hits:
            m["percent"] = int(round(m["score"] * 100))
            m["wt_label"] = WT_LABELS.get(m["work_type"], m["work_type"])
        slide["matches"] = hits
    return jsonify({"ok": True, "slides": slides,
                    "templates": [{"key": t["key"], "label": t["label"]}
                                  for t in slide_generator.CONTENT_TEMPLATES]})


def _editable_card(staged):
    """The staged slide rendered as the EDITABLE SLIDE CARD the review page shows --
    click straight onto its title, paragraph, capability. Same macro, same markup, same
    CSS (templates/_slide_editor.html). This IS the builder's preview: a picture of a
    slide is something you look at; this is something you fix."""
    kind = staged.get("content_type", "case_study")
    return render_template("builder_editor.html", kind=kind,
                           s=slide_schema.view_model(staged),
                           schema=slide_schema.fields_for(kind))


@bp.route("/builder/slide", methods=["POST"])
def builder_slide():
    """Build ONE slide from the pasted content, stage it, and return it as an editable
    slide card -- exactly what /review shows for a deck built from scratch."""
    content = _pasted_content(request.form, request.files)
    work_type = (request.form.get("work_type", "") or "").strip().upper()
    # Work type is REQUIRED, not merely nice-to-have: it decides the id prefix when the
    # case study is saved to the library. A blank one used to sail through here and
    # silently fail to save at finalize (owner-reported, 2026-07-08).
    if not content:
        return jsonify({"ok": False, "error": "Paste the content or attach a document first."})
    if work_type not in _WORK_TYPES:
        return jsonify({"ok": False, "error": "Choose a work type — it's needed to save "
                                              "this into the shared library."})
    industry = request.form.get("industry", "")
    client = request.form.get("client_name", "").strip()
    rec, tdef = slide_generator.build_content_slide(
        content, industry, request.form.get("template_hint", "auto"))
    staged = staging.add(rec, work_type, industry, client)
    return jsonify({"ok": True, "id": staged["id"], "title": staged.get("title", "Untitled"),
                    "content_type": staged.get("content_type", "case_study"),
                    "template_label": tdef["label"],
                    "html": _editable_card(staged)})


@bp.route("/builder/slide/<sid>", methods=["POST"])
def builder_slide_save(sid):
    """Save the edits typed straight onto one slide card.

    Edits go through the shape's own normalizer (slide_schema.apply_edits), so a
    hand-edit can't break an invariant the renderer relies on -- e.g. a case study must
    still have exactly 6 capabilities.

    The staged record is the single source of truth: downloading, folding into a deck,
    and saving to the library all read it, so an edit made here follows the slide
    everywhere without being threaded through any of those paths."""
    rec = staging.get(sid)
    if not rec:
        abort(404)
    edits = request.get_json(silent=True) or {}
    fields = slide_schema.apply_edits(rec, edits, request.args.get("industry", ""))
    updated = staging.update_fields(sid, fields)
    if not updated:
        abort(404)
    return jsonify({"ok": True, "id": sid, "title": updated.get("title", "")})


@bp.route("/builder/preview/<sid>")
def builder_preview(sid):
    """Render an existing library case (AIP/WFS/MSS) so a near-duplicate can be SEEN
    before it's reused, in the same preview area a freshly-built slide would use."""
    sid = sid.upper()
    rec = case_library.record(sid)
    if not rec:
        abort(404)
    return jsonify({"ok": True, "id": sid, "title": rec.get("title", ""),
                    "content_type": "case_study", "png": preview.case_png(sid)})


def _queue_items(ids, industry):
    """The queued ids -> the items skills.build_into renders, in queue order.
    An id is either a content-store case (reused from the duplicate check) or a
    staging id (a slide built here). Returns (order, items, staged_pairs)."""
    order, items, staged_pairs = [], [], []
    for raw in ids:
        rec = case_library.record(raw.upper())
        if rec:                                   # a reused library case
            order.append(raw.upper())
            items.append({"id": raw.upper(), "template": "case_study_v2",
                          "kind": "case_study", "record": rec})
            continue
        stg = staging.get(raw)
        if not stg:
            continue
        oid = "NEW:" + raw
        order.append(oid)
        items.append(deck_build.staged_item(oid, stg, industry or stg.get("industry", "")))
        staged_pairs.append((raw, stg))
    return order, items, staged_pairs


@bp.route("/builder/download")
def builder_download():
    """Assemble the queue into one .pptx of just these slides, in queue order --
    no title/closing bookends; this is a handful of slides, not a pitch deck.

    This is also a COMMIT point: every case study in the queue is saved to the shared
    library here (JSON + the source Excel + an embedding), so a future build can match
    it instead of writing it again. Idempotent -- folding the same slide into a deck
    later won't save it twice."""
    ids = [x.strip() for x in request.args.get("ids", "").split(",") if x.strip()]
    if not ids:
        abort(400, "Nothing queued to download.")
    industry = request.args.get("industry", "")
    client = request.args.get("client_name", "").strip()
    order, items, staged_pairs = _queue_items(ids, industry)
    if not items:
        abort(400, "None of those slides could be found.")

    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    fname = "J2W_Slides_" + safe_filename(client or "Custom") + ".pptx"
    path = os.path.join(config.OUTPUT_DIR, fname)

    # skills.build_into copies each template slide into a real deck on disk, so start
    # from an empty presentation sized to match the templates (13.33in x 7.5in).
    tpl = Presentation(config.SKILLS_TEMPLATES_PPTX)
    blank = Presentation()
    blank.slide_width, blank.slide_height = tpl.slide_width, tpl.slide_height
    try:
        blank.save(path)
        skills.build_into(path, order, items)
    except (PermissionError, OSError) as e:
        return file_busy_page(str(e))

    # commit: the case studies in this queue join the shared library
    for stg_id, stg in staged_pairs:
        item = next((i for i in items if i["id"] == "NEW:" + stg_id), None)
        if item and item["kind"] == "case_study":
            deck_build.promote_case(stg_id, stg, item["record"],
                                    stg.get("work_type", ""), industry or stg.get("industry", ""))

    return send_file(path, as_attachment=True, download_name=fname)
