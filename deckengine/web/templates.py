# -*- coding: utf-8 -*-
"""templates.py  --  the /templates page: fill-on-demand templates, saved decks
(upload & reuse), and the deck re-skinner (upload -> J2W-branded -> download)."""

import io
import os
import re
import shutil
import time
import uuid

from flask import Blueprint, render_template, request, send_file, redirect, abort
from pptx import Presentation

from deckengine import config
from deckengine.services.rendering import slide_generator, reskin, templatize
from deckengine.services import saved_templates
from .view_helpers import shell

bp = Blueprint("templates", __name__)

# temp holding area for an uploaded deck while the owner picks a slide + reviews
# the AI-proposed role mapping (the "learn a template" multi-step flow) -- keyed
# by a random token threaded through the form, cleaned up once saved/cancelled.
_PENDING_DIR = os.path.join(config.LEARNED_TEMPLATES_DIR, "_pending")


@bp.route("/templates")
def templates_page():
    items = []
    try:
        for name, slide in slide_generator.list_templates().items():
            markers = set()
            text = ""
            for sh in slide.shapes:
                if sh.has_text_frame:
                    markers.update(re.findall(r"\{\{[A-Z]+\}\}", sh.text_frame.text))
                    text += sh.text_frame.text
            status = "placeholder" if "Generated slide" in text else "active"
            items.append({"name": name, "markers": sorted(markers), "status": status})
    except Exception:
        items = []
    body = render_template("templates_page.html", items=items,
                           saved=saved_templates.all_templates(),
                           learned=templatize.all_templates())
    return shell(body, active="templates", crumb="<b>Templates</b>")


# ── F3: save an uploaded deck as a reusable template ──────────────────────────
@bp.route("/template/save", methods=["POST"])
def template_save():
    f = request.files.get("deck_file")
    if not f or not getattr(f, "filename", ""):
        return redirect("/templates")
    data = f.read()
    if not data:
        return redirect("/templates")
    name = request.form.get("name", "").strip() or os.path.splitext(f.filename)[0]
    try:
        slides = len(Presentation(io.BytesIO(data)).slides)
    except Exception:
        slides = 0
    saved_templates.save(name, data, slides)
    return redirect("/templates")


@bp.route("/template/saved/<tid>/download")
def template_download(tid):
    p = saved_templates.file_path(tid)
    if not p:
        abort(404)
    row = saved_templates.get(tid)
    return send_file(p, as_attachment=True,
                     download_name=(row["name"] + ".pptx") if row else "template.pptx")


@bp.route("/template/saved/<tid>/delete", methods=["POST"])
def template_delete(tid):
    saved_templates.delete(tid)
    return redirect("/templates")


# ── F2: re-skin an uploaded deck into the J2W design ──────────────────────────
def _slug(s):
    return re.sub(r"[^A-Za-z0-9_-]+", "_", s or "").strip("_") or "Deck"


@bp.route("/template/reskin", methods=["POST"])
def template_reskin():
    f = request.files.get("deck_file")
    if not f or not getattr(f, "filename", ""):
        return redirect("/templates")
    data = f.read()
    try:                                          # reject non-pptx up front
        Presentation(io.BytesIO(data))
    except Exception:
        body = render_template("templates_page.html", items=[],
                               saved=saved_templates.all_templates(),
                               reskin_error="Couldn't read that file — please upload a valid .pptx.")
        return shell(body, active="templates", crumb="<b>Templates</b>")
    # rebrand the original deck in place (nothing dropped), then render true previews
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    fname = "Reskinned_" + _slug(os.path.splitext(f.filename)[0]) + ".pptx"
    out = os.path.join(config.OUTPUT_DIR, fname)
    try:
        reskin.restyle_deck(data, out)
    except (PermissionError, OSError) as e:
        from .view_helpers import file_busy_page
        return file_busy_page(str(e))
    rdir = os.path.join(config.RENDERS_DIR, "reskin")
    shutil.rmtree(rdir, ignore_errors=True)
    try:
        pngs = reskin.render_pngs(out, rdir)
    except Exception:
        pngs = []
    images = ["/static/renders/reskin/" + os.path.basename(p) for p in pngs]
    nslides = len(Presentation(out).slides._sldIdLst)
    body = render_template("reskin_preview.html", images=images, nslides=nslides,
                           filename=fname, src_name=f.filename)
    return shell(body, active="templates", crumb="<b>Templates</b> / Re-skin preview")


@bp.route("/template/save_output", methods=["POST"])
def template_save_output():
    """Save an already-built output deck (e.g. a re-skinned one) as a reusable template."""
    fname = os.path.basename(request.form.get("filename", ""))
    p = os.path.join(config.OUTPUT_DIR, fname)
    if fname and os.path.exists(p):
        with open(p, "rb") as fh:
            data = fh.read()
        name = request.form.get("name", "").strip() or os.path.splitext(fname)[0].replace("Reskinned_", "")
        try:
            slides = len(Presentation(io.BytesIO(data)).slides)
        except Exception:
            slides = 0
        saved_templates.save(name, data, slides)
    return redirect("/templates")


# ── F4: "learn a template" -- upload -> pick a slide -> AI-propose roles ──────
#       -> owner confirms/fixes -> save. See templatize.py for the full design
#       rationale (why this needs a human-confirm step before it's ever used).
def _pending_path(token):
    return os.path.join(_PENDING_DIR, token + ".pptx")


def _sweep_stale_pending(max_age_seconds=6 * 3600):
    """An abandoned upload (picked a slide, never saved) leaves a temp file
    behind forever otherwise -- swept opportunistically on the next upload
    rather than a scheduled job, since this is a low-traffic, low-stakes flow."""
    try:
        if not os.path.isdir(_PENDING_DIR):
            return
        now = time.time()
        for fn in os.listdir(_PENDING_DIR):
            p = os.path.join(_PENDING_DIR, fn)
            if now - os.path.getmtime(p) > max_age_seconds:
                os.remove(p)
    except OSError:
        pass


@bp.route("/templatize/upload", methods=["POST"])
def templatize_upload():
    f = request.files.get("deck_file")
    if not f or not getattr(f, "filename", ""):
        return redirect("/templates")
    data = f.read()
    try:
        Presentation(io.BytesIO(data))   # reject non-pptx up front
    except Exception:
        return redirect("/templates")
    _sweep_stale_pending()
    os.makedirs(_PENDING_DIR, exist_ok=True)
    token = uuid.uuid4().hex
    with open(_pending_path(token), "wb") as out:
        out.write(data)
    return redirect(f"/templatize/pick/{token}")


@bp.route("/templatize/pick/<token>")
def templatize_pick(token):
    p = _pending_path(token)
    if not os.path.exists(p):
        return redirect("/templates")
    with open(p, "rb") as f:
        data = f.read()
    slides = templatize.list_slides(data)
    body = render_template("templatize_pick.html", token=token, slides=slides)
    return shell(body, active="templates", crumb="<b>Templates</b> / Learn a template")


@bp.route("/templatize/review/<token>/<int:slide_index>")
def templatize_review(token, slide_index):
    p = _pending_path(token)
    if not os.path.exists(p):
        return redirect("/templates")
    with open(p, "rb") as f:
        data = f.read()
    shapes = templatize.shapes_of(data, slide_index)
    roles = templatize.propose_roles(shapes)
    rows = [dict(s, role=roles.get(str(s["idx"]), "skip")) for s in shapes]
    body = render_template("templatize_review.html", token=token, slide_index=slide_index,
                           rows=rows, role_vocab=templatize.ROLE_VOCAB,
                           role_labels=templatize.ROLE_LABELS)
    return shell(body, active="templates", crumb="<b>Templates</b> / Confirm mapping")


@bp.route("/templatize/save", methods=["POST"])
def templatize_save():
    token = request.form.get("token", "")
    slide_index = int(request.form.get("slide_index", "0"))
    name = request.form.get("name", "").strip()
    p = _pending_path(token)
    if not os.path.exists(p) or not name:
        return redirect("/templates")
    with open(p, "rb") as f:
        data = f.read()
    role_map = {}
    for key, val in request.form.items():
        if key.startswith("role__") and val in templatize.ROLE_VOCAB:
            role_map[key[len("role__"):]] = val
    templatize.save_learned_template(data, slide_index, role_map, name)
    try:
        os.remove(p)
    except OSError:
        pass
    return redirect("/templates")


@bp.route("/templatize/<tid>/activate", methods=["POST"])
def templatize_activate(tid):
    templatize.set_active(tid)
    return redirect("/templates")


@bp.route("/templatize/<tid>/deactivate", methods=["POST"])
def templatize_deactivate(tid):
    templatize.deactivate(tid)
    return redirect("/templates")


@bp.route("/templatize/<tid>/delete", methods=["POST"])
def templatize_delete(tid):
    templatize.delete(tid)
    return redirect("/templates")
