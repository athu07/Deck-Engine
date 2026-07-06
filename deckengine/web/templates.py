# -*- coding: utf-8 -*-
"""templates.py  --  the /templates page: fill-on-demand templates, saved decks
(upload & reuse), and the deck re-skinner (upload -> J2W-branded -> download)."""

import io
import os
import re
import shutil

from flask import Blueprint, render_template, request, send_file, redirect, abort
from pptx import Presentation

from deckengine import config
from deckengine.services.rendering import slide_generator, reskin
from deckengine.services import saved_templates
from .view_helpers import shell

bp = Blueprint("templates", __name__)


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
                           saved=saved_templates.all_templates())
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
