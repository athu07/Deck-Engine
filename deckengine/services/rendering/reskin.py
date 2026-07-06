# -*- coding: utf-8 -*-
"""
reskin.py  --  rebrand an uploaded PowerPoint into the J2W look WITHOUT changing
any of its content.

Approach: restyle the ORIGINAL deck in place. We open the uploaded file and keep
every shape exactly where it is — all text, tables, images (including vector WMF/EMF),
charts, positions and the slide size are untouched — and only:
  * swap fonts to the J2W pair (Oswald for headings/subtitles, Raleway for body),
  * recolour dark/inherited heading text to J2W teal (light headings are left alone
    so white-on-colour titles stay readable),
  * add a J2W brand overlay to every slide (a slim teal+red top accent bar and a
    'JoulesToWatts' wordmark).

Because nothing is rebuilt, nothing can be dropped — the exact tradeoff the owner
asked for ("every word, table, image, heading preserved"). The preview is produced
by rendering the restyled deck to real page images via LibreOffice (render_pngs), so
what you see is exactly what downloads.
"""

import glob
import os
import subprocess

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN

# ── J2W identity ──────────────────────────────────────────────────────────────
TEAL = RGBColor(0x2C, 0x6E, 0x66)
RED = RGBColor(0xC0, 0x20, 0x26)
INK = RGBColor(0x11, 0x11, 0x10)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xE7, 0xF0, 0xEE)
HEAD_FONT = "Oswald"        # matches the case-study heading font
BODY_FONT = "Raleway"       # matches the case-study content font

_HEAD_PH = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)


# ── 1) restyle the original deck in place ─────────────────────────────────────
def restyle_deck(data, out_path):
    """Rebrand the uploaded pptx bytes to J2W and save to out_path (content intact)."""
    prs = Presentation(_as_stream(data))
    for slide in prs.slides:
        _restyle_shapes(slide.shapes)
        _brand_slide(slide, prs.slide_width, prs.slide_height)
    out_dir = os.path.dirname(out_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    prs.save(out_path)
    return out_path


def _as_stream(data):
    import io
    return io.BytesIO(data) if isinstance(data, (bytes, bytearray)) else data


def _role(shape):
    """'head' | 'sub' | 'body' — how a text shape should be styled."""
    name = (shape.name or "").lower()
    try:
        if shape.is_placeholder:
            t = shape.placeholder_format.type
            if t in _HEAD_PH:
                return "head"
            if t == PP_PLACEHOLDER.SUBTITLE:
                return "sub"
    except Exception:
        pass
    if "subtitle" in name:
        return "sub"
    if "title" in name:
        return "head"
    return "body"


def _restyle_shapes(shapes):
    for sh in shapes:
        try:
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                _restyle_shapes(sh.shapes)
                continue
        except Exception:
            pass
        try:
            if sh.has_table:                       # tables: J2W palette, data untouched
                _restyle_table(sh.table)
                continue
        except Exception:
            pass
        try:
            if not sh.has_text_frame:
                continue
        except Exception:
            continue
        role = _role(sh)
        font = HEAD_FONT if role in ("head", "sub") else BODY_FONT
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                run.font.name = font
                if role == "head":
                    _teal_if_dark(run)


def _restyle_table(table):
    """Recolour a table to the J2W palette (teal header, striped body, Raleway text)
    WITHOUT touching any cell text — only fills and fonts change."""
    for ri, row in enumerate(table.rows):
        header = (ri == 0)
        for cell in row.cells:
            try:
                cell.fill.solid()
                cell.fill.fore_color.rgb = TEAL if header else (SOFT if ri % 2 == 0 else WHITE)
            except Exception:
                pass
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.name = BODY_FONT
                    try:
                        run.font.color.rgb = WHITE if header else INK
                        if header:
                            run.font.bold = True
                    except Exception:
                        pass


def _teal_if_dark(run):
    """Recolour a heading run to J2W teal, but only when it is currently dark or
    inherits its colour — never override an explicit light colour (white titles on a
    dark section header must stay light)."""
    try:
        col = run.font.color
        ctype = col.type
        if ctype == MSO_COLOR_TYPE.RGB:
            rgb = col.rgb
            if (rgb[0] + rgb[1] + rgb[2]) < 460:      # dark-ish -> teal
                run.font.color.rgb = TEAL
        else:                                          # inherited/theme -> assume dark
            run.font.color.rgb = TEAL
    except Exception:
        pass


def _brand_slide(slide, sw, sh):
    """Add the J2W overlay: a slim teal top bar with a red segment + a wordmark."""
    bar_h = Inches(0.13)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, sw, bar_h)
    _flat_fill(bar, TEAL)
    seg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(1.7), bar_h)
    _flat_fill(seg, RED)
    wm = slide.shapes.add_textbox(sw - Inches(2.3), sh - Inches(0.4), Inches(2.15), Inches(0.3))
    tf = wm.text_frame
    tf.word_wrap = False
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = "JoulesToWatts"
    r.font.size = Pt(9)
    r.font.bold = True
    r.font.name = HEAD_FONT
    r.font.color.rgb = TEAL


def _flat_fill(shape, rgb):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb
    shape.line.fill.background()
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


# ── 2) render the restyled deck to page images (true preview) ─────────────────
def _on_path(name):
    from shutil import which
    return which(name) is not None


def render_pngs(pptx_path, out_dir, dpi=120):
    """Render every slide of pptx_path to a PNG in out_dir via LibreOffice (pptx->pdf)
    then poppler (pdf->png). Returns the ordered list of PNG paths, or [] if the tools
    aren't available (the caller then falls back to a download-only preview)."""
    soffice = "libreoffice" if _on_path("libreoffice") else ("soffice" if _on_path("soffice") else None)
    if not soffice or not _on_path("pdftoppm"):
        return []
    os.makedirs(out_dir, exist_ok=True)
    profile = "file://" + os.path.join(out_dir, ".loprofile")
    try:
        subprocess.run(
            [soffice, "-env:UserInstallation=" + profile, "--headless", "--convert-to",
             "pdf", "--outdir", out_dir, pptx_path],
            timeout=150, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        return []
    base = os.path.splitext(os.path.basename(pptx_path))[0]
    pdf = os.path.join(out_dir, base + ".pdf")
    if not os.path.exists(pdf):
        return []
    try:
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf, os.path.join(out_dir, "slide")],
                       timeout=150, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        return []
    return sorted(glob.glob(os.path.join(out_dir, "slide*.png")))
