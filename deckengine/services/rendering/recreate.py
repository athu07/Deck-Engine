# -*- coding: utf-8 -*-
"""
recreate.py  --  "Recreate with AI": rebuild an uploaded deck from scratch in
J2W's own visual language, slide by slide, instead of restyling the original
shapes in place (see reskin.py for that -- kept as a separate, faster, always-
guaranteed-fidelity option; owner's spec, 2026-07-09).

Owner's reference case (2026-07-09): a GenSpark-produced version of an
uploaded capability deck kept every fact/number from the source verbatim but
gave each slide a layout matched to what that slide's content actually is --
a stat-tile overview, a 3-feature "pillar" deep-dive, a named list with score
chips, a small data table -- instead of a font/colour swap of the original
shapes. This module reproduces that: for each ORIGINAL slide, extract its
text, classify which content SHAPE it is (from the SAME registry
slide_generator.CONTENT_TEMPLATES the "Already have the content?" flow uses,
plus one local escape hatch for slides that don't fit), extract structured
fields for that shape, and render a FRESH J2W-styled slide from those fields.

A slide that doesn't fit any known shape (a title/cover, a closing/contact
slide, or something genuinely data-heavy/visual a text template would lose
meaning from) falls back to reskin's in-place restyle for THAT slide only --
never forced into a mould that doesn't fit.

Bookended the same way reskin.py is: our own title slide (CS01) first, our
own closing slide (CS08) last, pulled verbatim from the master deck. The
FIRST and LAST slide of the uploaded deck are skipped outright (almost always
the source's own cover/closing, made redundant by our bookends); every slide
in between is recreated or restyled.
"""

import io
import json
import os

from pptx import Presentation

from deckengine import config
from deckengine.services.rendering import reskin, skills, slide_generator
from deckengine.services.rendering.slide_generator import CONTENT_TEMPLATES

_NONE_KEY = "_none"

# key -> (mapping fn for the header markers, draw fn for the programmatic
# body). case_study is handled separately (its own template file + the
# active-learned-template precedence -- see _render_case_study).
_MAPPING_FNS = {
    "four_box": skills._mapping_four_box,
    "roadmap_board": skills._mapping_roadmap_head,
    "box_grid": skills._mapping_box_grid_head,
    "pillar_deepdive": skills._mapping_pillar_head,
    "scored_list": skills._mapping_scored_list_head,
    "stat_overview": skills._mapping_stat_overview_head,
    "data_table": skills._mapping_data_table_head,
}
_DRAW_FNS = {
    "roadmap_board": skills._draw_roadmap_columns,
    "box_grid": skills._draw_box_grid,
    "pillar_deepdive": skills._draw_pillar_blocks,
    "scored_list": skills._draw_scored_rows,
    "stat_overview": skills._draw_stat_overview,
    "data_table": skills._draw_data_table,
}


def _as_stream(data):
    return io.BytesIO(data) if isinstance(data, (bytes, bytearray)) else data


def _slide_text(slide):
    """All of a slide's own text, in document order -- the same shape the
    paste-content classifier already expects."""
    lines = []
    for sh in slide.shapes:
        try:
            if sh.has_text_frame and sh.text_frame.text.strip():
                lines.append(sh.text_frame.text.strip())
        except Exception:
            pass
    return "\n".join(lines)


def _classify_slide(text):
    """Same registry, same A/B/C... pattern as slide_generator.classify_content,
    but with one extra escape hatch: NONE of these, when a slide is a title/
    closing slide or is genuinely data-heavy/visual. Fails safe to _NONE_KEY
    (restyle in place) rather than forcing a bad fit."""
    text = (text or "").strip()
    if not text:
        return _NONE_KEY
    letters = "ABCDEFGHIJK"
    choices = "\n".join(f"({letters[i]}) {t['classify_desc']}"
                        for i, t in enumerate(CONTENT_TEMPLATES))
    none_letter = letters[len(CONTENT_TEMPLATES)]
    key_by_letter = {letters[i]: t["key"] for i, t in enumerate(CONTENT_TEMPLATES)}
    key_by_letter[none_letter] = _NONE_KEY
    prompt = (
        "Does the content below (one slide's own text) read as one of these "
        "template shapes?\n" + choices +
        f"\n({none_letter}) NONE OF THESE -- a title/cover slide, a closing/"
        "contact/CTA slide, or content that's genuinely data-heavy/visual (a "
        "real chart, dense table, or diagram the text alone can't represent) "
        "and would lose meaning forced into a text-only template.\n\n"
        "If genuinely unsure, or the content doesn't clearly fit one shape "
        f"well, answer ({none_letter}).\n\nCONTENT:\n\"\"\"\n" + text[:4000] + "\n\"\"\"\n\n"
        'Reply with ONLY this JSON: {"choice":"A"} -- the single letter of your pick.'
    )
    try:
        from deckengine.services.infra import load_env
        load_env()
        from openai import OpenAI
        resp = OpenAI().chat.completions.create(
            model=slide_generator.MODEL, temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You classify one slide's text "
                 "into one of several template shapes, or 'none' if it "
                 "doesn't fit. Reply with one JSON object only."},
                {"role": "user", "content": prompt},
            ],
        )
        data = json.loads(resp.choices[0].message.content)
        letter = str(data.get("choice", "")).strip().upper()[:1]
        return key_by_letter.get(letter, _NONE_KEY)
    except Exception:
        return _NONE_KEY


def _restyle_original_slide(dest, src_slide, sw, sh, dest_sw, dest_sh):
    """Fallback for a slide that doesn't fit any content shape: copy the
    ORIGINAL slide and run reskin's own in-place restyle machinery on it
    (role detection, heading bar, brand overlay -- all proportioned to the
    slide's OWN original canvas, sw/sh), so it still gets J2W fonts/colours/
    branding even though its own layout is kept as-authored. THEN rescale the
    whole result from that original canvas to the STANDARD J2W canvas
    (dest_sw/dest_sh) the rest of this deck is being built at (see
    recreate_deck) -- every template-rendered slide is natively authored at
    that standard size already; only this restyle-in-place fallback path
    keeps the source's own (possibly different) coordinates and needs
    scaling. A no-op when the source already IS that standard size."""
    new = slide_generator._copy_slide(dest, src_slide)
    size_ranks = reskin._slide_sizes(new.shapes)
    dark_bg = reskin._slide_is_dark(new, new.shapes, sw, sh)
    reskin._restyle_shapes(new.shapes, size_ranks, sw, sh, dark_bg)
    reskin._fix_text_on_box_contrast(new.shapes)
    reskin._fix_picture_contrast(new.shapes)
    reskin._resolve_bottom_banner_overlap(new.shapes, sw, sh)
    heading = reskin._first_heading_shape(new.shapes, size_ranks)
    if heading is not None:
        reskin._add_heading_bar(new, heading)
    reskin._whiten_slide_background(new)
    reskin._brand_slide(new, sw, sh)
    if sw and sh and (sw, sh) != (dest_sw, dest_sh):
        reskin._scale_shapes(new.shapes, dest_sw / sw, dest_sh / sh)
    return new


def _render_case_study(dest, record):
    """Mirrors skills.build_into's own case_study_v2 dispatch exactly: the
    owner's ACTIVE learned template wins if one exists, else the built-in
    case_study_v2 -- kept consistent with every other AI-created case study
    in the app, not a parallel reimplementation."""
    from deckengine.services.rendering import templatize as _templatize
    from deckengine.services.rendering import fill_case_study as _fcs
    active = _templatize.active_template()
    if active:
        _templatize.fill_into(dest, active, record)
        return
    case_tpl = skills.find_template(Presentation(config.CASE_TEMPLATE_PPTX), "case_study_v2")
    if case_tpl is None:
        return
    new = slide_generator._copy_slide(dest, case_tpl)
    _fcs.apply_markers(new, _fcs.build_mapping(record))


def _render_shape(dest, tfile, key, record):
    template_slide = skills.find_template(tfile, record.get("template", key))
    if template_slide is None:
        return False
    new = slide_generator._copy_slide(dest, template_slide)
    mapping_fn = _MAPPING_FNS.get(key)
    if mapping_fn:
        skills.fill_markers(new, mapping_fn(record))
    draw_fn = _DRAW_FNS.get(key)
    if draw_fn:
        draw_fn(new, record)
    return True


def recreate_deck(data, out_path, industry=""):
    """Rebuild the uploaded pptx bytes into a FRESH J2W-styled deck, slide by
    slide, and save to out_path. Returns (out_path, stats) where stats =
    {"recreated": n, "restyled": n, "skipped": n} for the caller to report."""
    src_prs = Presentation(_as_stream(data))
    src_slides = list(src_prs.slides)
    sw, sh = src_prs.slide_width, src_prs.slide_height   # the UPLOADED deck's own canvas

    # The output canvas is ALWAYS the standard J2W size (matches the master
    # deck / skills_templates.pptx / case_study_v2, all 13.33x7.5in) -- NOT
    # the uploaded deck's own size. Every template-rendered slide below is
    # natively authored at that standard size, so it needs no scaling; only
    # the restyle-in-place fallback (which keeps the source's own shape
    # coordinates) needs rescaling to fit, same as the bookend slides
    # already do. Preserving the source's own (possibly different) canvas
    # size instead was the root cause of a real bug: on a source deck
    # smaller than 13.33x7.5, every template slide overflowed past the
    # actual slide edge and every font rendered oversized relative to it
    # (owner-reported, 2026-07-09, with the actual output file attached).
    try:
        master = Presentation(config.MASTER_DECK)
        dest_sw, dest_sh = master.slide_width, master.slide_height
    except Exception:
        dest_sw, dest_sh = sw, sh   # fail-safe: fall back to the source's own size

    dest = Presentation()
    dest.slide_width, dest.slide_height = dest_sw, dest_sh
    tfile = Presentation(config.SKILLS_TEMPLATES_PPTX)

    stats = {"recreated": 0, "restyled": 0, "skipped": 0}
    middle = src_slides[1:-1] if len(src_slides) > 2 else src_slides
    stats["skipped"] = len(src_slides) - len(middle)

    for slide in middle:
        text = _slide_text(slide)
        key = _classify_slide(text)
        if key == _NONE_KEY:
            _restyle_original_slide(dest, slide, sw, sh, dest_sw, dest_sh)
            stats["restyled"] += 1
            continue
        tdef = next((t for t in CONTENT_TEMPLATES if t["key"] == key), None)
        if tdef is None:
            _restyle_original_slide(dest, slide, sw, sh, dest_sw, dest_sh)
            stats["restyled"] += 1
            continue
        record = tdef["builder"](text, {"industry": industry})
        if key == "case_study":
            from deckengine.services.rendering.deck_build import ai_to_store_record
            _render_case_study(dest, ai_to_store_record(record, industry))
            stats["recreated"] += 1
            continue
        if _render_shape(dest, tfile, key, record):
            stats["recreated"] += 1
        else:
            _restyle_original_slide(dest, slide, sw, sh, dest_sw, dest_sh)
            stats["restyled"] += 1

    _append_bookends(dest, dest_sw, dest_sh)
    out_dir = os.path.dirname(out_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    dest.save(out_path)
    return out_path, stats


def _append_bookends(dest, sw, sh):
    """Our own title slide (CS01) first, our own closing slide (CS08) last --
    pulled verbatim from the master deck, same source reskin.py's bookends
    use, scaled to the working canvas size."""
    from deckengine.services.content.build_library import read_id
    try:
        master = Presentation(config.MASTER_DECK)
    except Exception:
        return
    title_src = winback_src = None
    for slide in master.slides:
        sid = read_id(slide)
        if sid == "CS01":
            title_src = slide
        elif sid == "CS08":
            winback_src = slide
    msw, msh = master.slide_width, master.slide_height
    sx = (sw / msw) if msw else 1.0
    sy = (sh / msh) if msh else 1.0
    if title_src is not None:
        new_title = slide_generator._copy_slide(dest, title_src)
        reskin._scale_shapes(new_title.shapes, sx, sy)
        reskin._move_slide_to_front(dest, new_title)
    if winback_src is not None:
        new_close = slide_generator._copy_slide(dest, winback_src)
        reskin._scale_shapes(new_close.shapes, sx, sy)
