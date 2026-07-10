# -*- coding: utf-8 -*-
"""
preview.py  --  render an EXISTING library case study to a PNG.

Used by one thing: the Custom Slide Builder's duplicate check. When it says "MSS042 is
91% similar to what you pasted", you need to SEE MSS042 before deciding to reuse it,
and that slide lives only in the content store -- there is nothing on screen to look at.

Slides you BUILD are not rendered here. They come back as editable slide cards (the
review page's `.se` inputs, via templates/_slide_editor.html), because a picture of a
slide is something you look at and a card is something you fix. Rendering them to PNG
also cost ~3 seconds each, through a headless LibreOffice, for an image you'd then have
to edit somewhere else anyway.

Rendering is slow, so every PNG is cached under static/renders/builder/ and rendered once.

FAIL-SAFE, like every other optional step in this app: if LibreOffice or poppler is
missing, or a render times out, case_png() returns None and the caller says so plainly.
"""

import os
import shutil
import tempfile

from deckengine import config

# Where the browser fetches previews from, and where they're cached on disk. Under
# static/renders/ (config.RENDERS_DIR) so Flask serves them without a route.
_CACHE_DIR = os.path.join(config.RENDERS_DIR, "builder")
_URL_PREFIX = "/static/renders/builder/"


def _cache_path(key):
    return os.path.join(_CACHE_DIR, str(key) + ".png")


def cached_url(key):
    """The public URL for an already-rendered preview, or None if it isn't cached."""
    return _URL_PREFIX + str(key) + ".png" if os.path.exists(_cache_path(key)) else None


def _render_first_slide(pptx_path, dest_png):
    """pptx -> PNG of slide 1, at dest_png. Raises if the tools aren't available."""
    from deckengine.services.rendering import reskin

    with tempfile.TemporaryDirectory() as td:
        pngs = reskin.render_pngs(pptx_path, td)
        if not pngs:
            raise RuntimeError("LibreOffice/poppler unavailable")
        os.makedirs(os.path.dirname(dest_png), exist_ok=True)
        shutil.move(pngs[0], dest_png)
    return dest_png


def case_png(case_id, force=False):
    """Render an EXISTING content-store case (AIP/WFS/MSS) to a cached PNG -- so a
    near-duplicate can be seen before it's reused. Returns its URL, or None."""
    from deckengine.services.content import case_library
    from deckengine.services.rendering import fill_case_study

    key = "case_" + str(case_id)
    if not force:
        hit = cached_url(key)
        if hit:
            return hit
    rec = case_library.record(case_id)
    if not rec:
        return None
    try:
        with tempfile.TemporaryDirectory() as td:
            deck = os.path.join(td, "slide.pptx")
            fill_case_study.fill_row(rec, deck)
            _render_first_slide(deck, _cache_path(key))
        return _URL_PREFIX + key + ".png"
    except Exception:
        return None
