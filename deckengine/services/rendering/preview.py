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

import glob
import os
import shutil
import tempfile
import threading

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


# ── the master deck's own slides (CSxx) ──────────────────────────────────────
# The review page shows a library slide as a title and a subtitle, because those are the
# only two fields the engine can edit. Everything else on the slide -- the bullets, the
# tables, the charts, the diagrams -- was invisible until you downloaded the .pptx. This
# renders the real thing.
#
# The whole master deck converts in ONE LibreOffice pass (40 slides, ~8.5s measured), so
# render it once and key each page to its slide id, rather than paying a separate 3-second
# render per slide per review page. The cache is keyed on the master deck's mtime, so
# editing the master invalidates every preview automatically.
_MASTER_LOCK = threading.Lock()


def _master_cache_dir():
    try:
        stamp = int(os.path.getmtime(config.MASTER_DECK))
    except OSError:
        stamp = 0
    return os.path.join(config.RENDERS_DIR, "master", str(stamp))


def _build_master_cache(cache_dir):
    """Render every slide of the master deck to <cache_dir>/<J2W_ID>.png, once."""
    from pptx import Presentation
    from deckengine.services.content.build_library import read_id
    from deckengine.services.rendering import reskin

    ids = [read_id(s) for s in Presentation(config.MASTER_DECK).slides]
    with tempfile.TemporaryDirectory() as td:
        pngs = reskin.render_pngs(config.MASTER_DECK, td)
        if not pngs:
            raise RuntimeError("LibreOffice/poppler unavailable")
        if len(pngs) != len(ids):
            # a mismatch means the page order can't be trusted to name the slides
            raise RuntimeError("rendered %d pages for %d slides" % (len(pngs), len(ids)))
        tmp_dir = cache_dir + ".partial"
        shutil.rmtree(tmp_dir, ignore_errors=True)
        os.makedirs(tmp_dir, exist_ok=True)
        for sid, png in zip(ids, pngs):
            if sid:                       # template slides carry no J2W_ID
                shutil.move(png, os.path.join(tmp_dir, sid + ".png"))
        os.replace(tmp_dir, cache_dir)    # atomic: a reader sees all of it or none
    return cache_dir


def master_slide_png(sid):
    """The rendered PNG path for one master-deck slide (CSxx), or None.

    First call for a given master deck renders all 40 slides (~8.5s) and caches them;
    every call after that is a stat(). Serialised, so ten browsers asking at once render
    it once. Fail-safe: no LibreOffice, no preview -- the caller shows the text card."""
    cache_dir = _master_cache_dir()
    path = os.path.join(cache_dir, sid + ".png")
    if os.path.exists(path):
        return path
    with _MASTER_LOCK:
        if os.path.exists(path):         # another thread built it while we waited
            return path
        try:
            _build_master_cache(cache_dir)
        except Exception:
            return None
        # drop caches for older versions of the master deck
        for stale in glob.glob(os.path.join(config.RENDERS_DIR, "master", "*")):
            if stale != cache_dir:
                shutil.rmtree(stale, ignore_errors=True)
    return path if os.path.exists(path) else None


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
