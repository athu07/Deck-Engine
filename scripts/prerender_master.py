# -*- coding: utf-8 -*-
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from deckengine import config  # noqa: E402  (repo-root data paths)
"""
prerender_master.py
Render every slide of the master deck to a PNG, once, and COMMIT the result.

WHY THIS EXISTS
The review page shows a picture of each library slide, so the salesperson can see the
whole deck before downloading it. Rendering those at runtime was wrong twice over:

  * A plain Render web service has no LibreOffice and no pdftoppm. The renderer found
    neither on the PATH, returned nothing, and every preview 404'd -- the deployed app
    showed "this server can't render slide previews" on all thirteen slides
    (owner-reported, 2026-07-10).
  * A server that DOES have them paid a 40-slide LibreOffice conversion on the first
    request after every cold start. Render's disk is ephemeral, so that meant every
    deploy, on a small shared CPU, with LibreOffice's memory footprint on a 512MB box.

So the previews are built HERE, on a machine that has the tools, and shipped:

    static/previews/master/<content-hash>/CS01.png ...

The directory is named for the master deck's CONTENT HASH, not its mtime -- `git
checkout` rewrites mtimes, so an mtime key could never match a build-time render. Change
the deck, the hash changes, the old previews are ignored and this script must be re-run.

RUN IT whenever data/decks/WORKING_COPY_Master_Deck.pptx changes:

    python scripts/prerender_master.py

Needs LibreOffice + poppler (`soffice`, `pdftoppm`). The Dockerfile runs it at build
time; if you deploy without Docker, run it locally and commit what it writes.
"""

import shutil

from pptx import Presentation

from deckengine.services.content.build_library import read_id
from deckengine.services.rendering import preview, reskin


def _save_optimised(src_png, dest_webp, max_width=1280):
    """These get committed, so keep them small: WebP is ~58% lighter than PNG for these
    graphics, at a size no one can tell from the original. On-screen size is unchanged --
    the review card frames them at 16:9 regardless of pixel dimensions. The review card is
    ~1170px wide; anything past 1280 is bytes in git that nobody sees."""
    from PIL import Image
    im = Image.open(src_png).convert("RGB")
    if im.width > max_width:
        im = im.resize((max_width, round(im.height * max_width / im.width)), Image.LANCZOS)
    im.save(dest_webp, "WEBP", quality=85, method=6)


def main(dpi=90):
    key = preview.master_key()
    out_dir = preview._shipped_dir()
    root = _os.path.join(config.SLIDE_PREVIEWS_DIR, "master")

    ids = [read_id(s) for s in Presentation(config.MASTER_DECK).slides]
    print("master deck : %s" % config.MASTER_DECK)
    print("content hash: %s   (%d slides, %d with a J2W_ID)"
          % (key, len(ids), sum(1 for i in ids if i)))

    if _os.path.isdir(out_dir) and _os.listdir(out_dir):
        print("already rendered -> %s" % out_dir)
        return 0

    import tempfile
    with tempfile.TemporaryDirectory() as td:
        print("rendering (LibreOffice -> PDF -> PNG)...")
        pngs = reskin.render_pngs(config.MASTER_DECK, td, dpi=dpi)
        if not pngs:
            print("FAILED: LibreOffice and/or pdftoppm are not on the PATH.\n"
                  "        Install them (apt: libreoffice-impress poppler-utils) and re-run.")
            return 1
        if len(pngs) != len(ids):
            print("FAILED: rendered %d pages for %d slides -- the page order can't be "
                  "trusted to name the slides." % (len(pngs), len(ids)))
            return 1

        tmp_dir = out_dir + ".partial"
        shutil.rmtree(tmp_dir, ignore_errors=True)
        _os.makedirs(tmp_dir, exist_ok=True)
        written = 0
        for sid, png in zip(ids, pngs):
            if not sid:                      # template slides carry no J2W_ID
                continue
            _save_optimised(png, _os.path.join(tmp_dir, sid + ".webp"))
            written += 1
        _os.replace(tmp_dir, out_dir)        # atomic

    # a previous deck's previews are dead weight in the image and in git
    for stale in _os.listdir(root):
        if stale != key:
            shutil.rmtree(_os.path.join(root, stale), ignore_errors=True)
            print("removed stale previews for %s" % stale)

    total = sum(_os.path.getsize(_os.path.join(out_dir, f)) for f in _os.listdir(out_dir))
    print("wrote %d previews -> %s  (%.1f MB)" % (written, out_dir, total / 1048576))
    print("\nCommit static/previews/ so the deployed app has them.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
