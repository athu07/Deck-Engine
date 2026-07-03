# -*- coding: utf-8 -*-
"""
renderer.py  --  render a .pptx to high-fidelity slide images using the installed
PowerPoint (COM automation). This gives a true, professional PowerPoint-quality
preview (exactly how the slides look), unlike a best-effort JS renderer.

Images are written under static/renders/<key>/ (key = filename + mtime) so they
serve directly via Flask's /static and are cached — re-viewing is instant.
"""

import os

RENDER_ROOT = os.path.join("static", "renders")
WIDTH = 1600   # export width in px; height kept proportional by PowerPoint


def _key(pptx_path):
    base = os.path.splitext(os.path.basename(pptx_path))[0]
    return "%s_%d" % (base, int(os.path.getmtime(pptx_path)))


def render(pptx_path):
    """Return a list of web paths to slide PNGs, e.g.
    ['/static/renders/Deck_123/slide_001.png', ...]. Empty list on failure."""
    key = _key(pptx_path)
    out_dir = os.path.join(RENDER_ROOT, key)
    web = "/static/renders/" + key

    if os.path.isdir(out_dir):                       # cache hit
        pngs = sorted(f for f in os.listdir(out_dir) if f.endswith(".png"))
        if pngs:
            return [web + "/" + f for f in pngs]

    os.makedirs(out_dir, exist_ok=True)
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    pres = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        pres = app.Presentations.Open(os.path.abspath(pptx_path),
                                       WithWindow=False, ReadOnly=True)
        paths = []
        for i in range(1, pres.Slides.Count + 1):
            fn = os.path.abspath(os.path.join(out_dir, "slide_%03d.png" % i))
            pres.Slides(i).Export(fn, "PNG", WIDTH)
            paths.append(web + "/slide_%03d.png" % i)
        return paths
    except Exception:
        return []
    finally:
        try:
            if pres is not None:
                pres.Close()                  # release the file lock
        except Exception:
            pass
        # NOTE: never app.Quit() — it could close the user's open PowerPoint.
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    import sys
    imgs = render(sys.argv[1])
    print("rendered %d slides:" % len(imgs))
    for p in imgs:
        print(" ", p)
