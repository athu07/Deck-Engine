# -*- coding: utf-8 -*-
"""
test_fixes.py  --  hermetic guards for the two core fixes (no network).

  1. Issue 1: a mismatch-flagged capability in the WRONG domain is demoted by
     rank_cases (so it can't win the fill).
  2. Issue 2: build_context saves/loads the rich context by build_id.
  3. Issue 2: draft_case_study actually feeds research + profile + notes into the
     generation prompt (they used to be discarded).
"""

import os
import pathlib

REPO = pathlib.Path(__file__).resolve().parent.parent


def _chdir():
    os.chdir(REPO)


# ── Issue 1: mismatch-flag demotion ───────────────────────────────────────────
def test_avoid_flag_demotes_cross_domain_case():
    _chdir()
    from deckengine.services.matching import relevance
    rows = [
        {"slide_id": "SAME", "title": "Computer Vision Inspection", "keywords": "computer vision",
         "primary_industry": "MANUFACTURING", "primary_function": "", "work_types": "MS",
         "search_text": "computer vision inspection manufacturing"},
        {"slide_id": "CROSS", "title": "Computer Vision for Web QA", "keywords": "computer vision",
         "primary_industry": "TECH_IT", "primary_function": "", "work_types": "MS",
         "search_text": "computer vision web qa"},
    ]
    ranked = relevance.rank_cases(
        "we need computer vision", rows, industry="MANUFACTURING", use_semantic=False,
        avoid=[{"capability": "computer vision", "reason": "meant for web QA, not this domain"}])
    score = {r["row"]["slide_id"]: r["score"] for r in ranked}
    # the flagged, wrong-domain case is pushed well below the same-domain one
    assert score["CROSS"] < score["SAME"]
    assert score["CROSS"] < 0            # penalised below the keep floor

    # without the avoid flag the cross-domain case is NOT penalised
    plain = relevance.rank_cases("we need computer vision", rows,
                                 industry="MANUFACTURING", use_semantic=False)
    plain_score = {r["row"]["slide_id"]: r["score"] for r in plain}
    assert plain_score["CROSS"] > score["CROSS"]


# ── Issue 2: build-context persistence ────────────────────────────────────────
def test_build_context_roundtrip(tmp_path, monkeypatch):
    from deckengine import config
    from deckengine.services import build_context
    monkeypatch.setattr(config, "BUILD_CONTEXT_DIR", str(tmp_path))
    build_context.save("abc123DEF", {"research": "R", "profile": "P", "transcript": "T",
                                     "industry": "MANUFACTURING"})
    got = build_context.load("abc123DEF")
    assert got["research"] == "R" and got["profile"] == "P" and got["industry"] == "MANUFACTURING"
    assert build_context.load("does-not-exist") == {}   # missing -> {} (never raises)
    assert build_context.load("") == {}                 # empty id -> {}


# ── Issue 2: generation prompt actually receives the rich context ─────────────
def test_draft_case_study_feeds_research_profile_notes(monkeypatch):
    _chdir()
    import openai

    captured = {}

    class _FakeClient:
        class chat:
            class completions:
                @staticmethod
                def create(**kw):
                    captured["prompt"] = kw["messages"][-1]["content"]
                    raise RuntimeError("stop after capture")   # -> draft takes fail-safe path

    monkeypatch.setattr(openai, "OpenAI", lambda *a, **k: _FakeClient())
    from deckengine.services.rendering import slide_generator
    out = slide_generator.draft_case_study(
        "computer vision inspection",
        {"industry": "MANUFACTURING", "recipient": "CTO", "function": "Quality",
         "research": "RESEARCH_MARKER_XYZ", "profile": "PROFILE_MARKER_XYZ",
         "notes": "TRANSCRIPT_MARKER_XYZ"})
    p = captured.get("prompt", "")
    assert "RESEARCH_MARKER_XYZ" in p, "deep research not in generation prompt"
    assert "PROFILE_MARKER_XYZ" in p, "profile not in generation prompt"
    assert "TRANSCRIPT_MARKER_XYZ" in p, "transcript not in generation prompt"
    # fail-safe still returns a well-formed placeholder (6 caps / 3 results)
    assert len(out["capabilities"]) == 6 and len(out["results"]) == 3


# ── F2: re-skin restyles in place and preserves ALL content ───────────────────
def test_reskin_preserves_content(tmp_path):
    _chdir()
    import io
    from pptx import Presentation
    from pptx.util import Inches
    from deckengine.services.rendering import reskin
    src = Presentation()
    s = src.slides.add_slide(src.slide_layouts[5])
    s.shapes.title.text = "Hello Deck"
    tbl = s.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(4), Inches(1)).table
    tbl.cell(0, 0).text = "A"; tbl.cell(0, 1).text = "B"
    tbl.cell(1, 0).text = "1"; tbl.cell(1, 1).text = "ZZZVAL"
    buf = io.BytesIO(); src.save(buf)
    out = str(tmp_path / "reskinned.pptx")
    reskin.restyle_deck(buf.getvalue(), out)
    r = Presentation(out)
    text = " ".join(sh.text_frame.text for sl in r.slides for sh in sl.shapes if sh.has_text_frame)
    cells = " ".join(c.text for sl in r.slides for sh in sl.shapes if sh.has_table
                     for row in sh.table.rows for c in row.cells)
    assert "Hello Deck" in text          # title preserved
    assert "ZZZVAL" in cells             # table data preserved untouched
    # J2W branding applied: the title run is now Oswald
    title_fonts = {run.font.name for sl in r.slides for sh in sl.shapes
                   if sh.has_text_frame and "Hello Deck" in sh.text_frame.text
                   for p in sh.text_frame.paragraphs for run in p.runs}
    assert "Oswald" in title_fonts


# ── Library bulk download: several slides -> one .zip of per-slide .pptx ───────
def test_bulk_slides_zip():
    _chdir()
    import io
    import zipfile
    import app as appmod
    c = appmod.app.test_client()
    r = c.get("/slides/download?ids=CS01,MSS001")
    assert r.status_code == 200
    assert r.headers["Content-Type"] == "application/zip"
    names = zipfile.ZipFile(io.BytesIO(r.data)).namelist()
    assert "CS01.pptx" in names and "MSS001.pptx" in names   # master + content-store case
    assert c.get("/slides/download?ids=").status_code == 400  # nothing selected


# ── F3: saved-templates store round-trip ──────────────────────────────────────
def test_saved_templates_roundtrip(tmp_path, monkeypatch):
    from deckengine import config
    from deckengine.services import saved_templates
    monkeypatch.setattr(config, "SAVED_TEMPLATES_DIR", str(tmp_path))
    row = saved_templates.save("My Deck", b"PK-fake-pptx-bytes", 5)
    assert saved_templates.get(row["id"])["name"] == "My Deck"
    assert saved_templates.file_path(row["id"])
    assert saved_templates.delete(row["id"]) is True
    assert saved_templates.all_templates() == []
