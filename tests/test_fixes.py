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
