# -*- coding: utf-8 -*-
"""
Smoke-test safety net for the Deck Engine refactor.

Goal: give the refactor a fast, hermetic regression guard so each phase can be
verified green. Two things are checked:

  1. test_matcher_golden — matcher.plan(ctx, use_ai=False) returns an EXACT, frozen
     list of slide IDs for fixed contexts. This is the behaviour-diff guard: if a
     refactor changes matching output, the ID lists change and this fails.

  2. test_build_review_finalize — the real Flask flow /build -> /review -> /finalize
     runs end-to-end and produces a .pptx with slides.

Hermetic by design:
  - OPENAI_API_KEY is blanked, so every AI call (embeddings / extract / explain)
    fails safe to its offline default — no network, fully deterministic.
  - Output and meeting-log dirs are redirected to a temp dir, so the repo working
    tree is never touched.
  - The AI-accept / staging.promote path is never exercised, so the living master
    deck (WORKING_COPY_Master_Deck.pptx) is never mutated.

Run from anywhere:  .venv/bin/python -m pytest tests/ -q
"""

import os
import pathlib
import tempfile

# Redirect the app's writable dirs to throwaway temp dirs BEFORE anything imports
# config/app, so the real output/, meetings/ and staging/ are never touched.
_TMP = tempfile.mkdtemp(prefix="deck_smoke_")
os.environ["DECK_OUTPUT_DIR"] = os.path.join(_TMP, "out")
os.environ["DECK_MEETINGS_DIR"] = os.path.join(_TMP, "meet")
os.environ["DECK_STAGING_DIR"] = os.path.join(_TMP, "stage")
os.environ["DECK_BUILD_CONTEXT_DIR"] = os.path.join(_TMP, "bctx")
for _d in ("out", "meet", "stage", "bctx"):
    os.makedirs(os.path.join(_TMP, _d), exist_ok=True)

import pytest

REPO = pathlib.Path(__file__).resolve().parent.parent


# ── frozen baseline (captured offline from the current code) ──────────────────
GOLDEN = {
    "ms_manuf": (
        {"client_name": "Schneider Electric", "industry": "MANUFACTURING",
         "work_types": ["MS"], "functions": [], "recipient": "CTO",
         "transcript": "We need test automation, predictive maintenance and "
                       "digital twin for our plants."},
        ['CS01', 'CS02', 'CS03', 'CS04', 'CS05', 'CS06', 'CS18', 'CS19',
         'MSS012', 'MSS005', 'MSS034', 'CS07', 'CS08'],
    ),
    "wf_health": (
        {"client_name": "HPE", "industry": "HEALTHCARE",
         "work_types": ["WORKFORCE"], "functions": [], "recipient": "Head of Talent",
         "transcript": "Looking for RPO and cloud migration staffing for healthcare IT."},
        ['CS01', 'CS02', 'CS03', 'CS04', 'CS05', 'CS06', 'CS10', 'CS11', 'CS12',
         'CS13', 'WFS022', 'WFS004', 'WFS002', 'CS07', 'CS08'],
    ),
}


@pytest.fixture(scope="session", autouse=True)
def _run_from_repo_root():
    """The app resolves every data file relative to CWD; pin CWD to the repo root."""
    old = os.getcwd()
    os.chdir(REPO)
    try:
        yield
    finally:
        os.chdir(old)


@pytest.fixture(autouse=True)
def _hermetic_ai(monkeypatch):
    """Force every OpenAI call to fail fast so the app takes its offline fail-safe
    paths — deterministic, no network, no cost. A non-empty sentinel key stops
    infra.load_env from repopulating the real key from .env (it only skips keys that
    are already set to a non-empty value); an unreachable base URL makes any call
    that is still attempted fail instantly instead of hitting the real API."""
    monkeypatch.setenv("OPENAI_API_KEY", "sk-hermetic-test-not-a-real-key")
    monkeypatch.setenv("OPENAI_BASE_URL", "http://127.0.0.1:1")


@pytest.fixture
def client():
    """Flask test client. Writable dirs are already redirected to temp via the
    DECK_*_DIR env vars set at import (see top of file)."""
    import app as appmod
    appmod.app.config.update(TESTING=True)
    appmod.app._smoke_out_dir = os.environ["DECK_OUTPUT_DIR"]   # handed to the test
    return appmod.app.test_client()


# ── 1. behaviour-diff guard (offline, deterministic) ──────────────────────────
@pytest.mark.parametrize("name", list(GOLDEN))
def test_matcher_golden(name):
    from deckengine.services.matching import matcher
    ctx, expected = GOLDEN[name]
    ids = [p["slide_id"] for p in matcher.plan(ctx, use_ai=False)["picks"]]
    assert ids == expected, f"{name} matching drifted:\n got: {ids}\n want: {expected}"


# ── 2. full pipeline: /build -> /review -> /finalize -> .pptx ─────────────────
def test_build_review_finalize(client):
    from pptx import Presentation

    ctx, ids = GOLDEN["ms_manuf"]
    form = {
        "client_name": "Smoke Test Co",
        "industry": ctx["industry"],
        "work_types": ctx["work_types"],
        "functions": ctx["functions"],
        "phase": "First meeting",
        "recipient": ctx["recipient"],
        "transcript": ctx["transcript"],
    }

    # /build — exercises matcher.plan(use_ai=True) offline + full page render
    r = client.post("/build", data=form)
    assert r.status_code == 200, r.status_code
    assert b"Review &amp; edit" in r.data or b"Why this deck matches" in r.data

    # /review — server builds editable cards for the picked ids
    order = ",".join(ids)
    r = client.post("/review", data={**form, "order": order})
    assert r.status_code == 200, r.status_code

    # /finalize — assembles the real .pptx (core slides + MSS store cases rendered)
    r = client.post("/finalize", data={**form, "order": order})
    assert r.status_code == 200, r.status_code
    assert b"Tailored_Deck_Smoke_Test_Co.pptx" in r.data

    built = pathlib.Path(client.application._smoke_out_dir) / "Tailored_Deck_Smoke_Test_Co.pptx"
    assert built.exists(), f"deck not written to {built}"
    assert len(Presentation(str(built)).slides) > 0


# ── 3. every live page route renders (guards routing/template refactors) ───────
@pytest.mark.parametrize("route", ["/", "/new", "/dashboard", "/library",
                                   "/staging", "/templates", "/meetings", "/deck"])
def test_get_routes_ok(client, route):
    r = client.get(route)
    assert r.status_code == 200, f"{route} -> {r.status_code}"
