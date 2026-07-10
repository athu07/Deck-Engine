# -*- coding: utf-8 -*-
"""
test_fixes.py  --  hermetic guards for the two core fixes (no network).

  1. Issue 1: a mismatch-flagged capability in the WRONG domain is demoted by
     rank_cases (so it can't win the fill).
  2. Issue 2: build_context saves/loads the rich context by build_id.
  3. Issue 2: draft_case_study actually feeds research + profile + notes into the
     generation prompt (they used to be discarded).

Plus the Custom Slide Builder guards (2026-07-10):
  4. dedupe.similar_cases fails safe to [] with no API key.
  5. staging.add round-trips EVERY shape's fields (it used to drop five of them).
  6. deck_build.staged_item routes all 8 shapes (five used to become case studies).
  7. promote_case is idempotent -- download AND deck-build saves the case once.
"""

import os
import pathlib

REPO = pathlib.Path(__file__).resolve().parent.parent


def _chdir():
    os.chdir(REPO)


def _isolate_staging(tmp_path, monkeypatch):
    """Point staging at a throwaway file -- these tests must never touch staging.json."""
    from deckengine.services.rendering import staging
    monkeypatch.setattr(staging, "STAGE_DIR", str(tmp_path))
    monkeypatch.setattr(staging, "STAGE_JSON", str(tmp_path / "staging.json"))
    return staging


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
    assert "HELLO DECK" in text          # title preserved (headings are uppercased by design)
    assert "ZZZVAL" in cells             # table data preserved untouched
    # J2W branding applied: the title run is now Oswald
    title_fonts = {run.font.name for sl in r.slides for sh in sl.shapes
                   if sh.has_text_frame and "HELLO DECK" in sh.text_frame.text
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


# ── Custom Slide Builder (2026-07-10) ─────────────────────────────────────────
def test_dedupe_fails_safe_without_a_key(monkeypatch):
    """A duplicate check must never be the reason a slide can't be built."""
    _chdir()
    from deckengine.services.matching import dedupe, relevance
    monkeypatch.setattr(relevance, "embed_texts", lambda texts: None)   # offline
    assert dedupe.similar_cases("any pasted content at all") == []
    assert dedupe.similar_cases("") == []                               # nothing to check


def test_dedupe_ranks_and_gates_by_work_type(monkeypatch):
    """Above the bar -> a match, best first; below it -> silence. Work type gates."""
    _chdir()
    from deckengine.services.matching import dedupe, relevance
    from deckengine.services.content import case_library

    monkeypatch.setattr(relevance, "embed_texts", lambda texts: [[1.0, 0.0]])
    monkeypatch.setattr(relevance, "_load_case_embeddings",
                        lambda: {"MSS900": [1.0, 0.0],        # identical -> 1.00
                                 "MSS901": [0.9, 0.436],      # ~0.90 -> above the bar
                                 "MSS902": [0.0, 1.0],        # orthogonal -> 0.00
                                 "WFS900": [1.0, 0.0]})       # identical, wrong work type
    monkeypatch.setattr(case_library, "_load", lambda: [
        {"id": "MSS900", "title": "Exact", "work_type": "MS", "domain": ""},
        {"id": "MSS901", "title": "Close", "work_type": "MS", "domain": ""},
        {"id": "MSS902", "title": "Unrelated", "work_type": "MS", "domain": ""},
        {"id": "WFS900", "title": "Other work type", "work_type": "WORKFORCE", "domain": ""},
    ])

    hits = dedupe.similar_cases("text", allowed_work_types=["MS"])
    assert [h["id"] for h in hits] == ["MSS900", "MSS901"]      # ranked, orthogonal dropped
    assert hits[0]["score"] > hits[1]["score"] >= dedupe.THRESHOLD
    assert "WFS900" not in [h["id"] for h in hits]              # gated out by work type
    # an MS case is not a duplicate of a Workforce one, even word-for-word
    assert [h["id"] for h in dedupe.similar_cases("text", allowed_work_types=["WORKFORCE"])] == ["WFS900"]


def test_staging_round_trips_every_shape_field(tmp_path, monkeypatch):
    """staging.add() used to drop eyebrow/blocks/rows/stats/items/col_labels, so those
    five shapes rendered empty (or fell through to the case-study path)."""
    _chdir()
    staging = _isolate_staging(tmp_path, monkeypatch)

    grid = staging.add({"content_type": "box_grid", "template": "box_grid", "title": "G",
                        "boxes": [{"heading": "h", "body": "b"}]}, "MS", "BFSI")
    assert staging.get(grid["id"])["boxes"] == [{"heading": "h", "body": "b"}]

    stats = staging.add({"content_type": "stat_overview", "template": "stat_overview",
                         "title": "S", "intro": "i", "stats": [{"value": "42%", "label": "x"}],
                         "items": ["a", "b"]}, "MS", "BFSI")
    back = staging.get(stats["id"])
    assert back["stats"] == [{"value": "42%", "label": "x"}] and back["items"] == ["a", "b"]

    table = staging.add({"content_type": "data_table", "template": "data_table", "title": "T",
                         "col_labels": ["A", "B", "C"], "rows": [{"label": "r"}]}, "MS", "BFSI")
    assert staging.get(table["id"])["col_labels"] == ["A", "B", "C"]

    pillar = staging.add({"content_type": "pillar_deepdive", "template": "pillar_deepdive",
                          "title": "P", "eyebrow": "e", "blocks": [{"heading": "h"}]}, "MS", "BFSI")
    back = staging.get(pillar["id"])
    assert back["eyebrow"] == "e" and back["blocks"] == [{"heading": "h"}]


def test_staged_item_routes_all_shapes():
    """Only case_study carries a content-store record; every other shape goes down the
    generic draw path. Before this, five shapes were silently built AS case studies."""
    _chdir()
    from deckengine.services.rendering import deck_build

    for shape in ("four_box", "roadmap_board", "box_grid", "pillar_deepdive",
                  "scored_list", "stat_overview", "data_table"):
        item = deck_build.staged_item("NEW:G1", {"content_type": shape, "template": shape,
                                                 "title": "X"}, "BFSI")
        assert item["kind"] == shape and item["template"] == shape
        assert "record" not in item and item["data"]["title"] == "X"

    case = deck_build.staged_item("NEW:G2", {"content_type": "case_study", "title": "C",
                                             "challenge": "ch", "solution": "so"}, "BFSI")
    assert case["kind"] == "case_study" and case["template"] == "case_study_v2"
    assert case["record"]["challenge"] == "ch" and case["record"]["industry"] == "BFSI"

    # a record with no content_type is a case study (the historical default)
    assert deck_build.staged_item("NEW:G3", {"title": "C"}, "")["kind"] == "case_study"

    # /review edits win over the drafted text
    edited = deck_build.staged_item("NEW:G4", {"content_type": "case_study", "title": "old"},
                                    "", {"title": "new"})
    assert edited["record"]["title"] == "new"


def test_promote_case_saves_once_across_both_commit_points(tmp_path, monkeypatch):
    """A slide downloaded from the builder AND folded into a deck must reach the
    library exactly once -- promote_ai_case mints a fresh id on every call."""
    _chdir()
    staging = _isolate_staging(tmp_path, monkeypatch)
    from deckengine.services.rendering import deck_build
    from deckengine.services.content import case_library

    calls = []
    monkeypatch.setattr(case_library, "promote_ai_case",
                        lambda rec, wt, ind: (calls.append((wt, ind)), "MSS999")[1])

    stg = staging.add({"content_type": "case_study", "title": "A Case", "challenge": "c",
                       "solution": "s", "capabilities": ["One: a"], "results": ["r"]},
                      "MS", "BFSI")
    record = deck_build.staged_item("NEW:" + stg["id"], stg, "BFSI")["record"]

    first = deck_build.promote_case(stg["id"], staging.get(stg["id"]), record, "", "BFSI")
    second = deck_build.promote_case(stg["id"], staging.get(stg["id"]), record, "WORKFORCE", "BFSI")

    assert first == second == "MSS999"
    assert calls == [("MS", "BFSI")]                    # saved once; the record's OWN work type won
    assert staging.get(stg["id"])["promoted_id"] == "MSS999"


def test_promote_case_never_raises(tmp_path, monkeypatch):
    """A library-save failure must never block the deck."""
    _chdir()
    staging = _isolate_staging(tmp_path, monkeypatch)
    from deckengine.services.rendering import deck_build
    from deckengine.services.content import case_library

    def boom(*a, **k):
        raise OSError("Excel is open in another program")
    monkeypatch.setattr(case_library, "promote_ai_case", boom)

    stg = staging.add({"content_type": "case_study", "title": "A"}, "MS", "BFSI")
    assert deck_build.promote_case(stg["id"], stg, {"title": "A"}, "MS", "BFSI") is None
    assert not staging.get(stg["id"]).get("promoted_id")   # not marked -> retried next time


def test_paste_parser_splits_a_document_into_slides():
    """A whole deck's content, marked up by hand, cut into slides. A label that names a
    category outright picks the template; anything else is the slide's HEADING and the
    shape is left to the content (AUTO), never guessed from the title."""
    _chdir()
    from deckengine.services.content import paste_parser as pp

    slides = pp.parse(
        "slide 1 - casestudy\nA bank could not close procurement cycles.\n\n"
        "slide 2 - for box section\nFour pillars: intake, triage, delivery, assurance.\n\n"
        "Slide 3: road map\nPhase 1 discovery. Phase 2 pilot.\n\n"
        "slide 4 - How we think before we build\nThree principles guide us.\n\n"
        "slide 5\nNo label at all.\n")
    assert [s["template"] for s in slides] == [
        "case_study", "four_box", "roadmap_board", pp.AUTO, pp.AUTO]
    assert [s["number"] for s in slides] == [1, 2, 3, 4, 5]
    assert [s["matched"] for s in slides] == [True, True, True, False, False]
    assert slides[0]["content"].startswith("A bank")
    assert slides[1]["content"].endswith("assurance.")   # cut at the NEXT header, not beyond

    # a heading is kept AND prepended to the content, so the slide gets titled what the
    # author called it instead of the generator inventing a title
    assert slides[3]["heading"] == "How we think before we build"
    assert slides[3]["content"].startswith("How we think before we build\n")
    assert slides[4]["heading"] == ""                     # no label -> no heading


def test_paste_parser_has_no_length_limits():
    """A real header is not a short one. "Slide 4: Case Study 1, Reconciliation automation
    with agentic AI" is 64 characters, and a 60-char cap silently swallowed two slides out
    of nine in the owner's first real document. Length is never the signal."""
    _chdir()
    from deckengine.services.content import paste_parser as pp

    long_title = ("An extremely long slide title that goes on at length describing exactly "
                  "what the author intends to convey on this particular slide")
    slides = pp.parse("slide 1 - case study\nbody one\n\nslide 2: %s\nbody two\n" % long_title)
    assert len(slides) == 2
    assert slides[1]["heading"] == long_title

    real = pp.parse("Slide 4: Case Study 1, Reconciliation automation with agentic AI\nbody\n")
    assert len(real) == 1 and real[0]["number"] == 4


def test_paste_parser_uses_slide_numbers_to_reject_stray_lines():
    """Slide numbers validate each other: real headers form an increasing chain. A stray
    line that merely reads like a header can't extend it, so it stays as content."""
    _chdir()
    from deckengine.services.content import paste_parser as pp

    slides = pp.parse("slide 1 - case study\nreal body one\n\n"
                      "slide 2 - road map\nreal body two\n\n"
                      "Slide 1: as discussed above.\nthis is prose, not a slide\n\n"
                      "slide 3 - data table\nreal body three\n")
    assert [s["number"] for s in slides] == [1, 2, 3]         # the out-of-sequence one is gone
    assert "as discussed above" in slides[1]["content"]        # absorbed as content

    # gaps are fine -- increasing, not necessarily consecutive
    assert [s["number"] for s in pp.parse(
        "slide 2 - case study\na\n\nslide 5 - road map\nb\n\nslide 9 - data table\nc\n")] == [2, 5, 9]

    # prose with no separator after the number is never even a candidate
    one = pp.parse("Just one slide.\nSlide 2 of the printed deck covers the roadmap in detail.")
    assert len(one) == 1 and "covers the roadmap" in one[0]["content"]

    # a header with nothing under it is not a slide
    assert len(pp.parse("slide 1 - case study\n\nslide 2 - road map\nreal content")) == 1
    assert pp.parse("") == []


def test_paste_parser_only_matches_a_category_exactly():
    """Fuzzy label matching is gone. The heading "How we would engage with Voya" once
    matched the alias "named list with stats" on the single word "with" -- a preposition
    picked the template. A label either names a category, or it doesn't."""
    _chdir()
    from deckengine.services.content import paste_parser as pp
    assert pp.match_template("Road Map") == "roadmap_board"
    assert pp.match_template("FOUR BOX") == "four_box"
    assert pp.match_template("headline stats") == "stat_overview"
    assert pp.match_template("deep dive") == "pillar_deepdive"
    assert pp.match_template("for box section") == "four_box"
    # headings, not categories -> the content decides, not the title
    assert pp.match_template("How we would engage with Voya") == pp.AUTO
    assert pp.match_template("Case Study 1, Reconciliation automation") == pp.AUTO
    assert pp.match_template("What a first step looks like") == pp.AUTO
    assert pp.match_template("nonsense words") == pp.AUTO
    assert pp.match_template("") == pp.AUTO


def test_builder_parse_route_reports_the_split_without_building(tmp_path, monkeypatch):
    """/builder/parse must be cheap: no slide generation, no staging, no rendering.
    It reports the split and any library duplicate so a bad split is caught first.

    A slide whose category the author NAMED must not cost an AI call at all."""
    _chdir()
    staging = _isolate_staging(tmp_path, monkeypatch)
    import app as appmod
    from deckengine.services.matching import relevance
    from deckengine.services.rendering import slide_generator, preview

    monkeypatch.setattr(relevance, "embed_texts", lambda texts: None)      # offline
    def never(*a, **k):
        raise AssertionError("parse must not build, render, or classify a named category")
    monkeypatch.setattr(slide_generator, "build_content_slide", never)
    monkeypatch.setattr(preview, "case_png", never)
    monkeypatch.setattr(slide_generator, "classify_content_many", never)

    c = appmod.app.test_client()
    d = c.post("/builder/parse", data={
        "work_type": "MS",
        "content": "slide 1 - case study\nA procurement story.\n\nslide 2 - road map\nPhase 1.\n",
    }).get_json()

    assert d["ok"] is True
    assert [s["template"] for s in d["slides"]] == ["case_study", "roadmap_board"]
    assert all(s["matched"] for s in d["slides"])           # the author's own labels won
    assert all(s["matches"] == [] for s in d["slides"])     # offline -> no dup check, no block
    assert staging.all_items() == []                        # nothing staged
    assert "auto" not in {t["key"] for t in d["templates"]}  # every slide has a real shape

    # the same guards as every other builder route
    assert c.post("/builder/parse", data={"content": "x"}).get_json()["ok"] is False   # no work type
    assert c.post("/builder/parse", data={"work_type": "MS", "content": ""}).get_json()["ok"] is False


def test_parse_route_reads_the_shape_of_unlabelled_slides(tmp_path, monkeypatch):
    """The intelligence layer: where the author gave a category we honour it untouched;
    where they gave a heading or nothing, we read the CONTENT and map it to a category
    we actually have. One call for the whole document, and the heading rides along as a
    hint. No slide ever reaches the review screen as 'auto'."""
    _chdir()
    _isolate_staging(tmp_path, monkeypatch)
    import app as appmod
    from deckengine.services.matching import relevance
    from deckengine.services.rendering import slide_generator

    monkeypatch.setattr(relevance, "embed_texts", lambda texts: None)      # offline
    seen = {}
    def fake_classify(slides):
        seen["slides"] = slides
        return ["pillar_deepdive", "box_grid"]
    monkeypatch.setattr(slide_generator, "classify_content_many", fake_classify)

    c = appmod.app.test_client()
    d = c.post("/builder/parse", data={
        "work_type": "MS",
        "content": "slide 1 - case study\nA story.\n\n"
                   "slide 2: How we think before we build\nThree principles.\n\n"
                   "slide 3\nFour reassurances, unlabelled.\n",
    }).get_json()

    assert [s["template"] for s in d["slides"]] == ["case_study", "pillar_deepdive", "box_grid"]
    assert [s["matched"] for s in d["slides"]] == [True, False, False]
    assert [s.get("inferred", False) for s in d["slides"]] == [False, True, True]
    # ONE call, carrying only the slides that needed reading, heading passed as a hint
    assert len(seen["slides"]) == 2
    assert seen["slides"][0]["heading"] == "How we think before we build"
    assert seen["slides"][1]["heading"] == ""


def test_classify_content_many_fails_safe(monkeypatch):
    """Offline, or on a reply that lost count, every slide falls back to the registry's
    first shape -- never a partial mapping silently applied to the wrong slides."""
    _chdir()
    from deckengine.services.rendering import slide_generator
    import openai

    slides = [{"heading": "", "content": "a"}, {"heading": "", "content": "b"}]
    monkeypatch.setattr(openai, "OpenAI", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("offline")))
    assert slide_generator.classify_content_many(slides) == ["case_study", "case_study"]
    assert slide_generator.classify_content_many([]) == []


def test_parse_route_only_flags_duplicates_for_case_study_shapes(tmp_path, monkeypatch):
    """The library holds only case studies, so a roadmap can never be a duplicate of
    one -- offering to swap it in would replace the slide with a different slide."""
    _chdir()
    _isolate_staging(tmp_path, monkeypatch)
    import app as appmod
    from deckengine.services.rendering import slide_generator
    from deckengine.services.matching import dedupe

    monkeypatch.setattr(slide_generator, "classify_content_many",
                        lambda slides: ["case_study"] * len(slides))
    monkeypatch.setattr(dedupe, "similar_cases_many",
                        lambda texts, **k: [[{"id": "MSS001", "title": "T", "work_type": "MS",
                                              "domain": "", "score": 0.95}] for _ in texts])
    c = appmod.app.test_client()
    d = c.post("/builder/parse", data={
        "work_type": "MS",
        "content": "slide 1 - case study\nA story.\n\nslide 2 - road map\nPhase 1.\n"
                   "\nslide 3\nUnlabelled, reads as a case study.\n",
    }).get_json()

    by_num = {s["number"]: s for s in d["slides"]}
    assert by_num[1]["matches"][0]["id"] == "MSS001"        # a real possible duplicate
    assert by_num[1]["matches"][0]["percent"] == 95
    assert by_num[2]["matches"] == []                       # roadmap cannot duplicate a case study
    assert by_num[3]["matches"][0]["id"] == "MSS001"        # inferred case study -> still checked


def test_similar_cases_many_is_one_embedding_call(monkeypatch):
    """N slides must cost ONE embedding call, not N. Blanks keep their own empty slot."""
    _chdir()
    from deckengine.services.matching import dedupe, relevance
    from deckengine.services.content import case_library

    calls = []
    def fake_embed(texts):
        calls.append(list(texts))
        return [[1.0, 0.0] for _ in texts]
    monkeypatch.setattr(relevance, "embed_texts", fake_embed)
    monkeypatch.setattr(relevance, "_load_case_embeddings", lambda: {"MSS1": [1.0, 0.0]})
    monkeypatch.setattr(case_library, "_load",
                        lambda: [{"id": "MSS1", "title": "T", "work_type": "MS", "domain": ""}])

    out = dedupe.similar_cases_many(["alpha", "   ", "beta"], allowed_work_types=["MS"])
    assert len(calls) == 1 and calls[0] == ["alpha", "beta"]   # one call, blanks not sent
    assert out[1] == []                                        # the blank keeps its slot
    assert out[0][0]["id"] == "MSS1" and out[2][0]["id"] == "MSS1"


def test_staging_ids_never_collide(tmp_path, monkeypatch):
    """Ids used to be "G%03d" % (len(items)+1), so a discard freed a number that the
    next add re-used -- two records with one id, and staging.get() returned the wrong
    one. The builder also adds concurrently, which the old scheme couldn't survive."""
    _chdir()
    staging = _isolate_staging(tmp_path, monkeypatch)
    a = staging.add({"title": "A"}, "MS", "")
    b = staging.add({"title": "B"}, "MS", "")
    staging.discard(a["id"])                       # frees G001
    c = staging.add({"title": "C"}, "MS", "")
    assert c["id"] != b["id"]                      # must not collide with the live record
    assert staging.get(b["id"])["title"] == "B"
    assert staging.get(c["id"])["title"] == "C"


def test_slide_schema_covers_every_shape():
    """Every shape the builder can produce must be editable. A shape with no schema
    builds and renders but silently offers no fields."""
    _chdir()
    from deckengine.services.rendering import slide_generator, slide_schema
    for t in slide_generator.CONTENT_TEMPLATES:
        assert slide_schema.fields_for(t["key"]), "no editable fields for %s" % t["key"]


def test_every_shape_is_wired_end_to_end():
    """A slide shape is only usable if ALL of these exist. Half-wiring one means it
    builds and then renders empty, or renders and then can't be edited, or can't be
    named in a pasted document -- each a silent failure the user only sees in the .pptx.

    (case_study is the exception: it fills the branded case_study_v2 template through
    fill_case_study, not a drawn body, so it has no J2W_TEMPLATE frame of its own.)"""
    _chdir()
    from pptx import Presentation
    from deckengine import config
    from deckengine.services.content import paste_parser
    from deckengine.services.rendering import draw_templates, skills, slide_generator, slide_schema

    tags = set()
    for slide in Presentation(config.SKILLS_TEMPLATES_PPTX).slides:
        if slide.has_notes_slide:
            for line in (slide.notes_slide.notes_text_frame.text or "").splitlines():
                if line.strip().startswith("J2W_TEMPLATE:"):
                    tags.add(line.split(":", 1)[1].strip())

    drawn_by_skills = {"four_box", "roadmap_board", "box_grid", "pillar_deepdive",
                       "scored_list", "stat_overview", "data_table"}
    problems = []
    for t in slide_generator.CONTENT_TEMPLATES:
        key = t["key"]
        if key == "case_study":
            continue
        if key not in tags:
            problems.append("%s: no template slide in skills_templates.pptx" % key)
        if not (key in drawn_by_skills or key in draw_templates.DRAWERS):
            problems.append("%s: nothing draws its body" % key)
        if not slide_schema.fields_for(key):
            problems.append("%s: not editable (no schema)" % key)
        try:
            slide_schema.normalize(key, {})
        except Exception as e:
            problems.append("%s: normalizer blows up on an empty record (%s)" % (key, e))
        if not callable(t.get("builder")):
            problems.append("%s: no builder to turn pasted content into it" % key)
    assert not problems, "\n".join(problems)

    # every new shape can be named in a pasted document ("slide 3 - governance layer")
    for key in draw_templates.DRAWERS:
        assert paste_parser.match_template(key.replace("_", " ")) == key, \
            "no paste_parser alias resolves to %s" % key


def test_every_shape_is_reachable_from_every_entry_point():
    """A shape is only 'wired up all over the application' if it appears in ALL of these.
    Miss one and the shape works in the builder but renders as a header on /template/
    recreate, or never shows on /templates, or can't be edited on /review."""
    _chdir()
    from deckengine.services.rendering import draw_templates, recreate, slide_generator, slide_schema
    from deckengine.web import decks as decks_mod
    from deckengine.web.templates import _slide_shapes

    on_templates_page = {s["key"] for s in _slide_shapes()}
    # four_box is marker-filled only -- it has no drawn body in skills.py either
    recreate_known = set(recreate._MAPPING_FNS) | set(recreate._DRAW_FNS)

    problems = []
    for t in slide_generator.CONTENT_TEMPLATES:
        key = t["key"]
        if key not in on_templates_page:
            problems.append("%s: absent from the /templates page" % key)
        if not slide_schema.fields_for(key):
            problems.append("%s: not editable" % key)
        if key != "case_study":
            if key not in recreate_known:
                problems.append("%s: Recreate-with-AI would render only its header" % key)
            if key not in decks_mod._SHAPE_KINDS:
                problems.append("%s: /review can't edit it" % key)
    assert not problems, "\n".join(problems)


def test_recreate_classifier_has_a_letter_per_template():
    """recreate._classify_slide indexed a fixed 11-letter string; the registry grew to 18
    and the whole Recreate-with-AI feature died with IndexError before the AI was even
    called. The alphabet must always outrun the registry."""
    _chdir()
    import string
    from deckengine.services.rendering import slide_generator
    assert len(slide_generator.CONTENT_TEMPLATES) + 1 <= len(string.ascii_uppercase)

    # and building the prompt must not raise, whatever the registry size
    from deckengine.services.rendering import recreate
    key = recreate._classify_slide("four parallel pillars, each with a checklist")
    assert key in {t["key"] for t in slide_generator.CONTENT_TEMPLATES} | {recreate._NONE_KEY}


def test_new_shapes_normalize_and_draw_from_an_empty_record(tmp_path):
    """Every drawer must survive a record with nothing in it -- the AI builder fails safe
    to a placeholder, and a blank slide is far better than a stack trace at download."""
    _chdir()
    from pptx import Presentation
    from deckengine import config
    from deckengine.services.rendering import draw_templates, skills, slide_schema

    tpl = Presentation(config.SKILLS_TEMPLATES_PPTX)
    for kind in draw_templates.DRAWERS:
        rec = slide_schema.normalize(kind, {})
        assert rec["content_type"] == kind
        blank = Presentation()
        blank.slide_width, blank.slide_height = tpl.slide_width, tpl.slide_height
        path = str(tmp_path / (kind + ".pptx"))
        blank.save(path)
        built = skills.build_into(path, [kind],
                                  [{"id": kind, "template": kind, "kind": kind, "data": rec}])
        assert built == 1, "%s drew nothing from an empty record" % kind


def test_slide_schema_edits_round_trip_and_normalize():
    """An edit is re-run through the shape's own normalizer, so a hand-edit can't break
    an invariant the renderer depends on."""
    _chdir()
    from deckengine.services.rendering import slide_schema as ss

    case = {"content_type": "case_study", "industry": "BFSI", "title": "T", "subhead": "s",
            "challenge": "c", "solution": "so", "capabilities": ["A: x"], "results": ["r"]}
    out = ss.apply_edits(case, dict(ss.extract("case_study", case), title="New"))
    assert out["title"] == "New"
    # emptying them must NOT produce a slide the renderer can't fill
    bad = ss.apply_edits(case, {"capabilities": [], "results": [], "title": ""})
    assert len(bad["capabilities"]) == 6 and len(bad["results"]) == 3
    assert bad["title"]                                   # never blank

    roadmap = {"content_type": "roadmap_board", "title": "T", "subhead": "", "intro": "",
               "columns": [{"name": "n", "tag": "P1", "items": ["a"]}], "legend": [],
               "footer_title": "", "footer_body": ""}
    out = ss.apply_edits(roadmap, dict(ss.extract("roadmap_board", roadmap),
                                       columns=[{"name": "Phase 1", "tag": "P1", "items": ["x", "y"]}]))
    assert out["columns"][0]["items"] == ["x", "y"]       # nested list survives the trip


def test_slide_schema_ignores_fields_the_browser_invents():
    """The editor may only touch declared content fields -- never a record's bookkeeping,
    and never its shape."""
    _chdir()
    from deckengine.services.rendering import slide_schema as ss
    rec = {"content_type": "four_box", "id": "G001", "work_type": "MS", "promoted_id": "MSS9",
           "title": "T", "subhead": "s", "boxes": [{"heading": "h", "body": "b"}]}
    out = ss.apply_edits(rec, {"title": "ok", "content_type": "case_study",
                               "id": "HACK", "promoted_id": "STOLEN", "work_type": "AI_POD"})
    assert set(out) == {"title", "subhead", "boxes"}
    assert out["title"] == "ok"


def test_builder_returns_an_editable_slide_card_not_an_image(tmp_path, monkeypatch):
    """A built slide comes back as the SAME editable slide card the review page shows --
    click straight onto the text. Not a picture with an Edit button beside it."""
    _chdir()
    _isolate_staging(tmp_path, monkeypatch)
    import app as appmod
    from deckengine.services.rendering import slide_generator

    monkeypatch.setattr(slide_generator, "build_content_slide", lambda content, industry="", hint="auto": (
        {"content_type": "roadmap_board", "template": "roadmap_board", "title": "Plan",
         "subhead": "", "intro": "",
         "columns": [{"name": "Discovery", "tag": "P1", "items": ["interviews"]}],
         "legend": [], "footer_title": "", "footer_body": ""},
        {"key": "roadmap_board", "label": "Phased roadmap / board"}))

    d = appmod.app.test_client().post("/builder/slide", data={
        "content": "Phase 1 discovery.", "work_type": "MS", "industry": "BFSI"}).get_json()

    assert d["ok"] and d["content_type"] == "roadmap_board"
    assert "png" not in d                                    # no image; the card IS the slide
    html = d["html"]
    assert 'class="se ' in html                              # the review page's editable input
    assert 'data-path="title"' in html
    assert 'data-path="columns.0.name"' in html              # nested record position
    assert 'data-path="columns.0.items.0"' in html           # a list inside a record
    assert "Discovery" in html and "interviews" in html      # current values, not blanks


def test_builder_edit_saves_and_persists(tmp_path, monkeypatch):
    """Saving edits typed onto the card updates the staged record -- the single source of
    truth that the download, the deck, and the library all read."""
    _chdir()
    staging = _isolate_staging(tmp_path, monkeypatch)
    import app as appmod

    stg = staging.add({"content_type": "stat_overview", "template": "stat_overview",
                       "title": "Old", "subhead": "", "intro": "",
                       "stats": [{"value": "42%", "label": "l"}], "items": []}, "MS", "BFSI")
    c = appmod.app.test_client()

    r = c.post("/builder/slide/%s" % stg["id"],
               json={"title": "New", "stats": [{"value": "61%", "label": "faster"}]}).get_json()
    assert r["ok"] and r["title"] == "New"

    back = staging.get(stg["id"])                 # persisted, and bookkeeping untouched
    assert back["title"] == "New" and back["stats"] == [{"value": "61%", "label": "faster"}]
    assert back["work_type"] == "MS" and back["id"] == stg["id"]

    assert c.post("/builder/slide/NOPE", json={}).status_code == 404


def test_review_page_can_now_edit_every_pasted_shape(tmp_path, monkeypatch):
    """The five shapes that used to fall through to "no text to edit", plus the two that
    rendered read-only ("re-paste to change the text"), are all inline-editable now --
    through the same macro, and /finalize applies their edits to the staged record."""
    _chdir()
    staging = _isolate_staging(tmp_path, monkeypatch)
    import app as appmod
    c = appmod.app.test_client()

    stg = staging.add({"content_type": "box_grid", "template": "box_grid", "title": "Old title",
                       "subhead": "s", "boxes": [{"heading": "h1", "body": "b1"},
                                                 {"heading": "h2", "body": "b2"}]}, "MS", "BFSI")
    sid = "NEW:" + stg["id"]

    r = c.post("/review", data={"client_name": "Acme", "order": sid})
    html = r.data.decode()
    assert r.status_code == 200
    assert 'data-shape-id="%s"' % sid in html                 # rendered by the shared macro
    assert 'data-path="boxes.0.heading"' in html
    assert "not inline-editable yet" not in html              # the old read-only footer is gone

    # /finalize applies the edit to the staged record, which is what deck_build reads
    import json as _json
    c.post("/finalize", data={
        "client_name": "Acme", "phase": "Intro", "order": sid, "work_types": "MS",
        "shape__" + sid: _json.dumps({"title": "Edited title",
                                      "boxes": [{"heading": "H1", "body": "B1"}]}),
    })
    back = staging.get(stg["id"])
    assert back["title"] == "Edited title"
    assert back["boxes"][0] == {"heading": "H1", "body": "B1"}
    assert len(back["boxes"]) >= 2            # normalizer re-padded: box_grid needs 2+


def test_draft_with_ai_button_survives_quotes_in_the_gap_text():
    """The button used to be onclick="draftAI({{ m.name|tojson }}, ...)". |tojson emits
    real double quotes, so the HTML attribute ended at the first one and the browser only
    ever saw `onclick="draftAI("` -- every click threw "Unexpected end of input" and did
    nothing. The gap text is arbitrary prose, so it rides on data-* attributes now."""
    _chdir()
    import html as htmlmod
    import re
    from flask import render_template
    import app as appmod

    nasty = 'Catch "micro-cracks" & O\'Brien\'s <defects> before scrap'
    with appmod.app.test_request_context():
        page = render_template(
            "build.html",
            ctx={"client_name": "V", "industry": "BFSI", "transcript": "", "phase": "",
                 "recipient": "", "functions": [], "work_types": ["MS"]},
            picks=[], gaps=[], titles={}, all_slides=[], case_lib=[], suggestions=[],
            suggested=[], ai_used=True, persona_labels=[], rationale=[],
            missing=[{"name": 'Predictive "maintenance"', "description": nasty,
                      "domain": "Manufacturing", "use_case": "ingot line"}],
            research_read=True, research_failed=False, resume=False, build_id="abc",
            reopen_seed=None)

    assert "onclick=\"draftAI" not in page          # no JS source inside an HTML attribute
    assert 'class="btn draft-btn"' in page
    desc = re.search(r'data-desc="([^"]*)"', page)
    assert desc, "the data-desc attribute is not well-formed"
    assert htmlmod.unescape(desc.group(1)) == nasty  # every character survives the trip

    # the card only exists when there IS a gap to draft, and the manual form stays disabled
    assert 'id="ca-preview"' in page and "ca-topic" not in page


def test_static_assets_are_versioned():
    """Without an mtime stamp the browser keeps the old app.css / build.js after a change,
    and the user sees the OLD page against the NEW server. That is exactly how a working
    Draft-with-AI looked broken."""
    _chdir()
    import re
    import app as appmod
    c = appmod.app.test_client()
    page = c.get("/builder").data.decode()
    assert re.search(r'/static/js/builder\.js\?v=\d+', page)
    assert re.search(r'/static/app\.css\?v=\d+', page)
    # and the versioned URL still serves the file
    url = re.search(r'(/static/js/builder\.js\?v=\d+)', page).group(1)
    assert c.get(url).status_code == 200


def test_staging_reads_never_see_a_half_written_file(tmp_path, monkeypatch):
    """`json.dump(x, open(path,"w"))` truncates first, then streams. A reader landing in
    that window got a syntax error, _load() swallowed it and returned [], and the caller
    concluded the record did not exist -- so six of nine concurrent saves 404'd on slides
    that were in the file the whole time. Measured before the fix: 207/300 reads blind."""
    _chdir()
    import threading
    staging = _isolate_staging(tmp_path, monkeypatch)

    for i in range(60):                       # enough records that a write isn't instant
        staging.add({"title": "S%d" % i, "content_type": "case_study",
                     "challenge": "x" * 400, "solution": "y" * 400}, "MS", "BFSI")
    ids = [it["id"] for it in staging.all_items()]
    target = ids[30]

    stop = threading.Event()
    def churn():
        while not stop.is_set():
            staging.update_fields(ids[0], {"title": "churn"})
    writer = threading.Thread(target=churn, daemon=True)
    writer.start()
    try:
        misses = sum(1 for _ in range(400) if staging.get(target) is None)
    finally:
        stop.set()
        writer.join(timeout=2)

    assert misses == 0, "%d reads saw a torn staging file" % misses
    assert len(staging.all_items()) == 60      # and no record was lost to the churn


def test_concurrent_slide_saves_all_succeed(tmp_path, monkeypatch):
    """The builder saves every card at once when you download. Each POST reads the record
    then writes it back; with a non-atomic write most of them 404'd."""
    _chdir()
    staging = _isolate_staging(tmp_path, monkeypatch)
    import app as appmod
    from concurrent.futures import ThreadPoolExecutor

    ids = [staging.add({"content_type": "four_box", "template": "four_box",
                        "title": "T%d" % i, "subhead": "",
                        "boxes": [{"heading": "h", "body": "b" * 300}]}, "MS", "BFSI")["id"]
           for i in range(9)]
    c = appmod.app.test_client()

    def save(sid):
        return c.post("/builder/slide/%s" % sid,
                      json={"title": "Edited " + sid}).status_code

    with ThreadPoolExecutor(max_workers=9) as pool:
        codes = list(pool.map(save, ids))

    assert codes == [200] * 9, "some saves 404'd: %r" % codes
    assert all(staging.get(sid)["title"] == "Edited " + sid for sid in ids)


def test_builder_routes(monkeypatch):
    """The builder page renders, and every endpoint validates its input."""
    _chdir()
    import app as appmod
    from deckengine.services.matching import relevance
    monkeypatch.setattr(relevance, "embed_texts", lambda texts: None)   # offline
    c = appmod.app.test_client()

    assert c.get("/builder").status_code == 200
    # a work type is REQUIRED: without it the case study can't be saved to the library
    assert c.post("/builder/slide", data={"content": "x", "work_type": ""}).get_json()["ok"] is False
    assert c.get("/builder/preview/NOPE001").status_code == 404
    assert c.get("/builder/download?ids=").status_code == 400


def test_concurrent_slide_builds_get_distinct_ids(tmp_path, monkeypatch):
    """The browser fires four /builder/slide requests at once. staging.add() is a
    read-modify-write on one JSON file, so without its lock two of them would mint the
    same id and one slide would silently vanish from the queue."""
    _chdir()
    staging = _isolate_staging(tmp_path, monkeypatch)
    from concurrent.futures import ThreadPoolExecutor

    with ThreadPoolExecutor(max_workers=8) as pool:
        ids = list(pool.map(lambda i: staging.add({"title": "S%d" % i}, "MS", "")["id"],
                            range(24)))

    assert len(set(ids)) == 24                       # no collisions
    assert len(staging.all_items()) == 24            # and nothing was overwritten

