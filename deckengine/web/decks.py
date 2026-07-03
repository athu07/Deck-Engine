# -*- coding: utf-8 -*-
"""
decks.py  --  the core deck-building flow.

/ (or /new) -> /build (pick slides) -> /review (edit) -> /finalize (assemble .pptx),
plus /deck to resume a deck-in-progress held in the browser's deck tray.
"""

import json
import re
import uuid
from zipfile import BadZipFile

from flask import Blueprint, request, render_template, abort
from pptx import Presentation

from deckengine import config
from deckengine.constants import (COVERAGE_THRESHOLD, _GENERIC_NEEDS, INDUSTRIES,
                                  FUNCTIONS, PHASES)
from deckengine.services.matching import matcher, relevance, ai_matcher
from deckengine.services.rendering import skills, staging, deck_build
from deckengine.services.content import case_library, editor
from deckengine.services.content.build_library import read_id
from deckengine.services.rendering import assembler
from deckengine.services import ingest as research
from deckengine.services import meeting_log
from .view_helpers import (shell, file_busy_page, safe_filename, current_salesperson,
                           legacy_case_ids as _legacy_case_ids)

bp = Blueprint("decks", __name__)


@bp.route("/")
@bp.route("/new")
def home():
    try:
        lib_count = len(json.load(open(config.TAGGED_LIBRARY_JSON, encoding="utf-8")))
    except Exception:
        lib_count = 0
    body = render_template("new_form.html", industries=INDUSTRIES, functions=FUNCTIONS,
                                  phases=PHASES, library_count=lib_count, error="")
    return shell(body, active="new", crumb="<b>New deck</b> / Context")


@bp.route("/deck")
def deck_resume():
    """Re-open the deck-in-progress (held in the browser's deck tray). The list and
    context are hydrated client-side from localStorage; the server just supplies
    the slide catalogue."""
    titles = matcher._title_lookup()
    all_slides = sorted(((sid, t) for sid, t in titles.items()
                         if sid not in _legacy_case_ids()),
                        key=lambda kv: matcher._num(kv[0]))
    titles.update(case_library.title_map())   # so resumed store cases show their title
    case_lib = sorted(case_library.all_cases(), key=lambda c: (c["work_type"], c["title"]))
    empty_ctx = {"client_name": "", "industry": "", "transcript": "",
                 "phase": "", "recipient": "", "functions": [], "work_types": []}
    body = render_template("build.html", ctx=empty_ctx, picks=[], gaps=[],
                                  titles=titles, all_slides=all_slides, case_lib=case_lib,
                                  suggestions=[], suggested=[], ai_used=False,
                                  persona_labels=[], rationale=[], missing=[],
                                  research_read=False, research_failed=False,
                                  resume=True, build_id="")
    return shell(body, active="new", crumb="<b>New deck</b> / Your deck")


@bp.route("/build", methods=["POST"])
def build():
    ctx = {
        "client_name": request.form.get("client_name", "Client").strip(),
        "industry": request.form.get("industry", "").strip(),
        "work_types": request.form.getlist("work_types"),
        "functions": request.form.getlist("functions"),
        "phase": request.form.get("phase", "").strip(),
        "recipient": request.form.get("recipient", "").strip(),
        "salesperson": current_salesperson(),
        "transcript": request.form.get("transcript", "").strip(),
    }
    # optional deep-research file (PDF/text) -> read alongside the notes for matching
    research_text = research.extract_text(request.files.get("research_file"))
    research_given = bool(request.files.get("research_file") and
                          getattr(request.files.get("research_file"), "filename", ""))
    research_read = bool(research_text)                 # actually got text out of it
    research_failed = research_given and not research_read
    # optional STAKEHOLDER PROFILE (LinkedIn/bio) -> drives function/skill matching
    profile_text = research.extract_text(request.files.get("profile_file"))
    match_notes = ctx["transcript"]
    if research_text:
        match_notes = (ctx["transcript"] + "\n\n[DEEP RESEARCH BRIEF]\n"
                       + research_text).strip()
    # Backstop for the browser's at-least-one-work-type check (e.g. JS disabled).
    if not ctx["work_types"]:
        try:
            lib_count = len(json.load(open(config.TAGGED_LIBRARY_JSON, encoding="utf-8")))
        except Exception:
            lib_count = 0
        body = render_template("new_form.html", industries=INDUSTRIES, functions=FUNCTIONS,
                                      phases=PHASES, library_count=lib_count,
                                      error="Please select at least one work type.")
        return shell(body, active="new", crumb="<b>New deck</b> / Context")
    # The DEEP RESEARCH (or the notes) names the account's real interests. Extract
    # them, then split by coverage: each we HAVE a case for -> a PRIORITY pick
    # (pinned ahead of generic mail-thread matches); each we DON'T -> a build gap.
    # The account's needs come from BOTH the research brief AND the stakeholder's
    # profile (their function / current-role mandate) — balanced. For each need,
    # find OUR best case OF THE SELECTED WORK TYPES via a DIRECT skill->title match:
    # covered -> that case LEADS the deck; not covered -> a gap ("want to generate").
    priority_ids, missing = [], []
    profile_needs = []
    try:
        wanted = {w.upper() for w in ctx["work_types"]}
        wt_ids = {c["id"] for c in case_library.all_cases()
                  if c.get("work_type", "").upper() in wanted}
        research_needs = ai_matcher.extract_accelerators(research_text) if research_text else []
        profile_needs = ai_matcher.extract_profile(profile_text) if profile_text else []
        if not research_needs and not profile_needs and ctx["transcript"].strip():
            research_needs = ai_matcher.extract_accelerators(ctx["transcript"])
        needs = profile_needs + research_needs          # profile + research, balanced
        acct_fns = matcher._account_functions(set(ctx.get("functions", [])), match_notes)
        if needs:
            # match on the skill NAME (the crisp function term), not the prose
            # description (which adds generic words that match random cases)
            best = relevance.best_cases([a["name"] for a in needs],
                                        industry=ctx.get("industry", ""), functions=acct_fns,
                                        allowed_ids=wt_ids)
            for a, (bid, _adj, cos, thits) in zip(needs, best):
                if a["name"].strip().lower() in _GENERIC_NEEDS:
                    continue                                        # too generic to pick or flag
                covered = thits >= 1 or cos >= COVERAGE_THRESHOLD   # skill in a TITLE, or strong meaning
                if covered:
                    if bid and bid not in priority_ids:
                        priority_ids.append(bid)
                else:
                    missing.append({**a, "sim": round(cos, 2)})
            missing = missing[:6]
    except Exception:
        priority_ids, missing = [], []

    # research + the profile's crisp focus areas LEAD the ranking; mail is secondary
    lead_research = (research_text + "\n"
                     + " ".join(f"{n['name']}. {n['description']}" for n in profile_needs)).strip()
    result = matcher.plan({**ctx, "transcript": ctx["transcript"], "research": lead_research},
                          use_ai=True, priority_ids=priority_ids)
    # Gaps are FLAGS only now (no inline generation) — nothing to pre-fill here.
    titles = matcher._title_lookup()

    # --- skills slides (PURE Workforce only): auto-add after the standard block,
    #     before the case studies; labeled + removable in the panel ---
    sk = skills.candidates(ctx)
    if sk:
        picks = result["picks"]
        insert_at = next((i for i, p in enumerate(picks) if p["reason"].startswith("case")), None)
        if insert_at is None:
            insert_at = next((i for i, p in enumerate(picks) if p["slide_id"] in matcher.PIN_TO_END),
                             len(picks))
        sk_picks = []
        _chip = {"industry_strength": "IND", "skill_deepdive": "SKL",
                 "company_footprint": "FOOT"}
        for c in sk:
            titles[c["id"]] = c["label"]
            reason = {
                "industry_strength": "auto-added — RFI: industry strength slide",
                "skill_deepdive":    "auto-added — RFI: skills deployed slide",
                "company_footprint": "auto-added — RFI: existing client relationship",
            }.get(c["kind"], "auto-added — RFI data slide")
            sk_picks.append({"slide_id": c["id"], "reason": reason, "skill": True,
                             "tag": _chip.get(c["kind"], "SKL"),
                             "label": c["label"]})
        picks[insert_at:insert_at] = sk_picks
    all_slides = sorted(((sid, t) for sid, t in titles.items()
                         if sid not in _legacy_case_ids()),
                        key=lambda kv: matcher._num(kv[0]))
    # store-case titles power the picks/suggested display AND the JS title lookup
    # (so a manually-added or auto-picked AIP/WFS/MSS case shows its real title)
    titles.update(case_library.title_map())
    case_lib = sorted(case_library.all_cases(), key=lambda c: (c["work_type"], c["title"]))

    # --- "why this deck matches" rationale (deterministic, straight from picks) ---
    def _why(reason):
        r = re.sub(r"^case \[[^\]]*\] · (T\d · )?", "", reason or "")
        r = re.sub(r"\s*\(score [^)]*\)$", "", r)
        return r.strip() or "related to your notes"

    def _tier(reason):
        m = re.search(r"\bT(\d)\b", reason or "")
        return "T" + m.group(1) if m else ""

    rationale = [{"id": p["slide_id"], "title": titles.get(p["slide_id"], ""),
                  "why": _why(p["reason"]), "tier": _tier(p["reason"])}
                 for p in result["picks"] if p["slide_id"][:3] in ("AIP", "WFS", "MSS")]

    # a stakeholder-specific "why this resonates with THEM" line per case (AI).
    # We ALWAYS give it the account context (client/industry/role/work type) so a
    # reason is written for EVERY case even with no notes/research/profile; the
    # profile, extracted focus areas and notes enrich it when present.
    account_bits = [
        "CLIENT: " + (ctx.get("client_name") or "the client"),
        "INDUSTRY: " + (ctx.get("industry") or "unspecified"),
        "WORK TYPES: " + (", ".join(ctx.get("work_types") or []) or "unspecified"),
    ]
    if ctx.get("functions"):
        account_bits.append("FUNCTIONS: " + ", ".join(ctx["functions"]))
    if ctx.get("recipient"):
        account_bits.append("STAKEHOLDER ROLE: " + ctx["recipient"])
    person_ctx = "\n\n".join(x for x in [
        "ACCOUNT:\n" + "\n".join(account_bits),
        ("STAKEHOLDER PROFILE:\n" + profile_text[:4000]) if profile_text else "",
        ("THEIR FUNCTION/SKILLS: " + ", ".join(n["name"] for n in profile_needs)) if profile_needs else "",
        ("MEETING NOTES:\n" + match_notes[:3000]) if match_notes.strip() else "",
    ] if x)
    if rationale:
        _rec = {r["id"]: r for r in case_library._load()}
        picks_for_ai = [{"id": r["id"], "title": r["title"],
                         "blurb": (_rec.get(r["id"], {}).get("challenge", "") or "")[:160]}
                        for r in rationale][:12]
        try:
            fit = ai_matcher.explain_fit(person_ctx, ctx.get("recipient", ""), picks_for_ai)
        except Exception:
            fit = {}
        # bare, low-signal fallbacks the AI sentence should replace outright
        _bare = ("same industry", "related", "related to your notes")
        for r in rationale:
            if fit.get(r["id"]):
                r["fit"] = fit[r["id"]]
                if r.get("why", "").strip().lower() in _bare:
                    r["why"] = ""      # AI sentence carries it; hide the bare fallback

    # (priority picks + `missing` were computed before planning, above)
    body = render_template("build.html", ctx=ctx, picks=result["picks"],
                                  gaps=result["gaps"], titles=titles, all_slides=all_slides,
                                  case_lib=case_lib, rationale=rationale, missing=missing,
                                  research_read=research_read, research_failed=research_failed,
                                  suggestions=result.get("suggestions", []),
                                  suggested=result.get("suggested", []),
                                  ai_used=result.get("ai_used", False),
                                  persona_labels=result.get("persona_labels", []),
                                  resume=False, build_id=uuid.uuid4().hex)
    return shell(body, active="new", crumb="<b>New deck</b> / Suggested slides")


@bp.route("/review", methods=["POST"])
def review():
    client = request.form.get("client_name", "Client").strip()
    ids = [x for x in request.form.get("order", "").split(",") if x]
    industry = request.form.get("industry", "")
    transcript = request.form.get("transcript", "")
    prs = Presentation(assembler.SOURCE)
    by_id = {read_id(s): s for s in prs.slides if read_id(s)}

    # ---- existing library slides: editable title/subtitle (as before) ----
    cards = []
    for sid in ids:
        slide = by_id.get(sid)
        if not slide:
            continue
        fields = [(idx, label, text.replace("[CLIENT]", client))
                  for idx, label, text in editor.editable_fields(slide)]
        shown = {f[2] for f in fields}
        context = "\n".join(t.replace("[CLIENT]", client)
                            for t in editor.full_text(slide) if t not in shown)
        cards.append({"id": sid, "fields": fields, "context": context[:400]})

    # Gaps are FLAGS only now — they are never auto-written here. AI slides come
    # solely from the deliberate "Create a slide with AI" tool (NEW:<id> items).
    ai_cards, ai_ids = [], []

    body = render_template("review.html", client=client, order=",".join(ids),
                                  cards=cards, ai_cards=ai_cards, ai_ids=",".join(ai_ids),
                                  industry=industry, transcript=transcript,
                                  phase=request.form.get("phase", ""),
                                  recipient=request.form.get("recipient", ""),
                                  functions=request.form.getlist("functions"),
                                  work_types=request.form.getlist("work_types"))
    return shell(body, active="new", crumb="<b>New deck</b> / Review &amp; edit")


@bp.route("/finalize", methods=["POST"])
def finalize():
    client = request.form.get("client_name", "Client").strip()
    ids = [x for x in request.form.get("order", "").split(",") if x]
    if not ids:
        abort(400)
    import os
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    filename = f"Tailored_Deck_{safe_filename(client)}.pptx"
    path = os.path.join(config.OUTPUT_DIR, filename)

    # ---- AI slides: apply edits, then ACCEPT (promote -> library) or REJECT ----
    final_ids = list(ids)
    accepted = []
    for gid in [x for x in request.form.get("ai_ids", "").split(",") if x]:
        decision = request.form.get("ai_decision__" + gid, "accept")
        if decision == "reject":
            staging.discard(gid)
            continue
        staging.update_content(
            gid,
            title=request.form.get("ai_title__" + gid),
            keywords=request.form.get("ai_keywords__" + gid),
            bullets=[b.strip() for b in request.form.get("ai_bullets__" + gid, "").splitlines() if b.strip()],
        )
        new_cs = staging.promote(gid)          # full sign-off -> real, client-ready slide
        if new_cs:
            accepted.append(new_cs)
    # Slot the AI slides in BEFORE the closing slides (Next Steps / Let's win together),
    # not at the very end of the deck.
    if accepted:
        insert_at = next((i for i, s in enumerate(final_ids) if s in matcher.PIN_TO_END),
                         len(final_ids))
        final_ids[insert_at:insert_at] = accepted

    # collect inline title/subtitle edits from the review form
    edits = {}
    for key, val in request.form.items():
        if key.startswith("edit__"):
            _, sid, idx = key.split("__")
            edits.setdefault(sid, {})[int(idx)] = val

    try:
        deck_build.assemble(final_ids, path, client=client,
                            industry=request.form.get("industry", ""),
                            work_types=request.form.getlist("work_types"),
                            transcript=request.form.get("transcript", ""),
                            edits=edits)
    except (PermissionError, BadZipFile) as e:
        return file_busy_page(e)
    meeting_log.save(                          # auto-log this meeting (no extra step)
        client=client,
        industry=request.form.get("industry", ""),
        functions=request.form.getlist("functions"),
        work_types=request.form.getlist("work_types"),
        phase=request.form.get("phase", ""),
        recipient=request.form.get("recipient", ""),
        salesperson=current_salesperson(),
        slide_ids=final_ids,
        deck_file=filename,
    )
    count = len(list(Presentation(path).slides))
    body = render_template("preview.html", client=client, filename=filename, count=count)
    return shell(body, active="new", crumb="<b>New deck</b> / Done")
