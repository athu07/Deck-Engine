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
from deckengine.constants import (COVERAGE_THRESHOLD, CAPABILITY_COVER, _GENERIC_NEEDS,
                                  INDUSTRIES, FUNCTIONS, PHASES)
from deckengine.services.matching import matcher, relevance, ai_matcher
from deckengine.services.rendering import skills, staging, deck_build, fill_case_study
from deckengine.services.content import case_library, editor
from deckengine.services.content.content_store import content_store
from deckengine.services.content.build_library import read_id
from deckengine.services.rendering import assembler
from deckengine.services import ingest as research
from deckengine.services import meeting_log
from deckengine.services import build_context
from .view_helpers import (shell, file_busy_page, safe_filename, current_salesperson,
                           legacy_case_ids as _legacy_case_ids)

bp = Blueprint("decks", __name__)


def _is_avoided(rec, avoid):
    """True if this content-store case is ABOUT a mismatch-flagged capability (so it must
    not be picked to answer a need)."""
    if not avoid or not rec:
        return False
    kw = rec.get("keywords") or []
    text = rec.get("title", "") + " " + (" ".join(kw) if isinstance(kw, list) else str(kw))
    head = relevance._tokens(text)
    for a in avoid:
        cap = a.get("capability", "") if isinstance(a, dict) else str(a)
        terms = relevance.specific_terms(cap)
        if terms and len(terms & head) >= max(1, len(terms) // 2):
            return True
    return False


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
    # Persist this build's full context (deep research + profile + FULL transcript),
    # keyed by build_id, so the AI case-study generator can synthesise from it later.
    build_id = uuid.uuid4().hex
    build_context.save(build_id, {
        "research": research_text, "profile": profile_text,
        "transcript": ctx["transcript"], "industry": ctx["industry"],
        "recipient": ctx["recipient"], "functions": ctx["functions"],
        "client_name": ctx["client_name"],
    })
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
    profile_needs = []               # needs (name+description) — feeds lead_research
    avoid = []                       # the brief's mismatch flags
    priority_reasons = {}            # case_id -> {"why","signal"} for the rationale
    try:
        wanted = {w.upper() for w in ctx["work_types"]}
        wt_ids = {c["id"] for c in case_library.all_cases()
                  if c.get("work_type", "").upper() in wanted}
        acct_fns = matcher._account_functions(set(ctx.get("functions", [])), match_notes)
        # ONE structured pass over research + profile + transcript: needs (with domain/
        # use-case), mismatch flags, expressed interest, account (honours the reference
        # priority: research primary, transcript = prior interest, profile-only otherwise).
        brief = ai_matcher.extract_brief(research_text, profile_text, ctx["transcript"])
        if brief and brief.get("needs"):
            avoid = brief.get("avoid") or []
            expressed = brief.get("expressed_interest") or []
            needs = brief["needs"]
            profile_needs = needs                       # names+descriptions for lead_research
            # expressed-interest topics -> lightweight needs so PRIOR INTEREST is weighted
            seen_names = {n["name"].lower() for n in needs}
            all_needs = needs + [{"name": x, "description": "", "domain": "", "use_case": ""}
                                 for x in expressed if x.lower() not in seen_names]
            # cheap shortlist per need — query is the bare CAPABILITY NAME (a crisp capability
            # embeds best; adding the use-case sentence or the client industry drags the match
            # toward the wrong cases). Industry is applied separately as a light boost.
            queries = [n["name"] for n in all_needs]
            shortlists = relevance.shortlist_cases(queries, industry=ctx.get("industry", ""),
                                                   functions=acct_fns, allowed_ids=wt_ids, top_n=8)
            sl_by_name = {n["name"]: lst for n, lst in zip(all_needs, shortlists)}
            recs = {r["id"]: r for r in case_library._load()}
            expressed_lc = {x.lower() for x in expressed}
            # SELECTION is algorithmic (the semantic shortlist ranks capability reliably): take
            # the top candidate that is NOT mismatch-flagged and clears the coverage bar.
            # covered -> priority pick; nothing clears the bar -> an honest gap.
            picked = []                       # [{need,id,title,blurb}] for the explainer
            for n in all_needs:
                name = n["name"]
                if name.strip().lower() in _GENERIC_NEEDS:
                    continue
                cand = None
                for item in sl_by_name.get(name, []):
                    if item["id"] in priority_ids:
                        continue                        # already used for another need
                    if _is_avoided(recs.get(item["id"], {}), avoid):
                        continue                        # skip a mismatch-flagged case
                    cand = item
                    break
                if cand and (cand["cosine"] >= CAPABILITY_COVER or cand["title_hits"] >= 1):
                    if cand["id"] not in priority_ids:
                        priority_ids.append(cand["id"])
                        rc = recs.get(cand["id"], {})
                        picked.append({"need": name, "id": cand["id"], "title": rc.get("title", ""),
                                       "blurb": rc.get("challenge", "")})
                elif name.lower() not in expressed_lc and n.get("description"):
                    missing.append({"name": name, "description": n["description"],
                                    "sim": round(cand["cosine"], 2) if cand else 0.0})
            missing = missing[:6]
            # rich, honest "why this was picked" line per pick (one LLM call; the model only
            # EXPLAINS the already-chosen case, so it cannot mis-pick). Templated fallback.
            ex = ai_matcher.explain_picks(brief, picked)
            for it in picked:
                r = ex.get(it["id"])
                if r and r.get("reason"):
                    priority_reasons[it["id"]] = {"why": r["reason"], "signal": r.get("signal", "")}
                else:
                    priority_reasons[it["id"]] = {"why": "Proves " + it["need"].lower(),
                                                  "signal": "capability"}
        else:
            # fail-safe: the old name-only extraction path (offline / brief parse failed)
            research_needs = ai_matcher.extract_accelerators(research_text) if research_text else []
            profile_needs = ai_matcher.extract_profile(profile_text) if profile_text else []
            if not research_needs and not profile_needs and ctx["transcript"].strip():
                research_needs = ai_matcher.extract_accelerators(ctx["transcript"])
            needs = profile_needs + research_needs
            if needs:
                best = relevance.best_cases([a["name"] for a in needs],
                                            industry=ctx.get("industry", ""), functions=acct_fns,
                                            allowed_ids=wt_ids)
                for a, (bid, _adj, cos, thits) in zip(needs, best):
                    if a["name"].strip().lower() in _GENERIC_NEEDS:
                        continue
                    covered = thits >= 1 or cos >= COVERAGE_THRESHOLD
                    if covered:
                        if bid and bid not in priority_ids:
                            priority_ids.append(bid)
                    else:
                        missing.append({**a, "sim": round(cos, 2)})
                missing = missing[:6]
    except Exception:
        priority_ids, missing, avoid, priority_reasons = [], [], [], {}

    # research + the profile's crisp focus areas LEAD the ranking; mail is secondary
    lead_research = (research_text + "\n"
                     + " ".join(f"{n['name']}. {n.get('description','')}" for n in profile_needs)).strip()
    result = matcher.plan({**ctx, "transcript": ctx["transcript"], "research": lead_research},
                          use_ai=True, priority_ids=priority_ids, avoid=avoid)
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

    # "Why this resonates" per case: the LLM re-rank already justified each PRIORITY
    # pick (domain / prior-interest / role), so surface that as the fit line. Fill
    # cases (not tied to a specific need) keep their deterministic matcher reason.
    _bare = ("same industry", "related", "related to your notes")
    for r in rationale:
        pr = priority_reasons.get(r["id"])
        if pr and pr.get("why"):
            r["fit"] = pr["why"]
            if pr.get("signal"):
                r["signal"] = pr["signal"]
            if r.get("why", "").strip().lower() in _bare:
                r["why"] = ""          # the re-rank reason carries it; hide the bare fallback

    # (priority picks + `missing` were computed before planning, above)
    body = render_template("build.html", ctx=ctx, picks=result["picks"],
                                  gaps=result["gaps"], titles=titles, all_slides=all_slides,
                                  case_lib=case_lib, rationale=rationale, missing=missing,
                                  research_read=research_read, research_failed=research_failed,
                                  suggestions=result.get("suggestions", []),
                                  suggested=result.get("suggested", []),
                                  ai_used=result.get("ai_used", False),
                                  persona_labels=result.get("persona_labels", []),
                                  resume=False, build_id=build_id)
    return shell(body, active="new", crumb="<b>New deck</b> / Suggested slides")


@bp.route("/review", methods=["POST"])
def review():
    client = request.form.get("client_name", "Client").strip()
    ids = [x for x in request.form.get("order", "").split(",") if x]
    industry = request.form.get("industry", "")
    transcript = request.form.get("transcript", "")
    prs = Presentation(assembler.SOURCE)
    by_id = {read_id(s): s for s in prs.slides if read_id(s)}
    store = content_store()          # {id: record} for AIP/WFS/MSS store cases

    def _case_slide(sid, rec):
        """Editable slide data for a case study (content-store or AI-drafted)."""
        domain = rec.get("domain") or rec.get("industry") or ""
        function = rec.get("function") or ""
        sub = rec.get("subhead") or ""
        if (not domain or not function) and sub:      # AI drafts carry a 'subhead' string
            dm = re.search(r"Domain:\s*([^|]+)", sub)
            fm = re.search(r"Function:\s*([^|]+)", sub)
            domain = domain or (dm.group(1).strip() if dm else "")
            function = function or (fm.group(1).strip() if fm else "")
        caps = fill_case_study._pad(rec.get("capabilities", []), 6)
        return {"kind": "case", "id": sid, "title": rec.get("title", ""),
                "domain": domain, "function": function,
                "challenge": rec.get("challenge", ""), "solution": rec.get("solution", ""),
                "caps": [fill_case_study.split_capability(c) for c in caps],
                "results": fill_case_study._pad(rec.get("results", []), 3)}

    # Build one editable "slide" per id, in deck order: library slides (title/subtitle),
    # case studies (full body editable), and data-driven skills slides (read-only preview).
    slides = []
    for sid in ids:
        if sid in by_id:                                       # master library slide (CSxx)
            fields = [(idx, label, text.replace("[CLIENT]", client))
                      for idx, label, text in editor.editable_fields(by_id[sid])]
            slides.append({"kind": "library", "id": sid, "fields": fields})
        elif sid in store:                                     # content-store case
            slides.append(_case_slide(sid, store[sid]))
        elif sid.startswith("NEW:"):                           # AI-created case
            rec = staging.get(sid[4:])
            if rec:
                slides.append(_case_slide(sid, rec))
        elif sid.startswith("SK:") or sid.startswith("FP:"):   # data-driven skills slide
            slides.append({"kind": "skills", "id": sid})

    body = render_template("review.html", client=client, order=",".join(ids),
                                  slides=slides, industry=industry, transcript=transcript,
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

    # collect inline edits from the review slide-view: edit__ = library title/subtitle,
    # cs__ = case-study body fields (title/domain/function/challenge/solution/caps/results).
    edits = {}
    raw_cases = {}
    for key, val in request.form.items():
        if key.startswith("edit__"):
            _, sid, idx = key.split("__")
            edits.setdefault(sid, {})[int(idx)] = val
        elif key.startswith("cs__"):
            parts = key.split("__")
            if len(parts) != 3:
                continue
            _, sid, field = parts
            d = raw_cases.setdefault(sid, {"caps": {}, "results": {}})
            if field in ("title", "domain", "function", "challenge", "solution"):
                d[field] = val
            elif field.startswith("cap") and field.endswith("_t"):
                d["caps"].setdefault(int(field[3:-2]), {})["t"] = val
            elif field.startswith("cap") and field.endswith("_b"):
                d["caps"].setdefault(int(field[3:-2]), {})["b"] = val
            elif field.startswith("res"):
                d["results"][int(field[3:])] = val
    case_edits = {}
    for sid, d in raw_cases.items():
        ov = {k: d[k] for k in ("title", "domain", "function", "challenge", "solution") if k in d}
        caps = []
        for i in sorted(d["caps"]):
            t = (d["caps"][i].get("t") or "").strip()
            b = (d["caps"][i].get("b") or "").strip()
            if t or b:
                caps.append({"title": t, "body": b})
        if caps:
            ov["capabilities"] = caps
        results = [(d["results"][i] or "").strip() for i in sorted(d["results"])]
        results = [r for r in results if r]
        if results:
            ov["results"] = results
        if ov:
            case_edits[sid] = ov

    try:
        deck_build.assemble(final_ids, path, client=client,
                            industry=request.form.get("industry", ""),
                            work_types=request.form.getlist("work_types"),
                            transcript=request.form.get("transcript", ""),
                            edits=edits, case_edits=case_edits)
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
