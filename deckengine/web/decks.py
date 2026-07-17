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

from flask import Blueprint, request, render_template, abort, jsonify
from pptx import Presentation

from deckengine import config
from deckengine import constants
from deckengine.constants import (COVERAGE_THRESHOLD, CAPABILITY_COVER, _GENERIC_NEEDS,
                                  INDUSTRIES, FUNCTIONS, PHASES)
from deckengine.services.matching import matcher, relevance, ai_matcher
from deckengine.services.matching.tagger import INDUSTRY as _BUILTIN_INDUSTRIES
from deckengine.services.rendering import (skills, staging, deck_build, fill_case_study,
                                           slide_generator, slide_schema, client_context)
from deckengine.services.content import case_library, editor
from deckengine.services.content.content_store import content_store
from deckengine.services.content.build_library import read_id
from deckengine.services.rendering import assembler
from deckengine.services import ingest as research
from deckengine.services import meeting_log
from deckengine.services import build_context
from .view_helpers import (shell, file_busy_page, current_salesperson,
                           legacy_case_ids as _legacy_case_ids,
                           resolve_industry, remember_custom_industry)

bp = Blueprint("decks", __name__)

# the pasted shapes that render through the shared inline-editor macro on /review
# (a case study has its own richer card; library slides have their own title/subtitle form)
_SHAPE_KINDS = tuple(k for k in slide_schema.SCHEMA if k != "case_study")


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
    body = render_template("new_form.html", industries=constants.all_industries(), functions=FUNCTIONS,
                                  phases=PHASES, library_count=lib_count, error="",
                                  content_templates=slide_generator.CONTENT_TEMPLATES)
    return shell(body, active="new", crumb="<b>New deck</b> / Context")


def _resume_page(reopen_seed=None):
    """Shared render for /deck (empty browser-tray resume) and /deck/reopen
    (server-persisted version reload) — both hydrate the SAME client-side
    resume mechanism (build.js reads localStorage j2w_deck); reopen_seed just
    pre-seeds that storage before build.js's init() runs (see build.html)."""
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
                                  resume=True, build_id="", reopen_seed=reopen_seed)
    return shell(body, active="new", crumb="<b>New deck</b> / Your deck")


@bp.route("/deck")
def deck_resume():
    """Re-open the deck-in-progress (held in the browser's deck tray). The list and
    context are hydrated client-side from localStorage; the server just supplies
    the slide catalogue."""
    return _resume_page()


@bp.route("/deck/reopen")
def deck_reopen():
    """Re-open an already-finalized, already-downloaded deck for further editing
    (add/remove/reorder a slide, e.g. one more case study), continuing from its
    LATEST version. Finalizing again produces the next version — nothing is
    overwritten. If no version exists yet for this client+phase, falls back to
    a plain empty resume (same as /deck)."""
    client = request.args.get("client", "").strip()
    phase = request.args.get("phase", "").strip()
    rec, latest = meeting_log.latest_version(client, phase)
    if not latest:
        return _resume_page()
    seed = {
        "ctx": {
            "client_name": client,
            "industry": rec.get("industry", ""),
            "transcript": "",           # not persisted per-version; re-enter if needed
            "phase": phase,
            "recipient": latest.get("recipient", ""),
            "functions": latest.get("functions", []),
            "work_types": latest.get("work_types", []),
        },
        "order": latest.get("slide_ids", []),
        "active": True,
    }
    return _resume_page(reopen_seed=seed)


@bp.route("/build", methods=["POST"])
def build():
    ctx = {
        "client_name": request.form.get("client_name", "Client").strip(),
        # "Other…" -> the typed industry, never the raw sentinel (see resolve_industry)
        "industry": resolve_industry(request.form),
        "work_types": request.form.getlist("work_types"),
        "functions": request.form.getlist("functions"),
        "phase": request.form.get("phase", "").strip(),
        "recipient": request.form.get("recipient", "").strip(),
        "salesperson": current_salesperson(),
        "transcript": request.form.get("transcript", "").strip(),
    }
    # a free-typed "Other" industry (not in the built-in taxonomy) is remembered
    # so it shows up in the dropdown for every build after this one
    remember_custom_industry(ctx["industry"])
    # optional deep-research file (PDF/text) -> read alongside the notes for matching
    research_text = research.extract_text(request.files.get("research_file"))
    research_given = bool(request.files.get("research_file") and
                          getattr(request.files.get("research_file"), "filename", ""))
    research_failed = research_given and not research_text
    # optional STAKEHOLDER PROFILE (LinkedIn/bio) -> drives function/skill matching
    profile_text = research.extract_text(request.files.get("profile_file"))
    # optional AUTO-RESEARCH (the strategic brief from /research_account) -- the
    # salesperson reviewed/edited it client-side before submitting. Combined
    # ADDITIVELY with any uploaded research file text, never replacing it
    # (owner's spec, 2026-07-08: keep the manual upload option, add
    # auto-research alongside it).
    auto_brief = request.form.get("auto_company_text", "").strip()
    if auto_brief:
        research_text = (research_text + "\n\n" + auto_brief).strip() if research_text else auto_brief
    research_read = bool(research_text)                 # got text, from file and/or auto-research
    # optional CLIENT LOGO -> stamped onto the title slide next to the J2W wordmark
    # at finalize time (client_logo.stamp_into, called from deck_build.assemble).
    # Saved now (keyed by client name) so it survives all the way to /finalize
    # without needing new state threaded through /review's form. A manual
    # upload wins if both are present; otherwise, if "Find logo" was clicked
    # and confirmed a real image (see /logo_preview), its bytes ride along as
    # a data URI so /build uses EXACTLY what was previewed, not a re-rolled
    # search result.
    logo_file = request.files.get("client_logo_file")
    if logo_file and getattr(logo_file, "filename", ""):
        from deckengine.services.rendering import client_logo
        client_logo.save(ctx["client_name"], logo_file.read())
    else:
        data_uri = request.form.get("client_logo_data_uri", "").strip()
        if data_uri.startswith("data:") and "," in data_uri:
            import base64
            from deckengine.services.rendering import client_logo
            try:
                client_logo.save(ctx["client_name"], base64.b64decode(data_uri.split(",", 1)[1]))
            except Exception:
                pass
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
        body = render_template("new_form.html", industries=constants.all_industries(), functions=FUNCTIONS,
                                      phases=PHASES, library_count=lib_count,
                                      error="Please select at least one work type.",
                                      content_templates=slide_generator.CONTENT_TEMPLATES)
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
    brief = {}                       # set below; kept defined even if extraction fails,
                                     # so the (separate) strategic-inference pass can use it
    # shared prep, hoisted above both the literal-extraction and strategic-inference
    # passes below (pure/no I/O, safe either way needs it)
    wanted = {w.upper() for w in ctx["work_types"]}
    wt_ids = {c["id"] for c in case_library.all_cases()
             if c.get("work_type", "").upper() in wanted}
    acct_fns = matcher._account_functions(set(ctx.get("functions", [])), match_notes)
    try:
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
                                    "domain": n.get("domain", ""), "use_case": n.get("use_case", ""),
                                    "sim": round(cand["cosine"], 2) if cand else 0.0})
            missing = missing[:10]
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
                missing = missing[:10]
    except Exception:
        priority_ids, missing, avoid, priority_reasons = [], [], [], {}

    # Industry-level gap (owner's spec, 2026-07-14): a salesperson-typed "Other"
    # industry (constants.all_industries()'s custom entries) with genuinely NO
    # related case study anywhere in the selected work types gets its own honest
    # flag, folded into the same "not in our library" card -- a capability gap
    # says "we don't have a slide for X"; this says "we have no proof at all for
    # this INDUSTRY". Built-in industries (BFSI, HEALTHCARE...) are never
    # flagged this way -- they always have at least some tagged coverage.
    acct_industry = ctx.get("industry", "")
    if acct_industry and acct_industry.upper() not in _BUILTIN_INDUSTRIES:
        try:
            # matcher-SHAPED rows (keywords middot-joined, primary_industry set) --
            # the same shape relevance.py's other industry logic already expects,
            # not the raw store records (whose "keywords" is a list, not a string).
            wt_rows = [row for rows in case_library.candidate_rows(ctx["work_types"]).values()
                      for row in rows]
            if not relevance.industry_has_coverage(acct_industry, wt_rows):
                missing.append({
                    "name": "%s case studies" % acct_industry,
                    "description": "We don't yet have a proven case study specifically "
                                   "for %s -- this deck leans on function/capability "
                                   "matches instead of an industry-specific proof point."
                                   % acct_industry,
                    "domain": acct_industry, "use_case": "", "sim": 0.0,
                })
        except Exception:
            pass                          # never let this flag block a build

    # --- strategic inference: capability bets that are NEVER stated outright, but
    #     that a sharp account researcher would infer by connecting the STAKEHOLDER's
    #     own career/role to the COMPANY's business (owner's spec, 2026-07-08 — the
    #     Pankaj Kumar Pant / Waaree Group case: nothing in his profile says "we want
    #     predictive maintenance," but 2 years running a solar ingot/wafer plant + a
    #     career built on defect-catching makes it a reasoned bet). Kept as a fully
    #     SEPARATE pass from the literal extraction above — a failure here can never
    #     touch a literal pick, and a literal pick is never overridden by a bet.
    #     ai_matcher.infer_strategic_fit() only returns a bet when it can argue BOTH
    #     a personal angle (why HE'D want it) and a company angle (why THEY'D fund
    #     it) — either missing and it's dropped before it reaches here.
    try:
        bets = ai_matcher.infer_strategic_fit(research_text, profile_text, ctx["transcript"], brief)
        if bets:
            recs = {r["id"]: r for r in case_library._load()}
            bet_shortlists = relevance.shortlist_cases(
                [b["name"] for b in bets], industry=ctx.get("industry", ""),
                functions=acct_fns, allowed_ids=wt_ids, top_n=8)
            for b, shortlist in zip(bets, bet_shortlists):
                cand = None
                for item in shortlist:
                    if item["id"] in priority_ids:
                        continue                        # already used for a literal need
                    if _is_avoided(recs.get(item["id"], {}), avoid):
                        continue                        # skip a mismatch-flagged case
                    cand = item
                    break
                if cand and (cand["cosine"] >= CAPABILITY_COVER or cand["title_hits"] >= 1):
                    priority_ids.append(cand["id"])
                    priority_reasons[cand["id"]] = {
                        "why": b["stakeholder_why"], "signal": "strategic bet",
                        "stakeholder_why": b["stakeholder_why"], "company_why": b["company_why"]}
    except Exception:
        pass

    # research + the profile's crisp focus areas LEAD the ranking; mail is secondary
    lead_research = (research_text + "\n"
                     + " ".join(f"{n['name']}. {n.get('description','')}" for n in profile_needs)).strip()
    result = matcher.plan({**ctx, "transcript": ctx["transcript"], "research": lead_research},
                          use_ai=True, priority_ids=priority_ids, avoid=avoid,
                          prefer_high_impact=bool(brief.get("prefer_high_impact")),
                          asked_flags={"asks_differentiation": bool(brief.get("asks_differentiation")),
                                      "asks_why_not_big_si": bool(brief.get("asks_why_not_big_si"))})
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
                 "company_footprint": "FOOT", "target_skill_profile": "TSP"}
        for c in sk:
            titles[c["id"]] = c["label"]
            reason = {
                "industry_strength": "auto-added — RFI: industry strength slide",
                "skill_deepdive":    "auto-added — RFI: skills deployed slide",
                "company_footprint": "auto-added — RFI: existing client relationship",
                "target_skill_profile": "auto-added — skills mentioned: target skill profile slide",
            }.get(c["kind"], "auto-added — RFI data slide")
            sk_picks.append({"slide_id": c["id"], "reason": reason, "skill": True,
                             "tag": _chip.get(c["kind"], "SKL"),
                             "label": c["label"]})
        picks[insert_at:insert_at] = sk_picks

    # --- Client Context + Tailored Approach (Workforce, First/Second only):
    #     a real AI extraction from the notes, fails closed (see client_context.py) ---
    cc = client_context.candidates(ctx)
    if cc:
        insert_at = next((i for i, p in enumerate(result["picks"]) if p["reason"].startswith("case")), None)
        if insert_at is None:
            insert_at = next((i for i, p in enumerate(result["picks"]) if p["slide_id"] in matcher.PIN_TO_END),
                             len(result["picks"]))
        cc_picks = []
        for c in cc:
            titles[c["id"]] = c["label"]
            cc_picks.append({"slide_id": c["id"], "reason": "auto-added — client context drawn from your notes",
                             "skill": True, "tag": "CC", "label": c["label"]})
        result["picks"][insert_at:insert_at] = cc_picks

    # --- pre-built content slides (F1 "Already have the content?", queued
    #     across possibly several pastes): merge them in by shape, same
    #     boundary the skills/context auto-adds above use -- a case-study-
    #     shaped one joins the case-study zone; anything else (four_box,
    #     roadmap_board, box_grid, ...) sits just above it, with the standard
    #     slides (owner's spec, 2026-07-09) ---
    content_ids = [x for x in request.form.get("content_slide_ids", "").split(",") if x]
    if content_ids:
        picks = result["picks"]
        insert_at = next((i for i, p in enumerate(picks) if p["reason"].startswith("case")), None)
        if insert_at is None:
            insert_at = next((i for i, p in enumerate(picks) if p["slide_id"] in matcher.PIN_TO_END),
                             len(picks))
        standard_extra, case_extra = [], []
        for cid in content_ids:
            # A REUSED library case (the Custom Slide Builder's duplicate check offers
            # "use this one") rides as its own store id, not a staging id -- it needs no
            # NEW: prefix and no rebuild; it just joins the case-study zone. Skipping it
            # here would silently drop a slide the salesperson deliberately queued.
            store_rec = case_library.record(cid.upper())
            if store_rec:
                if cid.upper() in {p["slide_id"] for p in picks}:
                    continue                     # the matcher already picked it
                titles[cid.upper()] = store_rec.get("title", "")
                case_extra.append({"slide_id": cid.upper(), "reason": "you reused this slide",
                                   "skill": True, "tag": "YOU",
                                   "label": store_rec.get("title", "")})
                continue
            rec = staging.get(cid)
            if not rec:
                continue
            sid = "NEW:" + cid
            titles[sid] = rec.get("title", "Untitled")
            entry = {"slide_id": sid, "reason": "your pasted content", "skill": True,
                     "tag": "YOU", "label": rec.get("title", "Untitled")}
            if rec.get("content_type", "case_study") == "case_study":
                case_extra.append(entry)
            else:
                standard_extra.append(entry)
        picks[insert_at:insert_at] = standard_extra + case_extra

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
            if pr.get("company_why"):     # a strategic bet: show BOTH angles, not just one
                r["stakeholder_why"] = pr.get("stakeholder_why")
                r["company_why"] = pr["company_why"]
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
                                  resume=False, build_id=build_id, reopen_seed=None)
    return shell(body, active="new", crumb="<b>New deck</b> / Suggested slides")


def _case_slide(sid, rec):
    """Editable slide data for a case study (content-store or AI-drafted)."""
    domain = rec.get("domain") or rec.get("industry") or ""
    function = rec.get("function") or ""
    sub = rec.get("subhead") or ""
    if (not domain or not function) and sub:          # AI drafts carry a 'subhead' string
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


def _build_review_slides(ids, client):
    """One editable 'slide' per id, in deck order: library slides (title/subtitle),
    case studies (full body editable), and data-driven skills slides (read-only)."""
    prs = Presentation(assembler.SOURCE)
    by_id = {read_id(s): s for s in prs.slides if read_id(s)}
    store = content_store()          # {id: record} for AIP/WFS/MSS store cases
    slides = []
    for sid in ids:
        if sid in (client_context.CS_CONTEXT, client_context.CS_APPROACH):
            # Client Context / Tailored Approach: filled from an AI extraction at
            # FINALIZE time (see deck_build.assemble), not editable here — showing
            # the raw [BRACKET] placeholders as "editable fields" would be confusing.
            slides.append({"kind": "skills", "id": sid})
        elif sid in by_id:                                       # master library slide (CSxx)
            # A predefined slide is shown as a read-only PREVIEW -- the real rendered slide,
            # exactly as it appears in the download. It is not edited on /review; it goes
            # into the deck as-is (owner's spec, 2026-07-13).
            slides.append({"kind": "library", "id": sid})
        elif sid in store:                                     # content-store case
            slides.append(_case_slide(sid, store[sid]))
        elif sid.startswith("NEW:"):                           # pasted or AI-created slide
            rec = staging.get(sid[4:])
            if not rec:
                continue
            kind = rec.get("content_type", "case_study")
            if kind == "case_study":
                slides.append(_case_slide(sid, rec))
            else:
                # every other shape is drawn by the SHARED inline editor macro
                # (templates/_slide_editor.html), the same one the Custom Slide Builder
                # uses. Before this, only four_box and roadmap_board rendered here, and
                # both were read-only; the other five fell through to the "no text to
                # edit" placeholder below.
                slides.append({"kind": kind, "id": sid,
                               "vm": slide_schema.view_model(rec),
                               "schema": slide_schema.fields_for(kind)})
        elif sid.startswith("SK:") or sid.startswith("FP:") or sid.startswith("TSP:"):   # data-driven skills slide
            slides.append({"kind": "skills", "id": sid})
    return slides


def _render_review(client, ids, *, industry="", transcript="", phase="", recipient="",
                   functions=None, work_types=None):
    slides = _build_review_slides(ids, client)
    body = render_template("review.html", client=client, order=",".join(ids), slides=slides,
                           industry=industry, transcript=transcript, phase=phase, recipient=recipient,
                           functions=functions or [], work_types=work_types or [],
                           shape_kinds=_SHAPE_KINDS)
    return shell(body, active="new", crumb="<b>New deck</b> / Review &amp; edit")


@bp.route("/review", methods=["POST"])
def review():
    client = request.form.get("client_name", "Client").strip()
    ids = [x for x in request.form.get("order", "").split(",") if x]
    return _render_review(client, ids,
                          industry=request.form.get("industry", ""),
                          transcript=request.form.get("transcript", ""),
                          phase=request.form.get("phase", ""),
                          recipient=request.form.get("recipient", ""),
                          functions=request.form.getlist("functions"),
                          work_types=request.form.getlist("work_types"))


@bp.route("/from_content", methods=["POST"])
def from_content():
    """F1: the user already has the content for ONE slide — paste it or upload a
    document. Which template it becomes is either auto-detected (does this read as
    a case study, or a four-way breakdown?) or set explicitly via the template_hint
    dropdown (owner's spec, 2026-07-08: pasted content isn't always a case study).
    Builds ONE branded slide and stages it, returning JSON (not a page) — the
    caller is new-form.js's queue, which lets the salesperson repeat this for
    several slides (e.g. handed one-by-one by the MS team) before either taking
    just those straight to review, or continuing into the full context form
    below so they get merged into the generated deck in the right spot (owner's
    spec, 2026-07-09: content slides should slot in among the standard slides or
    the case studies, wherever their own shape belongs, not bolt on separately)."""
    client = request.form.get("client_name", "Client").strip()
    industry = request.form.get("industry", "")
    work_type = request.form.get("work_type", "").strip().upper()
    content = request.form.get("content", "").strip()
    ftext = research.extract_text(request.files.get("content_file"))
    if ftext:
        content = (content + "\n\n" + ftext).strip()
    # BOTH content and work_type are required -- work_type used to be silently
    # optional here, which meant staging.add() got "" and case_library.
    # promote_ai_case() (called unconditionally at /finalize) would look up an
    # empty-string prefix, find nothing, and return None -- the slide built and
    # downloaded fine, but NEVER reached the content store, so it was invisible
    # in Slide Library / search / the Excel registry with no error anywhere
    # (owner-reported, 2026-07-08: a real case study went missing this way).
    if not content or work_type not in ("WORKFORCE", "AI_POD", "MS"):
        err = ("Paste the case-study content or attach a document first." if not content
              else "Choose a work type — it's needed to save this into the shared library.")
        return jsonify({"ok": False, "error": err})
    rec, tdef = slide_generator.build_content_slide(
        content, industry, request.form.get("template_hint", "auto").strip())
    staged = staging.add(rec, work_type, industry, client)
    return jsonify({"ok": True, "id": staged["id"], "title": staged.get("title", "Untitled"),
                    "content_type": staged.get("content_type", "case_study"),
                    "template_label": tdef["label"]})


@bp.route("/finalize", methods=["POST"])
def finalize():
    client = request.form.get("client_name", "Client").strip()
    phase = request.form.get("phase", "")
    ids = [x for x in request.form.get("order", "").split(",") if x]
    if not ids:
        abort(400)
    import os
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    # every version gets its own file, never overwritten (owner: keep every
    # version forever) -- e.g. "Joulestowatts_Acme Bank FV1.pptx", then FV2...
    # reserve_version() claims the number+filename ATOMICALLY, before the build
    # runs, so two near-simultaneous finalizes for the same client+phase can't
    # be handed the same version and silently overwrite each other's file.
    version, filename = meeting_log.reserve_version(client, phase)
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

    # Pasted non-case shapes (four-box, roadmap, box grid, deep-dive, scored list, stat
    # overview, data table) are edited through the shared inline-editor macro, which posts
    # ONE json field per slide. Apply them to the staged record straight away -- that
    # record is what deck_build reads, so the edit needs no further threading. Each goes
    # through its own normalizer (slide_schema.apply_edits), which also means a hand-edit
    # can't hand the renderer a shape it can't fill.
    for key, val in request.form.items():
        if not key.startswith("shape__NEW:"):
            continue
        stg_id = key[len("shape__NEW:"):]
        rec = staging.get(stg_id)
        if not rec:
            continue
        try:
            fields = slide_schema.apply_edits(rec, json.loads(val),
                                              request.form.get("industry", ""))
        except (ValueError, TypeError):
            continue                # a malformed payload leaves the slide as it was built
        staging.update_fields(stg_id, fields)

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
                            edits=edits, case_edits=case_edits,
                            phase=request.form.get("phase", ""))
    except (PermissionError, BadZipFile) as e:
        # the build never produced a real deck -- free the claimed version
        # number/filename so the next attempt doesn't skip one for nothing.
        meeting_log.cancel_reservation(client, phase, version)
        return file_busy_page(e)
    meeting_log.finalize_version(               # auto-log this version (no extra step)
        client=client, phase=phase, version=version,
        industry=request.form.get("industry", ""),
        functions=request.form.getlist("functions"),
        work_types=request.form.getlist("work_types"),
        recipient=request.form.get("recipient", ""),
        salesperson=current_salesperson(),
        slide_ids=final_ids,
        edits=edits,
        case_edits=case_edits,
    )
    count = len(list(Presentation(path).slides))
    body = render_template("preview.html", client=client, filename=filename, count=count)
    return shell(body, active="new", crumb="<b>New deck</b> / Done")
