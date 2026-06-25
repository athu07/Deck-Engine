# -*- coding: utf-8 -*-
"""
skills.py  --  the two data-driven "skills" slides (Workforce only).

Source data = J2W_Skills_Inventory.xlsx (aggregated sheets ONLY; the
"Consultant Detail" sheet is NEVER read into a deck).

Two master-deck template slides carry the markers:
  J2W_TEMPLATE: skills            -> capability slide (Skills Master sheet)
  J2W_TEMPLATE: company_footprint -> footprint slide (Client Footprint sheet)

Rules (see candidates()):
  - GATE: only a PURE Workforce deck (work types == {WORKFORCE}) gets these.
  - Capability: one slide per skill_area that matches the notes (keyword) OR the
    deck's industry (industries_served).
  - Footprint: one slide if the form's client matches a Client Footprint row.
  - Staleness: a row whose last_verified is > 90 days old is flagged.
"""
import re
from datetime import date, datetime

import openpyxl
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.dml.color import RGBColor

# J2W brand palette for chart slices (teal shades + ink)
BRAND = [RGBColor(0x2C, 0x6E, 0x66), RGBColor(0x7F, 0xB2, 0xA9),
         RGBColor(0xCF, 0xE7, 0xE2), RGBColor(0x11, 0x11, 0x10)]

EXCEL = "J2W_Skills_Inventory.xlsx"
SHEET_SKILLS = "Skills Master (Aggregated)"
SHEET_FOOT = "Client Footprint (Aggregated)"
STALE_DAYS = 90

# marker -> column
SK_MAP = {"SKILL_AREA": "skill_area", "TOTAL_CONSULTANTS": "total_consultants",
          "EXPERT_COUNT": "expert", "AVG_RAMP_UP": "avg_ramp_up_weeks",
          "AVAILABLE_NOW": "available_now", "INDUSTRIES_SERVED": "industries_served",
          "EXAMPLE_CLIENTS": "example_clients"}
FP_MAP = {"CLIENT_NAME": "client_name", "TOTAL_DEPLOYED": "total_deployed",
          "ACTIVE_ENGAGEMENTS": "active_engagements", "SKILL_AREAS_COUNT": "skill_areas_count",
          "RELATIONSHIP_DURATION": "relationship_duration", "SKILLS_DEPLOYED": "skills_deployed",
          "DIVISIONS_SERVED": "divisions_served", "EXPANSION_AREAS": "expansion_areas",
          "SNAPSHOT_DATE": "snapshot_date"}


def _rows(sheet):
    wb = openpyxl.load_workbook(EXCEL, data_only=True)
    ws = wb[sheet]
    hdr = [c.value for c in ws[1]]
    out = []
    for r in range(2, ws.max_row + 1):
        rec = {hdr[i]: ws.cell(r, i + 1).value for i in range(len(hdr))}
        out.append(rec)
    return out


def load_skills():
    # skip the TOTAL / BENCH aggregate row (no industries_served)
    return [r for r in _rows(SHEET_SKILLS) if (r.get("industries_served") or "").strip()]


def load_footprint():
    return [r for r in _rows(SHEET_FOOT) if (r.get("client_name") or "").strip()]


def _stale(val, days=STALE_DAYS):
    if not val:
        return False
    if isinstance(val, datetime):
        d = val.date()
    elif isinstance(val, date):
        d = val
    else:
        try:
            d = datetime.strptime(str(val)[:10], "%Y-%m-%d").date()
        except Exception:
            return False
    return (date.today() - d).days > days


def _industry_token(industry):
    """Deck industry -> a token to look for in industries_served. Handles e.g.
    TECH_IT -> 'tech' (matches 'Technology')."""
    return (industry or "").split("_")[0].strip().lower()


CAP_INDUSTRY = 3   # max industry-only capability slides (keyword matches are uncapped)


# Words too generic to be a reliable keyword hit on their own (they appear in many
# skill-area names — e.g. "engineering" is in 5 of them).
STOPWORDS = {"engineering", "solutions", "platforms", "services", "management",
             "modernization", "legacy"}


def _capability_reason(row, industry, transcript):
    """'keyword' if the skill area is mentioned in the notes (strong signal),
    'industry' if only the deck's industry matches (broad signal), else None."""
    area = (row.get("skill_area") or "").lower()
    notes = (transcript or "").lower()
    if area and area in notes:                        # whole skill-area phrase
        return "keyword"
    for w in re.findall(r"[a-z]{4,}", area):          # distinctive words only
        if w in STOPWORDS:
            continue
        if re.search(r"\b" + re.escape(w) + r"\b", notes):
            return "keyword"
    tok = _industry_token(industry)
    if tok and tok in (row.get("industries_served") or "").lower():
        return "industry"
    return None


def _mapping(row, marker_map):
    return {m: ("" if row.get(col) is None else str(row.get(col))) for m, col in marker_map.items()}


def candidates(context):
    """Return the skills slides to auto-add for this deck context (already
    work-type-gated). Each: {id, kind, label, mapping, stale, last_verified}."""
    wts = {str(w).upper() for w in (context.get("work_types") or [])}
    if wts != {"WORKFORCE"}:                          # GATE: pure Workforce only
        return []
    industry = context.get("industry", "")
    transcript = context.get("transcript", "")
    client = (context.get("client_name", "") or "").strip()
    out = []

    # keyword matches are uncapped; industry-only matches are capped (keyword-first)
    kw_rows, ind_rows = [], []
    for r in load_skills():
        reason = _capability_reason(r, industry, transcript)
        if reason == "keyword":
            kw_rows.append(r)
        elif reason == "industry":
            ind_rows.append(r)
    def _n(v):
        try:
            return int(v)
        except (TypeError, ValueError):
            return 0

    for r in kw_rows + ind_rows[:CAP_INDUSTRY]:
        area = (r.get("skill_area") or "").replace(",", " ")
        out.append({"id": "SK:" + area, "kind": "capability",
                    "label": "Capability — " + area, "template": "skills",
                    "mapping": _mapping(r, SK_MAP),
                    "chart": {"Expert": _n(r.get("expert")),
                              "Intermediate": _n(r.get("intermediate")),
                              "Junior": _n(r.get("junior"))},
                    "stale": _stale(r.get("last_verified")),
                    "last_verified": str(r.get("last_verified") or "")})

    for r in load_footprint():
        if client and (r.get("client_name") or "").strip().lower() == client.lower():
            name = (r.get("client_name") or "").replace(",", " ")
            out.append({"id": "FP:" + name, "kind": "footprint",
                        "label": "Footprint — " + name, "template": "company_footprint",
                        "mapping": _mapping(r, FP_MAP),
                        "stale": _stale(r.get("last_verified")),
                        "last_verified": str(r.get("last_verified") or "")})
            break
    return out


def by_id(context, sid):
    """Re-derive a single candidate by its synthetic id (used at assembly time)."""
    return next((c for c in candidates(context) if c["id"] == sid), None)


def fill_markers(slide, mapping):
    """Replace {{MARKER}} tokens in a slide's text, keeping each paragraph's first
    run formatting."""
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            full = "".join(run.text for run in p.runs)
            if "{{" not in full:
                continue
            new = full
            for m, v in mapping.items():
                new = new.replace("{{" + m + "}}", v)
            if new != full and p.runs:
                p.runs[0].text = new
                for run in p.runs[1:]:
                    run.text = ""


def add_proficiency_chart(slide, data):
    """If the slide has a {{CHART}} placeholder shape, replace it with a native
    doughnut of the proficiency mix (brand colours). No placeholder -> no-op (so it
    never overlaps a template that hasn't made room yet)."""
    holder = None
    for sh in slide.shapes:
        if sh.has_text_frame and "{{CHART}}" in sh.text_frame.text:
            holder = sh
            break
    if holder is None:
        return False
    left, top, width, height = holder.left, holder.top, holder.width, holder.height
    holder._element.getparent().remove(holder._element)        # drop the placeholder

    cats = [k for k, v in data.items() if v]
    vals = [int(v) for k, v in data.items() if v]
    if not vals:
        return False
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("Proficiency", vals)
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, left, top, width, height, cd)
    chart = gframe.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    for i, pt in enumerate(chart.plots[0].series[0].points):
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = BRAND[i % len(BRAND)]
    return True


def find_template(prs, name):
    """The master-deck slide tagged J2W_TEMPLATE: <name>."""
    tag = "J2W_TEMPLATE: " + name
    for s in prs.slides:
        if s.has_notes_slide:
            for line in (s.notes_slide.notes_text_frame.text or "").splitlines():
                if line.strip() == tag:
                    return s
    return None


def build_into(deck_path, order, cands, master_path="WORKING_COPY_Master_Deck.pptx"):
    """After the CS deck is assembled, copy in the chosen skills slides (filled from
    their Excel row) and put the whole deck into `order`. `order` = the full id list
    (CSxx + SK:/FP:). `cands` = the candidates() result (id -> mapping/template)."""
    from pptx import Presentation
    import slide_generator
    import assembler
    from build_library import read_id

    cand_by_id = {c["id"]: c for c in cands if c["id"] in order}
    if not cand_by_id:
        return 0
    prs = Presentation(deck_path)
    sld_id_lst = prs.slides._sldIdLst
    master = Presentation(master_path)
    tfile = None                                       # templates.pptx, loaded lazily

    skill_elem = {}
    for sid, c in cand_by_id.items():
        name = c["template"]
        if name in ("skills", "company_footprint"):    # templates live IN the master
            t = find_template(master, name)
        else:                                          # e.g. case_study_full -> templates.pptx
            if tfile is None:
                tfile = Presentation(slide_generator.TEMPLATES_FILE)
            t = find_template(tfile, name)
        if t is None:
            continue
        new = slide_generator._copy_slide(prs, t)     # text-only template -> copies cleanly
        if name == "case_study_full":
            slide_generator.fill_case_study(new, c["content"])
        else:
            fill_markers(new, c["mapping"])
            if c.get("kind") == "capability" and c.get("chart"):
                add_proficiency_chart(new, c["chart"])  # doughnut into the {{CHART}} box, if present
        skill_elem[sid] = list(sld_id_lst)[-1]         # the sldId just appended

    # reorder the whole deck to match `order` (append moves the node).
    # NB: use explicit None checks — lxml elements are "falsy" when childless, so
    # `cs_elem.get(sid) or skill_elem.get(sid)` would wrongly skip slides.
    cs_elem = {read_id(s): e for s, e in zip(prs.slides, list(sld_id_lst)) if read_id(s)}
    for sid in order:
        e = cs_elem.get(sid)
        if e is None:
            e = skill_elem.get(sid)
        if e is not None:
            sld_id_lst.append(e)

    assembler._atomic_save(prs, deck_path)
    return len(skill_elem)
