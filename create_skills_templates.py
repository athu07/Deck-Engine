# -*- coding: utf-8 -*-
"""
create_skills_templates.py
Builds skills_templates.pptx from scratch — no source PPT needed.
All 3 skills slides are drawn programmatically with the correct
{{PLACEHOLDER}} names the engine expects.

Run once (safe to re-run — overwrites skills_templates.pptx):
    py create_skills_templates.py
"""

from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

OUT = "skills_templates.pptx"

# ── Brand colours ─────────────────────────────────────────────────────────────
TEAL   = RGBColor(0x2C, 0x6E, 0x66)   # J2W primary dark teal
TEAL_M = RGBColor(0x7F, 0xB2, 0xA9)   # mid teal (dividers)
TEAL_L = RGBColor(0xCF, 0xE7, 0xE2)   # very light teal (cards, subtitle)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
DARK   = RGBColor(0x11, 0x11, 0x10)
GREY   = RGBColor(0x55, 0x55, 0x55)   # secondary text
BG     = RGBColor(0xF0, 0xF2, 0xF1)   # slide background

# ── Slide dimensions (13.33" x 7.50" widescreen) ─────────────────────────────
SW, SH = 13.33, 7.50
EMU    = 914400
I      = lambda v: int(v * EMU)   # inches to EMU

# ── Shared layout constants ───────────────────────────────────────────────────
HEADER_H = 1.05    # teal header bar height
MARGIN   = 0.20    # outer margin
PAD      = 0.15    # inner padding inside panels


# ── Low-level helpers ─────────────────────────────────────────────────────────
def _blank_layout(prs):
    for layout in prs.slide_layouts:
        if (layout.name or "").lower().strip() == "blank":
            return layout
    return prs.slide_layouts[-1]


def _no_border(shape):
    spPr = shape._element.spPr
    ln = spPr.find(qn('a:ln'))
    if ln is None:
        ln = etree.SubElement(spPr, qn('a:ln'))
    for c in list(ln):
        ln.remove(c)
    etree.SubElement(ln, qn('a:noFill'))


def rect(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(1, I(l), I(t), I(w), I(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    _no_border(s)
    return s


def txt(slide, l, t, w, h, text, size, color,
        bold=False, align=PP_ALIGN.LEFT, italic=False,
        anchor='t', wrap=True):
    tb = slide.shapes.add_textbox(I(l), I(t), I(w), I(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf._txBody.bodyPr.set('anchor', anchor)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.bold = bold
    r.font.italic = italic
    return tb


def divider_line(slide, l, t, w):
    rect(slide, l, t, w, 0.022, TEAL_M)


def header_bar(slide, title, subtitle):
    rect(slide, 0, 0, SW, HEADER_H, TEAL)
    txt(slide, 0.30, 0.09, SW-0.40, 0.56, title, 17, WHITE, bold=True)
    txt(slide, 0.30, 0.67, SW-0.40, 0.32, subtitle, 9, TEAL_L, italic=True)


def panel_card(slide, l, t, w, h, accent="left"):
    rect(slide, l, t, w, h, WHITE)
    if accent == "left":
        rect(slide, l, t, 0.05, h, TEAL)
    elif accent == "top":
        rect(slide, l, t, w, 0.05, TEAL)


def panel_section_label(slide, l, t, w, label):
    txt(slide, l+PAD, t+PAD, w-PAD*2, 0.28, label, 11, TEAL, bold=True)
    divider_line(slide, l+PAD, t+PAD+0.32, w-PAD*2)


def set_notes(slide, tag):
    slide.notes_slide.notes_text_frame.text = tag


# ── SLIDE 1 — skill_deepdive ──────────────────────────────────────────────────
#
#  Placeholders  (filled at runtime by skills.py > _mapping_skills):
#    {{SKILLS_HEADER}}   joined skill names shown in the header
#    {{SKILL_SUMMARY}}   multi-line: per-skill count + top companies
#    {{CHART}}           horizontal bar chart (skill vs consultant count)
#
def build_slide1(prs):
    slide = prs.slides.add_slide(_blank_layout(prs))
    rect(slide, 0, 0, SW, SH, BG)

    header_bar(slide,
               "WHERE J2W HAS DELIVERED  —  {{SKILLS_HEADER}}",
               "Skills deployed across J2W client engagements")

    PANEL_T = HEADER_H + 0.10
    PANEL_H = SH - PANEL_T - 0.10    # 6.25"
    LEFT_W  = 5.55
    GAP     = 0.23
    RIGHT_L = MARGIN + LEFT_W + GAP  # 6.18"
    RIGHT_W = SW - RIGHT_L - MARGIN  # 6.95"

    # Left panel: deployment summary text
    LL, LT, LW, LH = MARGIN, PANEL_T, LEFT_W, PANEL_H
    panel_card(slide, LL, LT, LW, LH, accent="left")
    panel_section_label(slide, LL, LT, LW, "DEPLOYMENT SUMMARY")
    txt(slide,
        LL+PAD, LT+PAD+0.38,
        LW-PAD-0.10, LH-PAD-0.48,
        "{{SKILL_SUMMARY}}", 10, DARK, wrap=True, anchor='t')

    # Right panel: skill distribution chart
    RL, RT, RW, RH = RIGHT_L, PANEL_T, RIGHT_W, PANEL_H
    panel_card(slide, RL, RT, RW, RH, accent="top")
    panel_section_label(slide, RL, RT, RW, "SKILL DISTRIBUTION")
    txt(slide,
        RL+PAD, RT+PAD+0.38,
        RW-PAD*2, RH-PAD-0.48,
        "{{CHART}}", 10, GREY, anchor='t')

    set_notes(slide, "J2W_TEMPLATE: skill_deepdive")
    print("  Slide 1 (skill_deepdive): done")


# ── SLIDE 2 — industry_strength ───────────────────────────────────────────────
#
#  Placeholders  (filled at runtime by skills.py > _mapping_industry):
#    {{INDUSTRY_NAME}}      industry label, e.g. "Banking & Financial Services"
#    {{TOTAL_CONSULTANTS}}  headcount
#    {{NUM_COMPANIES}}      distinct companies
#    {{NUM_FUNCTIONS}}      distinct functions
#    {{NUM_SKILLS}}         distinct skills
#    {{TOP_SKILL_1/2/3}}    top 3 skills by company coverage
#    {{CHART}}              pie chart (function breakdown %)
#
def build_slide2(prs):
    slide = prs.slides.add_slide(_blank_layout(prs))
    rect(slide, 0, 0, SW, SH, BG)

    header_bar(slide,
               "OUR DELIVERY FOOTPRINT IN  {{INDUSTRY_NAME}}",
               "J2W track record and capability across this industry")

    # 4 metric tiles
    TILE_H   = 1.40
    TILE_GAP = 0.14
    TILE_W   = (SW - 2*MARGIN - 3*TILE_GAP) / 4   # 3.12"
    TILE_T   = HEADER_H + 0.10

    TILES = [
        ("{{TOTAL_CONSULTANTS}}", "Consultants deployed"),
        ("{{NUM_COMPANIES}}",     "Companies served"),
        ("{{NUM_FUNCTIONS}}",     "Functions delivered"),
        ("{{NUM_SKILLS}}",        "Distinct skills deployed"),
    ]
    for i, (marker, label) in enumerate(TILES):
        tl = MARGIN + i * (TILE_W + TILE_GAP)
        panel_card(slide, tl, TILE_T, TILE_W, TILE_H, accent="top")
        # Metric value (large, centred)
        txt(slide, tl+PAD, TILE_T+0.10, TILE_W-PAD*2, 0.68,
            marker, 26, TEAL, bold=True, align=PP_ALIGN.CENTER)
        # Metric label
        txt(slide, tl+PAD, TILE_T+0.80, TILE_W-PAD*2, 0.52,
            label, 9, GREY, align=PP_ALIGN.CENTER, anchor='t')

    # Divider
    DIV_T = TILE_T + TILE_H + 0.08
    divider_line(slide, MARGIN, DIV_T, SW-2*MARGIN)

    # Lower section
    LOW_T = DIV_T + 0.10
    LOW_H = SH - LOW_T - 0.10   # ~4.52"
    LEFT_W  = 5.55
    GAP     = 0.23
    RIGHT_L = MARGIN + LEFT_W + GAP
    RIGHT_W = SW - RIGHT_L - MARGIN

    # Left lower panel: top 3 skill cards
    LL, LT, LW, LH = MARGIN, LOW_T, LEFT_W, LOW_H
    panel_card(slide, LL, LT, LW, LH, accent="left")
    panel_section_label(slide, LL, LT, LW, "TOP HIRED SKILLS  —  BY COMPANIES SERVED")

    CARD_T0  = LT + PAD + 0.40
    AVAIL    = LH - PAD*2 - 0.40
    CARD_GAP = 0.09
    CARD_H   = (AVAIL - 2*CARD_GAP) / 3   # ~1.37"
    BADGE_W  = 0.44

    for j, marker in enumerate(["{{TOP_SKILL_1}}", "{{TOP_SKILL_2}}", "{{TOP_SKILL_3}}"]):
        ct = CARD_T0 + j * (CARD_H + CARD_GAP)
        rect(slide, LL+PAD, ct, LW-PAD*2, CARD_H, TEAL_L)   # card bg
        rect(slide, LL+PAD, ct, BADGE_W, CARD_H, TEAL)        # rank badge
        txt(slide, LL+PAD, ct, BADGE_W, CARD_H,
            f"#{j+1}", 13, WHITE, bold=True,
            align=PP_ALIGN.CENTER, anchor='ctr')
        txt(slide, LL+PAD+BADGE_W+0.10, ct+0.08,
            LW-PAD*2-BADGE_W-0.14, CARD_H-0.16,
            marker, 11, DARK, wrap=True, anchor='t')

    # Right lower panel: pie chart
    RL, RT, RW, RH = RIGHT_L, LOW_T, RIGHT_W, LOW_H
    panel_card(slide, RL, RT, RW, RH, accent="top")
    panel_section_label(slide, RL, RT, RW, "FUNCTION DISTRIBUTION")
    txt(slide,
        RL+PAD, RT+PAD+0.38,
        RW-PAD*2, RH-PAD-0.48,
        "{{CHART}}", 10, GREY, anchor='t')

    set_notes(slide, "J2W_TEMPLATE: industry_strength")
    print("  Slide 2 (industry_strength): done")


# ── SLIDE 3 — company_footprint ───────────────────────────────────────────────
#
#  Placeholders  (filled at runtime by skills.py > _mapping_company):
#    {{COMPANY_NAME}}         client company display name
#    {{TOTAL_DEPLOYED}}       total consultants at this company
#    {{NUM_FUNCTIONS_CO}}     distinct functions at this company
#    {{NUM_SKILLS_CO}}        distinct skills at this company
#    {{ENGAGEMENT_TYPE}}      always "Existing client"
#    {{FUNCTION_BREAKDOWN}}   multi-line bullet list (function: count)
#    {{CHART}}                horizontal bar chart (function vs count)
#
def build_slide3(prs):
    slide = prs.slides.add_slide(_blank_layout(prs))
    rect(slide, 0, 0, SW, SH, BG)

    header_bar(slide,
               "OUR PARTNERSHIP WITH  {{COMPANY_NAME}}  —  DELIVERY OVERVIEW",
               "Our existing J2W relationship — people, skills & engagement")

    # 4 metric tiles
    TILE_H   = 1.40
    TILE_GAP = 0.14
    TILE_W   = (SW - 2*MARGIN - 3*TILE_GAP) / 4
    TILE_T   = HEADER_H + 0.10

    TILES = [
        ("{{TOTAL_DEPLOYED}}",    "Consultants deployed"),
        ("{{NUM_FUNCTIONS_CO}}", "Functions delivered"),
        ("{{NUM_SKILLS_CO}}",    "Distinct skills"),
        ("{{ENGAGEMENT_TYPE}}",  "Engagement type"),
    ]
    for i, (marker, label) in enumerate(TILES):
        tl = MARGIN + i * (TILE_W + TILE_GAP)
        panel_card(slide, tl, TILE_T, TILE_W, TILE_H, accent="top")
        txt(slide, tl+PAD, TILE_T+0.10, TILE_W-PAD*2, 0.68,
            marker, 26, TEAL, bold=True, align=PP_ALIGN.CENTER)
        txt(slide, tl+PAD, TILE_T+0.80, TILE_W-PAD*2, 0.52,
            label, 9, GREY, align=PP_ALIGN.CENTER, anchor='t')

    # Divider
    DIV_T = TILE_T + TILE_H + 0.08
    divider_line(slide, MARGIN, DIV_T, SW-2*MARGIN)

    # Lower section
    LOW_T = DIV_T + 0.10
    LOW_H = SH - LOW_T - 0.10
    LEFT_W  = 5.55
    GAP     = 0.23
    RIGHT_L = MARGIN + LEFT_W + GAP
    RIGHT_W = SW - RIGHT_L - MARGIN

    # Left lower panel: function breakdown text
    LL, LT, LW, LH = MARGIN, LOW_T, LEFT_W, LOW_H
    panel_card(slide, LL, LT, LW, LH, accent="left")
    panel_section_label(slide, LL, LT, LW, "FUNCTION BREAKDOWN")
    txt(slide,
        LL+PAD, LT+PAD+0.38,
        LW-PAD-0.10, LH-PAD-0.48,
        "{{FUNCTION_BREAKDOWN}}", 10, DARK, wrap=True, anchor='t')

    # Right lower panel: bar chart
    RL, RT, RW, RH = RIGHT_L, LOW_T, RIGHT_W, LOW_H
    panel_card(slide, RL, RT, RW, RH, accent="top")
    panel_section_label(slide, RL, RT, RW, "FUNCTION DISTRIBUTION")
    txt(slide,
        RL+PAD, RT+PAD+0.38,
        RW-PAD*2, RH-PAD-0.48,
        "{{CHART}}", 10, GREY, anchor='t')

    set_notes(slide, "J2W_TEMPLATE: company_footprint")
    print("  Slide 3 (company_footprint): done")


# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    prs = Presentation()
    prs.slide_width  = I(SW)
    prs.slide_height = I(SH)

    build_slide1(prs)
    build_slide2(prs)
    build_slide3(prs)

    prs.save(OUT)
    print(f"\nSaved: {OUT}  ({len(prs.slides)} slides)")
    for i, s in enumerate(prs.slides, 1):
        note = s.notes_slide.notes_text_frame.text.strip()
        print(f"  Slide {i} notes: {repr(note)}")
