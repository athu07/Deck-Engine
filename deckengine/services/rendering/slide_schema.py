# -*- coding: utf-8 -*-
"""
slide_schema.py  --  what is editable on each of the eight slide shapes.

The Custom Slide Builder shows a real rendered PNG of every slide it builds. A picture
you cannot correct is only half a tool: the salesperson spots a wrong number or a clumsy
line and has nowhere to fix it. This module says, per `content_type`, which fields the
slide actually has -- so ONE generic editor in the browser can edit all eight shapes
without a hand-written form per shape.

The field keys are exactly the keys `skills._mapping_*` / `_draw_*` and
`fill_case_study.build_mapping` read when rendering. If you add a shape to
`slide_generator.CONTENT_TEMPLATES`, add its fields here too, or it will build and
render but not be editable.

Every save is re-run through the shape's own `_normalize_*` function, so a hand-edit
cannot break an invariant the renderer depends on (a case study has exactly 6
capabilities and 3 results; a box grid has at least 2 boxes; and so on). Trusting the
browser to preserve those would be a slow-burning bug.

Field types:
    text      one line
    textarea  several lines
    strings   a list of one-line strings           (e.g. results, bullet items)
    objects   a list of records, each with `fields` (which may themselves be `strings`)
"""

from deckengine.services.rendering import slide_generator


def _t(key, label, type_="text"):
    return {"key": key, "label": label, "type": type_}


def _objects(key, label, item_label, fields):
    return {"key": key, "label": label, "type": "objects",
            "item_label": item_label, "fields": fields}


# content_type -> the editable shape of that slide
SCHEMA = {
    "case_study": [
        _t("title", "Title"),
        _t("subhead", "Client · Domain · Function"),
        _t("challenge", "The challenge", "textarea"),
        _t("solution", "The solution", "textarea"),
        _t("capabilities", "Key capabilities (Name: what it delivers)", "strings"),
        _t("results", "Results", "strings"),
    ],
    "four_box": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _objects("boxes", "The four boxes", "Box",
                 [_t("heading", "Heading"), _t("body", "Body", "textarea")]),
    ],
    "box_grid": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _objects("boxes", "Sections", "Section",
                 [_t("heading", "Heading"), _t("body", "Body", "textarea")]),
    ],
    "roadmap_board": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _t("intro", "Intro", "textarea"),
        _objects("columns", "Columns", "Column",
                 [_t("name", "Name"), _t("tag", "Phase / tag"),
                  _t("items", "Items", "strings")]),
        _objects("legend", "Legend", "Entry",
                 [_t("tag", "Tag"), _t("note", "Note")]),
        _t("footer_title", "Footer title"),
        _t("footer_body", "Footer body", "textarea"),
    ],
    "pillar_deepdive": [
        _t("eyebrow", "Eyebrow"),
        _t("title", "Title"),
        _objects("blocks", "Blocks", "Block",
                 [_t("heading", "Heading"), _t("body", "Body", "textarea"),
                  _t("subpoints", "Sub-points", "strings")]),
    ],
    "scored_list": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _objects("rows", "Rows", "Row",
                 [_t("name", "Name"), _t("description", "Description", "textarea"),
                  _t("stat", "Figure")]),
    ],
    "stat_overview": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _t("intro", "Intro", "textarea"),
        _objects("stats", "Headline stats", "Stat",
                 [_t("value", "Value"), _t("label", "Label")]),
        _t("items", "Named components", "strings"),
    ],
    "data_table": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _t("intro", "Intro", "textarea"),
        _t("col_labels", "Column headings", "strings"),
        _objects("rows", "Rows", "Row",
                 [_t("label", "Item"), _t("tag", "Category"), _t("value", "Value")]),
    ],
    # ── the ten style-guide shapes (owner's designs, 2026-07-10) ─────────────
    "pain_point_list": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _objects("rows", "The problems", "Problem",
                 [_t("label", "Label"), _t("body", "Consequence", "textarea")]),
    ],
    "platform_overview": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _objects("stats", "Headline stats", "Stat",
                 [_t("value", "Value"), _t("label", "Label")]),
        _t("capabilities", "Core capabilities", "strings"),
        _t("footer_title", "Footer band title"),
        _t("footer_items", "Footer band items", "strings"),
    ],
    "before_after_split": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _t("intro", "Intro line", "textarea"),
        _t("before_title", "Before: lane title"),
        _objects("before_stages", "Before: stages", "Stage",
                 [_t("tag", "Who does it"), _t("label", "Stage")]),
        _t("after_title", "After: lane title"),
        _objects("after_stages", "After: stages", "Stage",
                 [_t("tag", "Who does it"), _t("label", "Stage")]),
        _objects("questions", "The questions this raises", "Question",
                 [_t("title", "Question"), _t("body", "Detail", "textarea")]),
    ],
    "comparison_split": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _t("panel_title", "Left panel title"),
        _t("panel_intro", "Left panel intro", "textarea"),
        _objects("features", "Capability cards", "Capability",
                 [_t("heading", "Heading"), _t("body", "Body", "textarea")]),
        _t("table_title", "Table title"),
        _t("col_a", "Column A"),
        _t("col_b", "Column B"),
        _objects("rows", "Table rows", "Row",
                 [_t("metric", "Metric"), _t("a", "Value A"), _t("b", "Value B")]),
        _t("takeaway", "Closing line", "textarea"),
    ],
    "pillar_grid": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _objects("pillars", "Pillars", "Pillar",
                 [_t("heading", "Heading"), _t("body", "Body", "textarea"),
                  _t("points", "Supporting points", "strings")]),
    ],
    "option_columns": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _objects("options", "Options", "Option",
                 [_t("name", "Name"), _t("tag", "Characterisation"),
                  _objects("rows", "Dimensions", "Dimension",
                           [_t("label", "Label"), _t("value", "Value", "textarea")])]),
        _t("recommendation", "Recommendation", "textarea"),
    ],
    "agent_architecture": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _objects("agents", "Components", "Component",
                 [_t("name", "Name"), _t("body", "What it does", "textarea"),
                  _t("badge", "Metric badge")]),
        _t("footer", "Footer bar"),
    ],
    "governance_list": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _objects("items", "Layers", "Layer",
                 [_t("heading", "Heading"), _t("body", "Body", "textarea")]),
    ],
    "guardrail_columns": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _objects("columns", "Themes", "Theme",
                 [_t("heading", "Heading"),
                  _objects("points", "Points", "Point",
                           [_t("lead", "Bold lead-in"), _t("body", "Rest of the point", "textarea")])]),
        _t("callout_label", "Callout label"),
        _t("callout_body", "Callout text", "textarea"),
    ],
    "opportunity_cards": [
        _t("title", "Title"),
        _t("subhead", "Subhead"),
        _objects("cards", "Opportunities", "Opportunity",
                 [_t("heading", "Heading"), _t("opportunity", "The opportunity", "textarea"),
                  _t("outcome", "The outcome", "textarea")]),
    ],
}



def fields_for(content_type):
    """The editable fields of this shape, or [] if we don't know the shape."""
    return SCHEMA.get(content_type or "case_study", [])


def _field_keys(content_type):
    return {f["key"] for f in fields_for(content_type)}


def normalize(content_type, data, industry=""):
    """Re-run edited data through the shape's own normalizer, so a hand-edit can't
    break an invariant the renderer relies on. Returns the normalized record."""
    if content_type == "case_study":
        return slide_generator._normalize_case_study(data, industry)
    fn = {
        "four_box": slide_generator._normalize_four_box,
        "box_grid": slide_generator._normalize_box_grid,
        "roadmap_board": slide_generator._normalize_roadmap,
        "pillar_deepdive": slide_generator._normalize_pillar,
        "scored_list": slide_generator._normalize_scored_list,
        "stat_overview": slide_generator._normalize_stat_overview,
        "data_table": slide_generator._normalize_data_table,
        # the ten style-guide shapes
        "pain_point_list": slide_generator._normalize_pain_point_list,
        "platform_overview": slide_generator._normalize_platform_overview,
        "before_after_split": slide_generator._normalize_before_after_split,
        "comparison_split": slide_generator._normalize_comparison_split,
        "pillar_grid": slide_generator._normalize_pillar_grid,
        "option_columns": slide_generator._normalize_option_columns,
        "agent_architecture": slide_generator._normalize_agent_architecture,
        "governance_list": slide_generator._normalize_governance_list,
        "guardrail_columns": slide_generator._normalize_guardrail_columns,
        "opportunity_cards": slide_generator._normalize_opportunity_cards,
    }.get(content_type)
    return fn(data) if fn else dict(data)


def extract(content_type, record):
    """Just the editable fields of `record`, in schema order -- what the editor shows."""
    return {f["key"]: record.get(f["key"], [] if f["type"] in ("strings", "objects") else "")
            for f in fields_for(content_type)}


def view_model(record):
    """The record as the inline editor draws it (templates/_slide_editor.html).

    Normalized first, so the slots on screen are exactly the slots the template can
    render -- a case study always shows 6 capabilities and 3 results, a four-box always
    shows 4 boxes -- rather than however many the AI happened to produce.

    Capabilities are stored as "Name: what it delivers" strings; the editor gives them
    two fields, as the review page always has, so they're split here and re-joined by
    the browser on save.
    """
    content_type = record.get("content_type", "case_study")
    clean = normalize(content_type, dict(record), record.get("industry", ""))
    vm = extract(content_type, clean)
    if content_type == "case_study":
        from deckengine.services.rendering.fill_case_study import split_capability
        vm["capabilities"] = [split_capability(c) for c in vm.get("capabilities", [])]
    return vm


def apply_edits(record, edits, industry=""):
    """Merge the editor's values into a staged record and re-normalize.

    Only keys the schema declares are taken from `edits` -- a browser cannot introduce a
    field, change the shape, or overwrite the record's bookkeeping (id, work_type,
    promoted_id...). Returns the fields to persist.
    """
    content_type = record.get("content_type", "case_study")
    allowed = _field_keys(content_type)
    merged = dict(record)
    merged.update({k: v for k, v in (edits or {}).items() if k in allowed})
    clean = normalize(content_type, merged, industry or record.get("industry", ""))
    # the normalizers rewrite `template`; keep the record's own (case_study's normalizer
    # returns "case_study_full", which is the DRAFT template name, not the render key)
    clean.pop("template", None)
    clean.pop("content_type", None)
    return {k: v for k, v in clean.items() if k in allowed}
