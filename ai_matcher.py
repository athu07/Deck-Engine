# -*- coding: utf-8 -*-
"""
ai_matcher.py  --  Step 02 AI layer: refine transcript matching with judgment.

Keyword matching (in matcher.py) gives candidates; this asks the LLM to:
  - pick the case studies that GENUINELY fit the meeting (drop false positives
    that only share a generic keyword but are about a different topic), and
  - decide which OPTIONAL slides the transcript actually calls for.

Leader / people slides are handled by the caller, never here — they are never
auto-picked.

Provider: OpenAI (per project owner's choice). Key read from .env.
"""

import json

from secrets_loader import load_env

MODEL = "gpt-4o-mini"


def _client():
    load_env()
    from openai import OpenAI
    return OpenAI()                      # reads OPENAI_API_KEY from the environment


def refine(transcript, candidates_by_wt, optional_slides, top_n=3):
    """
    transcript        : the pasted meeting text
    candidates_by_wt  : {work_type: [{slide_id, title, keywords}]}
    optional_slides   : [{slide_id, title}]
    Returns {"cases": {work_type: [slide_id]}, "optional": [slide_id]}.
    """
    lines = [
        "A salesperson pasted this meeting transcript:",
        '"""', transcript[:6000], '"""', "",
        "Pick the case studies that GENUINELY match what this meeting is about.",
        "Do NOT pick a slide that only shares a generic keyword but is about a "
        "different industry or topic than the meeting.",
        "If a listed case study clearly relates to the meeting, you MUST include "
        "it — only return an empty list for a work type when none are relevant.",
        f"Pick at most {top_n} per work type. Use only the slide IDs shown.",
    ]
    for wt, rows in candidates_by_wt.items():
        lines.append(f"\nCASE STUDIES for {wt}:")
        for r in rows:
            lines.append(f"  {r['slide_id']}: {r['title']} — keywords: {r['keywords']}")
    if optional_slides:
        lines.append("\nOPTIONAL slides (include only if the transcript clearly calls for them):")
        for r in optional_slides:
            lines.append(f"  {r['slide_id']}: {r['title']}")
    lines.append(
        '\nReturn ONLY this JSON shape: '
        '{"cases": {"WORKFORCE": [ids], "AI_POD": [ids], "MS": [ids]}, "optional": [ids]}'
    )

    resp = _client().chat.completions.create(
        model=MODEL,
        temperature=0,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": "You select the most relevant sales slides "
                                          "for a meeting. Reply with one JSON object only."},
            {"role": "user", "content": "\n".join(lines)},
        ],
    )
    try:
        data = json.loads(resp.choices[0].message.content)
    except (json.JSONDecodeError, AttributeError, IndexError):
        return {"cases": {}, "optional": []}

    # The model is asked for plain slide IDs, but occasionally returns objects
    # (e.g. {"slide_id": "CS70"}). Normalise everything to a list of ID STRINGS
    # so the caller never has to guess (and never hits 'unhashable type: dict').
    def _ids(v):
        out = []
        for x in (v or []):
            if isinstance(x, str):
                out.append(x)
            elif isinstance(x, dict):
                sid = x.get("slide_id") or x.get("id")
                if sid:
                    out.append(str(sid))
        return out

    raw_cases = data.get("cases") if isinstance(data.get("cases"), dict) else {}
    cases = {wt: _ids(v) for wt, v in raw_cases.items()}
    return {"cases": cases, "optional": _ids(data.get("optional"))}


def extract_asks(transcript):
    """Pull the SPECIFIC capability / skill / technology asks the CLIENT made.

    Returns a short list of concise phrases, e.g. ['ADAS', 'fraud detection'].
    This is what lets the engine flag "X was asked but isn't in the deck" even
    for a topic that exists in NO slide yet (the whole point of the gap fix).
    Conservative by design — concrete asks only, capped, never generic words.
    Fails safe to [] on any error (the caller still has keyword detection)."""
    if not (transcript or "").strip():
        return []
    prompt = (
        "A salesperson pasted this client meeting transcript:\n"
        '"""\n' + transcript[:6000] + '\n"""\n\n'
        "List the SUBSTANTIAL capabilities, solutions, or domains the CLIENT asked "
        "for that would each justify a DEDICATED sales slide. Rules:\n"
        "- Capability/solution THEMES only (e.g. 'ADAS', 'fraud detection', "
        "'predictive maintenance', 'blockchain traceability', 'test automation').\n"
        "- Do NOT list individual tools, libraries, frameworks, or programming "
        "languages (e.g. Cucumber, Selenium, JUnit, React, Python, Docker, Jenkins). "
        "These are details, not deck themes. If the client only named tools, infer "
        "the capability they belong to (Selenium/Cucumber -> test automation) or skip.\n"
        "- Do NOT include generic words (software, team, quality, support, project, "
        "solution, technology, help, service).\n"
        "- Only things the client wants delivered — not background chit-chat.\n"
        "- At most 6 items. If none clearly merit their own slide, return an empty list.\n"
        "- Keep acronyms/proper nouns as written (ADAS, SAP); otherwise lowercase.\n"
        'Return ONLY this JSON: {"asks": ["...", "..."]}'
    )
    try:
        resp = _client().chat.completions.create(
            model=MODEL,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You extract concrete client asks from "
                                              "a sales meeting. Reply with one JSON object only."},
                {"role": "user", "content": prompt},
            ],
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        return []
    asks = data.get("asks") if isinstance(data, dict) else None
    out = []
    for a in (asks or []):
        if isinstance(a, str) and a.strip():
            out.append(a.strip())
        elif isinstance(a, dict):
            v = a.get("ask") or a.get("topic") or a.get("name")
            if v:
                out.append(str(v).strip())
    return out[:6]
