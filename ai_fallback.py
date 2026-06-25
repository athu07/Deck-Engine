# -*- coding: utf-8 -*-
"""
ai_fallback.py  --  Box 0, AI second-opinion for the UNCERTAIN tags only.

The keyword rules (tagger.py) handle everything they're confident about.
This script sends ONLY the shaky tags to an LLM for a second opinion:
  - kind  : for STANDARD/OPTIONAL slides (rules can't see "optional-ness")
  - work_type / industry / function : for CASE_STUDY slides where the
    keyword vote was missing or too close to call.

Safety / honesty built in:
  - The API key is read from .env via secrets_loader (never hard-coded).
  - DRY RUN by default: shows what WOULD be sent, makes no API calls, costs
    nothing. Add --apply to actually call the API.
  - The model may only choose from the Tag Dictionary's allowed values;
    anything else is ignored and the rule's answer is kept.
  - A tag a human has locked (confidence == HUMAN) is never touched.
  - AI answers are still marked confidence = AUTO (a machine chose them),
    with source = "AI" so you can see which came from the model.

Provider: OpenAI (per project owner's choice). Model is set below.
"""

import json
import sys

import tagger
from secrets_loader import load_env

MODEL = "gpt-4o-mini"   # cheap + good for classification; change here if you like

KINDS = ["STANDARD", "CASE_STUDY", "DIVIDER", "OPTIONAL"]
WORK_TYPES = list(tagger.WORK_TYPE.keys())     # WORKFORCE / AI_POD / MS
INDUSTRIES = list(tagger.INDUSTRY.keys())
FUNCTIONS = list(tagger.FUNCTION.keys())

ALLOWED = {
    "kind": KINDS,
    "work_type": WORK_TYPES,
    "industry": INDUSTRIES,
    "function": FUNCTIONS,
}


def _text_of(rec):
    t = rec.get("title", "")
    s = rec.get("subtitle", "")
    return " " + (rec.get("full_text", "") + " " + t + " " + s).lower() + " "


def _uncertain(text, keyword_map):
    """True if the keyword vote is missing or within 1 of the runner-up."""
    _, votes = tagger._score(text, keyword_map)
    if not votes:
        return True
    ranked = sorted(votes.values(), reverse=True)
    top = ranked[0]
    second = ranked[1] if len(ranked) > 1 else 0
    return (top - second) <= 1


def review_plan(rec):
    """Which dimensions of THIS slide deserve a second opinion."""
    text = _text_of(rec)
    kind = rec["tags"]["kind"]["value"]
    dims = []
    if kind in ("STANDARD", "OPTIONAL"):
        dims.append("kind")                       # catch the OPTIONAL slides
    elif kind == "CASE_STUDY":
        if _uncertain(text, tagger.WORK_TYPE):
            dims.append("work_type")
        if _uncertain(text, tagger.INDUSTRY):
            dims.append("industry")
        dims.append("function")                   # weakest tag — always review
    # don't review locked-by-human tags
    return [d for d in dims if rec["tags"].get(d, {}).get("confidence") != "HUMAN"]


def _prompt(rec, dims):
    lines = [
        "You are tagging a sales slide for a slide-picking engine.",
        "Pick the BEST value for each field from its allowed list. "
        "If none truly fits, return null. Do not invent values.",
        "",
        f"Slide title: {rec.get('title','')}",
        f"Slide subtitle: {rec.get('subtitle','')}",
        f"Slide text (truncated): {rec.get('full_text','')[:1200]}",
        "",
        "Return ONLY a JSON object with these fields and allowed values:",
    ]
    for d in dims:
        lines.append(f"  {d}: one of {ALLOWED[d]} or null")
    return "\n".join(lines)


def ask_model(client, rec, dims):
    resp = client.chat.completions.create(
        model=MODEL,
        temperature=0,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": "You are a precise classification assistant. "
                                          "Reply with a single JSON object, nothing else."},
            {"role": "user", "content": _prompt(rec, dims)},
        ],
    )
    raw = resp.choices[0].message.content
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return {}
    # keep only requested dims, and only allowed values (or None)
    out = {}
    for d in dims:
        v = data.get(d)
        if v in ALLOWED[d] or v is None:
            out[d] = v
    return out


def main(apply):
    recs = json.load(open("tagged_library.json", encoding="utf-8"))

    plans = [(r, review_plan(r)) for r in recs]
    todo = [(r, dims) for r, dims in plans if dims]
    total_calls = len(todo)
    by_dim = {}
    for _, dims in todo:
        for d in dims:
            by_dim[d] = by_dim.get(d, 0) + 1

    print(f"Slides needing a second opinion: {total_calls} of {len(recs)}")
    print(f"Tags to review by dimension     : {by_dim}")
    print(f"Model                           : {MODEL}")
    print()

    if not apply:
        print("DRY RUN — no API calls made, nothing spent.")
        print("Examples of what would be sent:")
        for r, dims in todo[:8]:
            print(f"  {r['slide_id']:5} {r['title'][:42]:42} -> review {dims}")
        print()
        print("When ready, run for real with:  py ai_fallback.py --apply")
        return

    # ---- LIVE: call the API ----
    load_env()
    from openai import OpenAI
    client = OpenAI()   # reads OPENAI_API_KEY from the environment

    changed = 0
    for i, (rec, dims) in enumerate(todo, 1):
        answer = ask_model(client, rec, dims)
        for d, new_val in answer.items():
            old = rec["tags"][d]["value"]
            rec["tags"][d] = {"value": new_val, "confidence": "AUTO", "source": "AI"}
            if new_val != old:
                changed += 1
                print(f"  {rec['slide_id']:5} {d:10} {str(old):16} -> {new_val}")
        print(f"[{i}/{total_calls}] {rec['slide_id']} done", file=sys.stderr)

    json.dump(recs, open("tagged_library.json", "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)
    print()
    print(f"Applied AI second opinions. Tags changed: {changed}. Saved -> tagged_library.json")


if __name__ == "__main__":
    main(apply="--apply" in sys.argv)
