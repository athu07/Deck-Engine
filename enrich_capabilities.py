# -*- coding: utf-8 -*-
"""
enrich_capabilities.py  --  Add a one-line DESCRIPTION under every capability
title in the content store (so capability cards aren't title-only).

WHY: the Excel ships capabilities as bare titles ("Automated Approval Routing").
The slide template has room for a short body under each title. This fills that
body, grounded in the case's OWN challenge + solution, with a hard rule:
never name any real client/company — only J2W is ever named.

HOW (read top-to-bottom, nothing hidden):
  - One AI call per case (gpt-4o-mini) returns one short body per capability.
  - Capabilities change from  ["title", ...]  to  [{"title","body"}, ...].
  - The store is BACKED UP first, and writes happen after every case, so a
    re-run resumes where it left off (already-enriched cases are skipped).

This is a ONE-TIME enrichment. After it runs, fill_case_study.py is free.

Run:
  py enrich_capabilities.py --limit 3      # test on 3 cases, print samples
  py enrich_capabilities.py                # enrich every not-yet-done case
  py enrich_capabilities.py --force        # redo all (ignores existing bodies)
  py enrich_capabilities.py --ids AIP001,MSS001
"""

import io
import json
import os
import shutil
import sys

STORE = "case_study_content_store.json"
BACKUP = "case_study_content_store.backup.json"
MODEL = "gpt-4o-mini"

SYSTEM = (
    "You write concise, factual B2B consulting slide copy for JoulesToWatts "
    "(J2W). Reply with one JSON object only."
)

# The single most important rule lives first and is repeated — confidentiality.
PROMPT_TMPL = """Write a short description for each capability of ONE J2W case study.

HARD RULES (non-negotiable):
- NEVER name any real client, customer, or company. The ONLY company name
  allowed anywhere is "J2W" / "JoulesToWatts". If the source text names a
  client, refer to them generically ("the client", "the organisation").
- Ground every description ONLY in the challenge/solution below. Do NOT invent
  numbers, metrics, tools, or facts that are not implied by the source.
- Each description: 6-16 words, plain sentence case, no trailing period, no
  marketing fluff. Describe WHAT the capability did / how it worked.

CASE TITLE: {title}
DOMAIN: {domain}

CHALLENGE:
{challenge}

SOLUTION:
{solution}

CAPABILITIES (write one description for each, IN THE SAME ORDER):
{caps}

Return ONLY JSON of this exact shape (same number of items, same order):
{{"bodies": [{slots}]}}"""


def _caps_titles(cap_list):
    """Return the list of capability TITLE strings, whatever the stored shape."""
    out = []
    for c in cap_list:
        if isinstance(c, dict):
            out.append((c.get("title") or "").strip())
        else:
            out.append(str(c).strip())
    return [t for t in out if t]


def _needs_work(rec):
    caps = rec.get("capabilities", [])
    if not caps:
        return False
    # done == every capability is a dict carrying a non-empty body
    return not all(isinstance(c, dict) and (c.get("body") or "").strip()
                   for c in caps)


def _ai_bodies(rec, client):
    titles = _caps_titles(rec.get("capabilities", []))
    if not titles:
        return []
    caps_block = "\n".join(f"{i+1}. {t}" for i, t in enumerate(titles))
    slots = ", ".join('"..."' for _ in titles)
    prompt = PROMPT_TMPL.format(
        title=rec.get("title", ""),
        domain=rec.get("domain", "") or "—",
        challenge=rec.get("challenge", "") or "(not provided)",
        solution=rec.get("solution", "") or "(not provided)",
        caps=caps_block,
        slots=slots,
    )
    resp = client.chat.completions.create(
        model=MODEL, temperature=0.3,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": prompt},
        ],
    )
    data = json.loads(resp.choices[0].message.content)
    bodies = [str(b).strip().rstrip(".") for b in (data.get("bodies") or [])]
    bodies = [(b[0].upper() + b[1:]) if b else b for b in bodies]   # sentence case
    # align length to the titles (pad short, trim long)
    bodies = (bodies + [""] * len(titles))[:len(titles)]
    return list(zip(titles, bodies))


def main():
    args = sys.argv[1:]
    force = "--force" in args
    limit = None
    only_ids = None
    if "--limit" in args:
        limit = int(args[args.index("--limit") + 1])
    if "--ids" in args:
        only_ids = {s.strip() for s in args[args.index("--ids") + 1].split(",")}

    recs = json.load(open(STORE, encoding="utf-8"))

    # one-time backup (don't clobber an existing backup)
    if not os.path.exists(BACKUP):
        shutil.copy2(STORE, BACKUP)
        print(f"Backed up store -> {BACKUP}")

    from secrets_loader import load_env
    load_env()
    from openai import OpenAI
    client = OpenAI()

    todo = []
    for r in recs:
        if only_ids and r["id"] not in only_ids:
            continue
        if force or _needs_work(r):
            todo.append(r)
    if limit:
        todo = todo[:limit]

    print(f"Cases to enrich: {len(todo)}"
          + (f"  (limit {limit})" if limit else "")
          + ("  [FORCE]" if force else ""))

    done = 0
    for r in todo:
        try:
            pairs = _ai_bodies(r, client)
        except Exception as e:                     # noqa: BLE001 - report & continue
            print(f"  {r['id']}: AI error - {e}")
            continue
        r["capabilities"] = [{"title": t, "body": b} for t, b in pairs]
        done += 1
        # save after every case so a re-run resumes cleanly
        with open(STORE, "w", encoding="utf-8") as f:
            json.dump(recs, f, ensure_ascii=False, indent=2)

        if limit:                                  # test mode -> show the result
            print(f"\n{r['id']}  {r['title'][:60]}")
            for t, b in pairs:
                print(f"   - {t}")
                print(f"       {b}")
        elif done % 10 == 0:
            print(f"  ...{done}/{len(todo)}")

    print(f"\nEnriched {done} case(s). Store saved -> {STORE}")
    if not limit:
        miss = [r["id"] for r in recs if _needs_work(r)]
        if miss:
            print(f"Still title-only ({len(miss)}): {miss[:20]}"
                  + (" ..." if len(miss) > 20 else ""))
        else:
            print("Every case now has capability descriptions.")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    main()
