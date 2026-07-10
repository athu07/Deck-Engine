# -*- coding: utf-8 -*-
"""
dedupe.py  --  "do we already have a slide for this?"

Before the Custom Slide Builder turns pasted content into a slide, it asks this
module whether the library already contains the same story. Rebuilding a case study
we already own is pure waste: it burns an AI call, it adds a near-duplicate record
that makes every future match noisier, and it hides a slide the team already trusts.

The check is deliberately cheap and reuses what already exists: ONE embedding call
for the pasted text (relevance.embed_texts), cosined against the per-case vectors
already sitting in data/case_embeddings.json. No new model, no new index.

FAIL-SAFE, like every AI path in this app: no API key, no embeddings file, or any
exception -> no matches, and the caller builds the slide as if we'd never asked.
A duplicate check must never be the reason a salesperson can't build a slide.
"""

from deckengine.services.content import case_library
from deckengine.services.matching import relevance

# At or above this cosine, the pasted content is "the same story" as an existing case
# and is worth interrupting the salesperson for. Set by the owner (2026-07-10) after
# weighing the two failure modes: a missed duplicate merely wastes a slide, while a
# false alarm on every paste trains people to click straight through the warning.
# Below this, nothing is shown at all -- silence, not a weak suggestion.
THRESHOLD = 0.80

# Never show more than a handful; the salesperson is deciding "reuse or build", not
# browsing the library (that's what /library is for).
TOP_N = 3


def similar_cases(text, top_n=TOP_N, threshold=THRESHOLD, allowed_work_types=None):
    """Content-store cases whose meaning matches `text` at or above `threshold`,
    best first. Returns [] when nothing clears the bar, and [] on any failure.

    `allowed_work_types` (e.g. {"MS"}) narrows the search to one work type -- the
    builder passes the work type chosen for the slide being pasted, since an MS case
    is not a duplicate of a Workforce one even when the words line up.

    Each match: {"id", "title", "work_type", "domain", "score"} with score in 0..1.
    """
    return similar_cases_many([text], top_n, threshold, allowed_work_types)[0]


def similar_cases_many(texts, top_n=TOP_N, threshold=THRESHOLD, allowed_work_types=None):
    """similar_cases for SEVERAL texts in ONE embedding call -- the pasted-document
    flow checks every slide before building any of them, and N round-trips for N
    slides is waste when the API takes a list.

    Returns a list of match-lists, positionally aligned with `texts` (so an empty or
    unembeddable text still gets its own [] slot). [] everywhere on any failure.
    """
    texts = list(texts or [])
    if not texts:
        return []
    empty = [[] for _ in texts]

    # embed_texts drops blanks, so embed only the non-blank ones and map back by position
    fill = [i for i, t in enumerate(texts) if (t or "").strip()]
    if not fill:
        return empty
    vecs = relevance.embed_texts([texts[i] for i in fill])
    if not vecs or len(vecs) != len(fill):
        return empty                    # offline / no key -> the check simply doesn't run

    case_vectors = relevance._load_case_embeddings()
    if not case_vectors:
        return empty                    # embeddings never built -> nothing to compare against

    wanted = {str(w).strip().upper() for w in (allowed_work_types or []) if str(w).strip()}
    cases = [rec for rec in case_library._load()
             if not wanted or (rec.get("work_type") or "").upper() in wanted]

    out = empty
    for slot, query in zip(fill, vecs):
        hits = []
        for rec in cases:
            vec = case_vectors.get(rec["id"])
            if not vec:
                continue                # a case added since the last embeddings rebuild
            score = relevance._cosine(query, vec)
            if score >= threshold:
                hits.append({"id": rec["id"], "title": rec.get("title", ""),
                             "work_type": rec.get("work_type", ""),
                             "domain": rec.get("domain", ""),
                             "score": round(score, 4)})
        hits.sort(key=lambda h: h["score"], reverse=True)
        out[slot] = hits[:top_n]
    return out
