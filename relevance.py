# -*- coding: utf-8 -*-
"""
relevance.py  --  How well does a case study match this meeting?

The rebuilt scoring core. For each case it blends:
  - SEMANTIC similarity : meaning-based match of the notes vs the whole case
    (title + challenge + solution + capabilities), via OpenAI embeddings. This is
    what catches paraphrased asks and codenames that share no literal words.
  - LEXICAL overlap      : words/phrases shared between the notes and the full
    case body (title + keywords weigh a little more). Works fully offline.
  - INDUSTRY / FUNCTION / PERSONA / WORK-TYPE : soft boosts, never hard filters.

Nothing here filters cases out — it RANKS all of them and hands the ordered list
back to matcher.py, which decides how many to keep.

Embeddings are optional and fail-safe: no API key / offline / no embeddings file
=> semantic term is simply 0 and the lexical+boost score still ranks sensibly.
"""

import json
import math
import re

import personas
import synonyms

EMB_FILE = "case_embeddings.json"
EMB_MODEL = "text-embedding-3-small"

# generic words that carry no matching signal
_STOP = set("""
the a an and or of to in for with on at by from into your our their they them this
that is are be as it we you i he she his her mr ms dr client company companies global
leading across over more most real time using use build built cut during before after
per not no yes them then than out up down about within without also plus via each any
all we our new key value team teams work works working solution solutions service
services platform platforms system systems tool tools data set sets need needs want
wants help so his her him thread say what who whom which will would can could should
""".split())

# ── weights (tuned against the Ericsson scorecard) ───────────────────────────
# Meaning match is the PRIMARY signal. FUNCTION/INDUSTRY fit is applied as a
# TIER bonus (a case that solves the person's functional problem beats one that's
# merely in their industry — see the two-axis matching model):
#   Tier 1 = function AND industry    Tier 2 = function only (still strong)
#   Tier 3 = industry only            Tier 4 = neither
W_SEMANTIC = 10.0    # best per-ask meaning match (0..1 cosine) — the primary signal
W_LEXICAL  = 2.5     # word/phrase overlap (0..1 normalised)
W_TITLE    = 0.7     # a topic named in the case TITLE strongly implies it's about it
TITLE_CAP  = 3
W_INDUSTRY = 1.5     # the account's OWN industry is preferred
W_FUNCTION = 0.8     # a case matching the stakeholder's function
W_PERSONA  = 0.3     # multiplier on personas.score_boost
W_METRIC   = 0.4     # a CXO gets a nudge toward business-outcome proof points
LEX_FULL   = 6       # this many weighted lexical hits == a full lexical score
# a case from ANOTHER industry is only eligible if its content match is this strong
# (same-industry cases are always eligible) — blocks random unrelated-industry picks
CROSS_INDUSTRY_MIN = 0.42
# how much a MAIL-thread-only match counts vs a DEEP-RESEARCH match (research leads)
MAIL_WEIGHT = 0.7

# outcome/business metrics that resonate with an executive audience
_OUTCOME = ("roi", "cost", "saving", "revenue", "margin", "risk", "spend",
            "downtime", "compliance", "efficiency", "productivity")


# ── tokenisation ─────────────────────────────────────────────────────────────
def _tokens(text):
    """Unigrams (>=3 chars, non-stop — keeps acronyms like GCC/ODC/ERP) + bigrams."""
    words = re.findall(r"[a-z0-9]+", (text or "").lower())
    uni = {w for w in words if len(w) >= 3 and w not in _STOP}
    bi = {f"{words[i]} {words[i+1]}"
          for i in range(len(words) - 1)
          if words[i] not in _STOP and words[i + 1] not in _STOP
          and len(words[i]) >= 3 and len(words[i + 1]) >= 3}
    return uni | bi


# generic business words that don't identify a FUNCTION — ignored when matching a
# person's skill to a case (so "procurement" matches, but "cost/optimization" don't)
_GENERIC_BIZ = {
    "cost", "optimization", "management", "intelligence", "platform", "solution",
    "solutions", "services", "service", "transformation", "automation", "enterprise",
    "global", "system", "systems", "operations", "operational", "strategy", "framework",
    "model", "build", "support", "engineering", "technology", "data", "analytics",
    "digital", "business", "capability", "capabilities", "delivery", "innovation",
    "planning", "process", "quality", "team", "teams", "project", "projects", "leading",
}


def specific_terms(text):
    """Discriminating single-word terms of `text` — the ones that name a real
    function/skill (procurement, contract, gcc, facilities), not generic filler."""
    return {t for t in _tokens(text) if " " not in t and t not in _GENERIC_BIZ}


def transcript_terms(transcript):
    """Terms from the notes, plus synonym expansion of the single words so a
    paraphrase ('observability' vs 'monitoring') still overlaps lexically."""
    terms = _tokens(transcript)
    expanded = set(terms)
    for t in terms:
        if " " not in t:
            expanded |= {s for s in synonyms.expand(t) if len(s) >= 4}
    return expanded


def lexical_hits(terms, row):
    """Weighted count of note-terms present in the case (title/keywords count 2x).
    Returns (weighted_count, matched_terms)."""
    body = row.get("_body_tokens")
    if body is None:
        body = _tokens(row.get("search_text", ""))
        row["_body_tokens"] = body
    head = row.get("_head_tokens")
    if head is None:
        head = _tokens(row.get("title", "") + " " + row.get("keywords", ""))
        row["_head_tokens"] = head
    matched = terms & body
    weighted = len(matched) + len(matched & head)   # +1 extra for head matches
    return weighted, matched


# ── embeddings (semantic) ────────────────────────────────────────────────────
_case_emb = None


def _load_case_embeddings():
    global _case_emb
    if _case_emb is None:
        try:
            with open(EMB_FILE, encoding="utf-8") as f:
                data = json.load(f)
            _case_emb = data.get("vectors", data)   # {id: [floats]}
        except (OSError, ValueError):
            _case_emb = {}
    return _case_emb


def _split_asks(transcript):
    """Break the notes into individual asks so each can be matched on its own.
    A multi-solution email is a blurry average as ONE embedding; per-ask matching
    lets a case match its single most relevant paragraph. Splits on blank lines,
    then falls back to sentence windows if the notes are one block."""
    t = (transcript or "").strip()
    if not t:
        return []
    chunks = [c.strip() for c in re.split(r"\n\s*\n", t) if len(c.strip()) >= 25]
    if len(chunks) <= 1:                      # no paragraphs -> window sentences
        sents = [s.strip() for s in re.split(r"(?<=[.!?])\s+", t) if len(s.strip()) >= 20]
        chunks = [" ".join(sents[i:i + 2]) for i in range(0, len(sents), 2)] or [t]
    return chunks[:24]                        # sane cap on embed calls


def ask_count(transcript):
    """How many distinct asks the notes contain — used to size the deck (a
    one-topic meeting gets a few proofs; a multi-solution email gets many)."""
    return len(_split_asks(transcript))


def embed_texts(texts):
    """Embed a list of strings in one call; returns list of vectors or None."""
    texts = [t for t in (texts or []) if t and t.strip()]
    if not texts:
        return None
    try:
        from secrets_loader import load_env
        load_env()
        from openai import OpenAI
        client = OpenAI()
        resp = client.embeddings.create(model=EMB_MODEL, input=[t[:6000] for t in texts])
        return [d.embedding for d in resp.data]
    except Exception:
        return None                    # offline / no key -> semantic simply off


def _cosine(a, b):
    if not a or not b:
        return 0.0
    dot = sum(x * y for x, y in zip(a, b))
    na = math.sqrt(sum(x * x for x in a))
    nb = math.sqrt(sum(y * y for y in b))
    if na == 0 or nb == 0:
        return 0.0
    return dot / (na * nb)


# ── the scorer ───────────────────────────────────────────────────────────────
def _has_outcome_metrics(row):
    text = (" ".join(row.get("_record", {}).get("results", []))
            + " " + row.get("search_text", "")).lower()
    return any(term in text for term in _OUTCOME)


def rank_cases(transcript, rows, *, industry="", functions=None,
               persona_codes=(), wanted=None, cxo=False, use_semantic=True, research=""):
    """Rank every case row (best first). The DEEP RESEARCH drives: a case that
    matches the research outranks one that only matches the generic mail thread.
    Returns dicts {row, score, sem, lex, matched, industry_hit, function_hit,
    eligible, persona_why}. NOTHING is filtered — matcher decides the cut."""
    functions = {f.upper() for f in (functions or set())}
    wanted = {w.upper() for w in (wanted or set())}
    industry = (industry or "").upper()

    terms = transcript_terms((transcript or "") + " " + (research or ""))
    # research asks weigh full; mail asks are down-weighted so the research leads
    res_chunks = _split_asks(research) if (use_semantic and (research or "").strip()) else []
    mail_chunks = _split_asks(transcript) if (use_semantic and (transcript or "").strip()) else []
    all_vecs = embed_texts(res_chunks + mail_chunks) if (res_chunks or mail_chunks) else None
    res_vecs = all_vecs[:len(res_chunks)] if all_vecs else []
    mail_vecs = all_vecs[len(res_chunks):] if all_vecs else []
    case_emb = _load_case_embeddings() if all_vecs else {}

    out = []
    for row in rows:
        weighted, matched = lexical_hits(terms, row)
        lex_norm = min(1.0, weighted / LEX_FULL)

        # a topic named in the case TITLE strongly implies the case is about it
        title_toks = row.get("_title_tokens")
        if title_toks is None:
            title_toks = _tokens(row.get("title", ""))
            row["_title_tokens"] = title_toks
        title_boost = min(TITLE_CAP, len(terms & title_toks)) * W_TITLE

        sem = 0.0
        if all_vecs:
            v = case_emb.get(row["slide_id"])
            if v:                          # best-matching ask; research weighs full
                sem_res = max((max(0.0, _cosine(a, v)) for a in res_vecs), default=0.0)
                sem_mail = max((max(0.0, _cosine(a, v)) for a in mail_vecs), default=0.0)
                sem = max(sem_res, MAIL_WEIGHT * sem_mail) if res_vecs else sem_mail

        ind_hit = bool(industry and (row.get("primary_industry") or "").upper() == industry)
        fn_hit = bool(functions and (row.get("primary_function") or "").upper() in functions)
        # same-industry cases are always eligible; a cross-industry case only if its
        # CONTENT match is strong (meaning or a solid title/keyword overlap)
        eligible = ind_hit or sem >= CROSS_INDUSTRY_MIN or weighted >= LEX_FULL
        p_boost, p_why = personas.score_boost(persona_codes, row)
        metric = W_METRIC if (cxo and _has_outcome_metrics(row)) else 0.0

        score = (W_SEMANTIC * sem
                 + W_LEXICAL * lex_norm
                 + title_boost
                 + (W_INDUSTRY if ind_hit else 0.0)
                 + (W_FUNCTION if fn_hit else 0.0)
                 + W_PERSONA * p_boost
                 + metric)

        out.append({
            "row": row, "score": score, "sem": sem, "lex": weighted,
            "matched": matched, "industry_hit": ind_hit, "function_hit": fn_hit,
            "eligible": eligible, "persona_why": p_why,
        })

    out.sort(key=lambda d: -d["score"])
    return out


def semantic_available():
    """True if a case-embeddings file is present (so the UI can say 'meaning match on')."""
    return bool(_load_case_embeddings())


_head_index = None


def _case_head_index():
    """{id: token set of (title + keywords)} — a lexical view of each case, for
    the 'do our cases even mention these words' check."""
    global _head_index
    if _head_index is None:
        import case_library
        _head_index = {r["id"]: _tokens(r.get("title", "") + " " + " ".join(r.get("keywords", [])))
                       for r in case_library._load()}
    return _head_index


def lexically_covered(text, min_overlap=0.6, allowed_ids=None):
    """True if some case's title/keywords already contain most of the words in
    `text`. If allowed_ids is given, only those cases count (work-type aware)."""
    toks = {t for t in _tokens(text) if " " not in t}
    if not toks:
        return False
    need = max(2, math.ceil(min_overlap * len(toks)))
    for cid, ht in _case_head_index().items():
        if allowed_ids is not None and cid not in allowed_ids:
            continue
        if len(toks & ht) >= need:
            return True
    return False


def coverage(texts):
    """For each text, (best_case_id, best_cosine) against the case embeddings —
    i.e. how well our library already covers that need. Empty/offline -> zeros."""
    embs = _load_case_embeddings()
    texts = list(texts or [])
    if not embs or not texts:
        return [(None, 0.0)] * len(texts)
    qv = embed_texts(texts)
    if not qv:
        return [(None, 0.0)] * len(texts)
    out = []
    for v in qv:
        best_id, best = None, -1.0
        for cid, cv in embs.items():
            c = _cosine(v, cv)
            if c > best:
                best, best_id = c, cid
        out.append((best_id, max(0.0, best)))
    return out


_case_meta = None


def _meta():
    global _case_meta
    if _case_meta is None:
        import case_library
        _case_meta = {r["id"]: ((r.get("industry") or "").upper(),
                                (r.get("function") or "").upper(),
                                _tokens(r.get("title", "")))
                      for r in case_library._load()}
    return _case_meta


def best_cases(texts, industry="", functions=None, allowed_ids=None):
    """For each text (a skill/need), the best case. A DIRECT term match — the
    skill's specific words appearing in the case TITLE/keywords (procurement,
    contract, gcc) — dominates, because for B2B functional terms it's far more
    reliable than noisy embedding similarity; semantic + function + industry
    refine. If allowed_ids is given, only those cases count (work-type aware).
    Returns [(case_id, adjusted_score, raw_cosine, title_hits)] — title_hits is the
    strong 'covered' signal (skill term in the case TITLE, not just a keyword)."""
    embs = _load_case_embeddings()
    meta = _meta()
    heads = _case_head_index()
    industry = (industry or "").upper()
    functions = {f.upper() for f in (functions or set())}
    texts = list(texts or [])
    if not embs or not texts:
        return [(None, 0.0, 0.0, 0)] * len(texts)
    qv = embed_texts(texts)
    if not qv:
        return [(None, 0.0, 0.0, 0)] * len(texts)
    out = []
    for text, v in zip(texts, qv):
        need = specific_terms(text)                   # the discriminating skill words
        best = (None, -1.0, 0.0, 0)
        for cid, cv in embs.items():
            if allowed_ids is not None and cid not in allowed_ids:
                continue
            c = _cosine(v, cv)
            ind, fn, ttoks = meta.get(cid, ("", "", set()))
            htoks = heads.get(cid, set())
            direct = len(need & htoks)                # skill term in title/keywords
            direct_title = len(need & ttoks)          # ...and specifically the title
            adj = (c
                   + 0.35 * min(3, direct)            # DIRECT term match dominates
                   + 0.25 * min(2, direct_title)
                   + (0.10 if fn in functions else 0.0)
                   + (0.02 if industry and ind == industry else 0.0))  # light tiebreak, not override
            if adj > best[1]:
                best = (cid, adj, c, direct_title)
        out.append(best)
    return out


def max_similarity(sid, others):
    """Highest case-to-case cosine between `sid` and any id in `others` (0 if
    no embeddings). Used to demote near-duplicate cases during selection."""
    emb = _load_case_embeddings()
    v = emb.get(sid)
    if not v or not others:
        return 0.0
    best = 0.0
    for o in others:
        vo = emb.get(o)
        if vo:
            c = _cosine(v, vo)
            if c > best:
                best = c
    return best
