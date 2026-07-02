# -*- coding: utf-8 -*-
"""
build_case_embeddings.py  --  One vector per case study, for meaning-based match.

Reads case_study_content_store.json, builds a rich text per case (title +
challenge + solution + capabilities + keywords via case_library._search_text),
asks OpenAI for an embedding, and writes case_embeddings.json:

    { "model": "...", "dim": 1536, "vectors": { "AIP001": [ ... ], ... } }

relevance.py loads this at match time and compares it to the meeting notes'
embedding by cosine similarity. This is what lets "certify AI agents before
deployment" find the AI-governance case even with no shared keywords.

One-time / re-runnable (cheap: ~160 short inputs on text-embedding-3-small).
Run after any change to the store's case text:
    py build_case_embeddings.py
"""

import io
import json
import sys

import case_library

STORE = "case_study_content_store.json"
OUT = "case_embeddings.json"
MODEL = "text-embedding-3-small"
BATCH = 100


def main():
    recs = json.load(open(STORE, encoding="utf-8"))
    ids = [r["id"] for r in recs]
    texts = [case_library._search_text(r) for r in recs]
    print(f"Embedding {len(texts)} case studies with {MODEL} ...")

    from secrets_loader import load_env
    load_env()
    from openai import OpenAI
    client = OpenAI()

    vectors = {}
    dim = None
    for i in range(0, len(texts), BATCH):
        chunk_ids = ids[i:i + BATCH]
        chunk_txt = [t[:8000] for t in texts[i:i + BATCH]]
        resp = client.embeddings.create(model=MODEL, input=chunk_txt)
        for cid, item in zip(chunk_ids, resp.data):
            vectors[cid] = item.embedding
            dim = len(item.embedding)
        print(f"  ...{min(i + BATCH, len(texts))}/{len(texts)}")

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"model": MODEL, "dim": dim, "vectors": vectors}, f)
    print(f"\nSaved {len(vectors)} vectors (dim {dim}) -> {OUT}")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    main()
