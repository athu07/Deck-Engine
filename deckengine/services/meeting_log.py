"""Versioned deck history + the deck filename convention.

Owner's naming convention (2026-07-07):
    Intro          -> "Joulestowatts_{Client} Intro.pptx"            (v1),
                      "Joulestowatts_{Client} Intro V2.pptx"          (v2+)
    First Meeting  -> "Joulestowatts_{Client} FV{n}.pptx"             (n = 1, 2, 3...)
    Second Meeting -> "Joulestowatts_{Client} SV{n}.pptx"
    Proposal       -> "Joulestowatts_{Client} PV{n}.pptx"             (not specified by the
                      owner; extends the same scheme for consistency -- flag if wrong)

One JSON file per client + phase (unchanged naming: J2W_<ClientName>_<PhaseCode>.json),
but the record now holds a VERSION HISTORY, not a single overwritten entry — every
version made for a client+phase is kept forever (owner's choice), each with its own
.pptx file on disk (never overwritten). This is what lets a salesperson reopen an
already-finalized, already-downloaded deck, add one more slide, and re-finalize as
the next version — see decks.py's /deck/reopen route.

Phase codes:  Intro=IN  First Meeting=FM  Second Meeting=SM  Proposal=PP
"""
import json
import os
import re
from datetime import datetime

from deckengine import config

MEETINGS_DIR = config.MEETINGS_DIR

PHASE_CODES = {
    "Intro": "IN",
    "First Meeting": "FM",
    "Second Meeting": "SM",
    "Proposal": "PP",
}

# the filename "version" prefix per phase (Intro is a special case -- see deck_filename)
_PHASE_FILE_PREFIX = {
    "First Meeting": "FV",
    "Second Meeting": "SV",
    "Proposal": "PV",
}


def phase_code(phase):
    """2-letter short form for a phase; 'XX' if unknown/blank."""
    return PHASE_CODES.get(phase, "XX")


def _safe_client(name):
    """Strip spaces and illegal characters: 'Acme Bank' -> 'AcmeBank'. Used for the
    JSON record filename only (not the .pptx name, which keeps natural spacing)."""
    return re.sub(r"[^A-Za-z0-9]+", "", name or "") or "Client"


def _safe_for_pptx_name(name):
    """Strip only OS-illegal filename characters; keep spaces/casing so the deck
    filename reads naturally ('Joulestowatts_Acme Bank Intro.pptx')."""
    return re.sub(r'[\\/:*?"<>|]+', "", (name or "").strip()) or "Client"


def record_name(client, phase):
    """The JSON record's own filename, e.g. 'J2W_AcmeBank_FM'."""
    return f"J2W_{_safe_client(client)}_{phase_code(phase)}"


def deck_filename(client, phase, version):
    """The .pptx filename for this client+phase+version, per the owner's convention."""
    client_disp = _safe_for_pptx_name(client)
    if phase == "Intro":
        suffix = "Intro" if version <= 1 else f"Intro V{version}"
    else:
        prefix = _PHASE_FILE_PREFIX.get(phase, "PV")
        suffix = f"{prefix}{version}"
    return f"Joulestowatts_{client_disp} {suffix}.pptx"


def _path(client, phase):
    return os.path.join(MEETINGS_DIR, record_name(client, phase) + ".json")


def _load_record(client, phase):
    try:
        with open(_path(client, phase), encoding="utf-8") as f:
            rec = json.load(f)
        rec.setdefault("versions", [])
        return rec
    except (OSError, ValueError):
        return {"client": client, "industry": "", "phase": phase,
                "phase_code": phase_code(phase), "versions": []}


def next_version_number(client, phase):
    """1 for a brand-new client+phase; len(existing)+1 otherwise."""
    return len(_load_record(client, phase).get("versions", [])) + 1


def save_version(client, industry, functions, work_types, phase, recipient,
                  salesperson, slide_ids, deck_file, edits=None, case_edits=None):
    """Append a new version to this client+phase's history. NEVER overwrites a
    prior version's entry or file — every version made is kept. Returns the new
    version dict (includes its 1-indexed `version` number)."""
    os.makedirs(MEETINGS_DIR, exist_ok=True)
    rec = _load_record(client, phase)
    version = len(rec["versions"]) + 1
    entry = {
        "version": version,
        "deck_file": deck_file,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "functions": functions,
        "work_types": work_types,
        "recipient": recipient,
        "salesperson": salesperson,
        "slide_ids": slide_ids,
        "edits": edits or {},
        "case_edits": case_edits or {},
    }
    rec["client"] = client
    rec["industry"] = industry
    rec["phase"] = phase
    rec["phase_code"] = phase_code(phase)
    rec["versions"].append(entry)
    with open(_path(client, phase), "w", encoding="utf-8") as f:
        json.dump(rec, f, indent=2, ensure_ascii=False)
    return entry


def latest_version(client, phase):
    """(record, latest_version_dict) for this client+phase, or (record, None) if
    none exist yet. Used by /deck/reopen to reload the most recent slide order."""
    rec = _load_record(client, phase)
    versions = rec.get("versions") or []
    return rec, (versions[-1] if versions else None)


def all_meetings():
    """Every version of every client+phase, newest first — one row per version
    (not per client+phase) so every historical deck stays visible/downloadable,
    matching the owner's 'keep every version forever' choice."""
    rows = []
    if not os.path.isdir(MEETINGS_DIR):
        return rows
    for fn in os.listdir(MEETINGS_DIR):
        if not fn.endswith(".json"):
            continue
        try:
            with open(os.path.join(MEETINGS_DIR, fn), encoding="utf-8") as f:
                rec = json.load(f)
        except (OSError, ValueError):
            continue
        base = {k: rec.get(k) for k in ("client", "industry", "phase", "phase_code")}
        for v in rec.get("versions", []):
            rows.append({**base, **v})
    rows.sort(key=lambda r: r.get("generated_at", ""), reverse=True)
    return rows


def all_latest():
    """One row per client+phase (the LATEST version only) — used for the
    'reopen to edit' list, so reopening always continues from the newest
    version rather than forking from an older one."""
    rows = []
    if not os.path.isdir(MEETINGS_DIR):
        return rows
    for fn in os.listdir(MEETINGS_DIR):
        if not fn.endswith(".json"):
            continue
        try:
            with open(os.path.join(MEETINGS_DIR, fn), encoding="utf-8") as f:
                rec = json.load(f)
        except (OSError, ValueError):
            continue
        versions = rec.get("versions") or []
        if not versions:
            continue
        base = {k: rec.get(k) for k in ("client", "industry", "phase", "phase_code")}
        rows.append({**base, **versions[-1]})
    rows.sort(key=lambda r: r.get("generated_at", ""), reverse=True)
    return rows
