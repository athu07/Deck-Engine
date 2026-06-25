"""Automatic meeting history.

Every time the engine generates a deck (/download or /finalize), one record is
saved here automatically — the salesperson does nothing extra.

Naming convention (owner's choice): one JSON file per client + phase, named
    J2W_<ClientName>_<PhaseCode>.json
The newest deck for a given client+phase OVERWRITES the previous record (only
the latest is kept). The full date & time is still stored inside the record.

Phase codes:  Pre-read=PR  First Meeting=FM  Second Meeting=SM  Proposal=PP
"""
import json
import os
import re
from datetime import datetime

MEETINGS_DIR = "meetings"

PHASE_CODES = {
    "Pre-read": "PR",
    "First Meeting": "FM",
    "Second Meeting": "SM",
    "Proposal": "PP",
}


def phase_code(phase):
    """2-letter short form for a phase; 'XX' if unknown/blank."""
    return PHASE_CODES.get(phase, "XX")


def _safe_client(name):
    """Strip spaces and illegal characters: 'Acme Bank' -> 'AcmeBank'."""
    return re.sub(r"[^A-Za-z0-9]+", "", name or "") or "Client"


def record_name(client, phase):
    """The self-explanatory record name, e.g. 'J2W_AcmeBank_FM'."""
    return f"J2W_{_safe_client(client)}_{phase_code(phase)}"


def save(client, industry, functions, work_types, phase, recipient,
         salesperson, slide_ids, deck_file):
    """Write one meeting record. Returns the file path written."""
    os.makedirs(MEETINGS_DIR, exist_ok=True)
    name = record_name(client, phase)
    record = {
        "name": name,
        "client": client,
        "industry": industry,
        "functions": functions,
        "work_types": work_types,
        "phase": phase,
        "phase_code": phase_code(phase),
        "recipient": recipient,
        "salesperson": salesperson,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "slide_ids": slide_ids,
        "deck_file": deck_file,
    }
    path = os.path.join(MEETINGS_DIR, name + ".json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(record, f, indent=2, ensure_ascii=False)
    return path


def all_meetings():
    """Every saved record, newest first (used by the 'Search past meetings' page)."""
    items = []
    if not os.path.isdir(MEETINGS_DIR):
        return items
    for fn in os.listdir(MEETINGS_DIR):
        if not fn.endswith(".json"):
            continue
        try:
            with open(os.path.join(MEETINGS_DIR, fn), encoding="utf-8") as f:
                items.append(json.load(f))
        except Exception:
            pass
    items.sort(key=lambda r: r.get("generated_at", ""), reverse=True)
    return items
