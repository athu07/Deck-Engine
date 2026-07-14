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
    """1 for a brand-new client+phase; len(existing)+1 otherwise. Read-only --
    for display/preview purposes. reserve_version() is what /finalize actually
    uses to CLAIM a number (see its docstring for why the two are different)."""
    return len(_real(_load_record(client, phase).get("versions"))) + 1


def _write_record(client, phase, rec):
    os.makedirs(MEETINGS_DIR, exist_ok=True)
    with open(_path(client, phase), "w", encoding="utf-8") as f:
        json.dump(rec, f, indent=2, ensure_ascii=False)


def _with_lock(client, phase, fn):
    """Run `fn()` while holding an exclusive lock file for this client+phase.

    os.O_CREAT | os.O_EXCL is atomic even across separate OS processes (unlike a
    plain Python threading.Lock, which only protects threads within ONE process --
    this app can run under gunicorn with several worker processes in production,
    per the deploy docs). Short critical section: only the read-modify-write of
    the JSON record happens inside it, never the actual deck assembly."""
    os.makedirs(MEETINGS_DIR, exist_ok=True)
    lock_path = _path(client, phase) + ".lock"
    import time
    for _ in range(50):                       # ~5s max wait, then give up loudly
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            break
        except FileExistsError:
            time.sleep(0.1)
    else:
        raise RuntimeError("Could not get the version lock for %s / %s -- "
                           "another finalize may be stuck." % (client, phase))
    try:
        return fn()
    finally:
        try:
            os.remove(lock_path)
        except OSError:
            pass


def reserve_version(client, phase):
    """Atomically claim the next version number AND its filename for this
    client+phase, writing a placeholder record immediately -- BEFORE the
    (multi-second) deck assembly runs. Returns (version, filename).

    Closes a real race: the old flow read "how many versions exist" once, spent
    several seconds building the deck, and only wrote the new version record
    afterwards. Two near-simultaneous /finalize calls for the SAME client+phase
    (a double-click, two open tabs, two people finalizing the same account) could
    both read the same starting count and be handed the SAME version number and
    filename -- the second one to finish would silently overwrite the first
    one's real .pptx bytes on disk, with no error to either salesperson (owner-
    reported, 2026-07-14: a real client's older version had vanished this way).
    Call finalize_version() once the deck is actually built, or
    cancel_reservation() if the build fails, so a failed attempt doesn't
    permanently burn a version number."""
    def _do():
        rec = _load_record(client, phase)
        version = len(rec["versions"]) + 1
        filename = deck_filename(client, phase, version)
        rec["client"] = client
        rec["phase"] = phase
        rec["phase_code"] = phase_code(phase)
        rec["versions"].append({"version": version, "deck_file": filename,
                                "generated_at": None, "_reserved": True})
        _write_record(client, phase, rec)
        return version, filename
    return _with_lock(client, phase, _do)


def finalize_version(client, phase, version, industry, functions, work_types,
                      recipient, salesperson, slide_ids, edits=None, case_edits=None):
    """Fill in the reservation's real fields once the deck has actually been
    built. The filename was already claimed by reserve_version() and never
    changes here."""
    def _do():
        rec = _load_record(client, phase)
        rec["industry"] = industry
        for v in rec["versions"]:
            if v.get("version") == version:
                v.update({
                    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "functions": functions, "work_types": work_types,
                    "recipient": recipient, "salesperson": salesperson,
                    "slide_ids": slide_ids, "edits": edits or {},
                    "case_edits": case_edits or {},
                })
                v.pop("_reserved", None)
                break
        _write_record(client, phase, rec)
    _with_lock(client, phase, _do)


def cancel_reservation(client, phase, version):
    """The build failed after reserve_version() claimed a number -- drop that
    placeholder so the version number (and its filename) is free to be tried
    again, instead of leaving a permanent gap with no real deck behind it."""
    def _do():
        rec = _load_record(client, phase)
        rec["versions"] = [v for v in rec["versions"]
                           if not (v.get("version") == version and v.get("_reserved"))]
        _write_record(client, phase, rec)
    _with_lock(client, phase, _do)


def _real(versions):
    """Drop reservation placeholders -- a version whose build is still running
    (or crashed before cancel_reservation() could clean it up) has no real
    generated_at/slide_ids yet and must never appear as a phantom deck."""
    return [v for v in (versions or []) if not v.get("_reserved")]


def latest_version(client, phase):
    """(record, latest_version_dict) for this client+phase, or (record, None) if
    none exist yet. Used by /deck/reopen to reload the most recent slide order."""
    rec = _load_record(client, phase)
    versions = _real(rec.get("versions"))
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
        for v in _real(rec.get("versions")):
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
        versions = _real(rec.get("versions"))
        if not versions:
            continue
        base = {k: rec.get(k) for k in ("client", "industry", "phase", "phase_code")}
        rows.append({**base, **versions[-1]})
    rows.sort(key=lambda r: r.get("generated_at", ""), reverse=True)
    return rows
