# -*- coding: utf-8 -*-
"""
secrets_loader.py  --  loads the .env file into the environment.

Why: keeping the API key in .env (which is git-ignored) means the secret
never lives inside our code and never gets committed or shared. Any script
can call load_env() at startup, then read the key via os.environ.

No external packages needed — this parses .env by hand.
"""

import os
from pathlib import Path


def load_env(path=".env"):
    """Read KEY=value lines from .env into os.environ (without overwriting
    a value that's already set in the real environment)."""
    p = Path(path)
    if not p.exists():
        return
    for raw in p.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and value and not os.environ.get(key):
            os.environ[key] = value


def get_key(name="ANTHROPIC_API_KEY"):
    """Convenience: load .env then return the named key (or None)."""
    load_env()
    return os.environ.get(name)


if __name__ == "__main__":
    # Safe self-check: tells you whether a key is present WITHOUT printing it.
    key = get_key("OPENAI_API_KEY")
    if not key:
        print("No OPENAI_API_KEY found yet. Open .env and paste your key after 'OPENAI_API_KEY='.")
    elif key.startswith("sk-"):
        print(f"OpenAI key loaded OK (starts with 'sk-', length {len(key)}).")
    else:
        print(f"A key is loaded (length {len(key)}), but it doesn't look like an OpenAI key (those start with 'sk-').")
