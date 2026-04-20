# -*- coding: utf-8 -*-
"""
local_config.py
===============
Loads credentials from .env file for local Python use (outside Colab).
Also sets up the inbox and export directories used by the scraper and
flat file generator.

Directory layout created on first run:
  ~/Documents/FlatFileGenerator/inbox/     ← scraper downloads here
  ~/Documents/FlatFileGenerator/exports/   ← flat file generator writes here

Override the root folder name by setting FLATFILE_ROOT in .env.
"""

import os
import sys
from pathlib import Path


# ── Folder layout ──────────────────────────────────────────────
# Override with FLATFILE_ROOT=/some/other/path in .env if needed
DEFAULT_ROOT = Path.home() / "Documents" / "FlatFileGenerator"


# ── Load .env file ─────────────────────────────────────────────
env_path = Path(__file__).parent / ".env"
if not env_path.exists():
    print(f"ERROR: .env file not found at {env_path}")
    print("       Copy .env.example to .env and fill in your credentials.")
    sys.exit(1)

env_vars = {}
with open(env_path) as f:
    for line in f:
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, val = line.split("=", 1)
            env_vars[key.strip()] = val.strip()
            os.environ[key.strip()] = val.strip()


# ── Resolve directories (env var wins, then default) ──────────
_root = Path(env_vars.get("FLATFILE_ROOT", str(DEFAULT_ROOT)))
INBOX_DIR    = str(_root / "inbox")
DOWNLOAD_DIR = INBOX_DIR                    # scraper alias
EXPORT_DIR   = str(_root / "exports")

os.makedirs(INBOX_DIR, exist_ok=True)
os.makedirs(EXPORT_DIR, exist_ok=True)


# ── Credentials ────────────────────────────────────────────────
HS_USERNAME = env_vars.get("HS_USERNAME", "")
HS_PASSWORD = env_vars.get("HS_PASSWORD", "")
ORS_API_KEY = env_vars.get("ORS_API_KEY", "")

REPO_DIR = str(Path(__file__).parent)

if not all([HS_USERNAME, HS_PASSWORD, ORS_API_KEY]):
    print("WARNING: Some credentials missing from .env")
else:
    print("Credentials loaded from .env")
    print(f"  Inbox:   {INBOX_DIR}")
    print(f"  Exports: {EXPORT_DIR}")
