"""
reset.py — Integrity Wall Systems PO Automation Reset Tool
Clears all processed data so you can test from a clean state.
Run with: python reset.py
"""

import os
import shutil
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

STORAGE_ROOT = os.getenv("STORAGE_ROOT", "./po-automation/data")

TARGETS = [
    ("FILE", os.path.join(STORAGE_ROOT, "po_tracking.csv")),
    ("DIR",  os.path.join(STORAGE_ROOT, "processed")),
    ("DIR",  os.path.join(STORAGE_ROOT, "flagged")),
    ("DIR",  os.path.join(STORAGE_ROOT, "inbox")),
    ("DIR",  os.path.join(STORAGE_ROOT, "archive")),
]

RECREATE = [
    os.path.join(STORAGE_ROOT, "inbox"),
    os.path.join(STORAGE_ROOT, "processed"),
    os.path.join(STORAGE_ROOT, "flagged"),
    os.path.join(STORAGE_ROOT, "archive"),
]

print("=" * 50)
print("  IWS PO Automation — Reset Tool")
print("=" * 50)
print(f"\nStorage root: {os.path.abspath(STORAGE_ROOT)}")
print("\nThe following will be deleted:")
for kind, path in TARGETS:
    exists = os.path.exists(path)
    print(f"  {'[FILE]' if kind == 'FILE' else '[DIR] '} {path}  {'exists' if exists else '— not found'}")

print()
confirm = input("Type YES to confirm reset: ").strip()

if confirm != "YES":
    print("Reset cancelled.")
    exit(0)

print()
deleted = 0
for kind, path in TARGETS:
    if not os.path.exists(path):
        print(f"  SKIP   {path} (not found)")
        continue
    try:
        if kind == "FILE":
            os.remove(path)
        else:
            shutil.rmtree(path, ignore_errors=True)
        print(f"  DELETE {path}")
        deleted += 1
    except Exception as e:
        print(f"  ERROR  {path} — {e}")

# Recreate required empty folders so automation can run immediately
print()
print("Recreating required folders...")
for folder in RECREATE:
    Path(folder).mkdir(parents=True, exist_ok=True)
    print(f"  OK     {folder}")

print()
print(f"Done. {deleted} item(s) removed, folders recreated.")
print("You can now run the automation fresh.")
print("=" * 50)