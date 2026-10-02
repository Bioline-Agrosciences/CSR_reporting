"""
One-off helper: searches your OneDrive/SharePoint-synced folders for the
Working Hours file and the BlueKanGo accidents export, so you don't have to
hunt for their exact paths by hand. Prints what it finds, ready to paste
into config.py (WORKING_HOURS_FILE and ACCIDENTS_DIR).

Not part of the regular monthly pipeline — run it once, fix config.py, then
forget about it (or run it again if a file ever gets moved).

Usage
-----
    uv run scripts/locate_safety_files.py

Searches under your home folder, skipping folders it can't read (OneDrive
"online-only" placeholder files are still found by name — Windows lists
them even before they're downloaded). Can take a minute on a large OneDrive.
"""
import os
from pathlib import Path

WORKING_HOURS_HINT = "working hours"
ACCIDENTS_HINT = "accidents_du_travail_bioline"


def search(root: Path):
    hours_matches, accidents_matches = [], []
    for dirpath, dirnames, filenames in os.walk(root, onerror=lambda e: None):
        # Skip a few folder trees that are large, irrelevant, and slow to walk.
        dirnames[:] = [d for d in dirnames if d not in
                       {".git", "node_modules", ".venv", "__pycache__", "AppData"}]
        for name in filenames:
            if not name.lower().endswith((".xlsx", ".xlsm")):
                continue
            lower = name.lower()
            path = Path(dirpath) / name
            if WORKING_HOURS_HINT in lower:
                hours_matches.append(path)
            elif ACCIDENTS_HINT in lower:
                accidents_matches.append(path)
    return hours_matches, accidents_matches


def main():
    root = Path.home()
    print(f"Searching under {root} (this can take a minute)...\n")
    hours_matches, accidents_matches = search(root)

    print("Working hours file — candidate(s) found:")
    if hours_matches:
        for p in hours_matches:
            print(f"  {p}")
        print(f"\n-> paste this into config.py:\n   WORKING_HOURS_FILE = Path(r\"{hours_matches[0]}\")")
    else:
        print("  (none found — check the file is actually synced to this machine)")

    print("\nBlueKanGo accidents export — candidate(s) found:")
    if accidents_matches:
        folders = sorted({p.parent for p in accidents_matches})
        for p in accidents_matches:
            print(f"  {p}")
        print(f"\n-> paste this into config.py:\n   ACCIDENTS_DIR = Path(r\"{folders[0]}\")")
        if len(folders) > 1:
            print("   (found exports in more than one folder — double-check which is the real one)")
    else:
        print("  (none found — check at least one export is actually synced to this machine)")


if __name__ == "__main__":
    main()
