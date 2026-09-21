"""
The scripts under scripts/ are not a package (no __init__.py, see
config.py): we import them the same way the scripts import each other, by
adding scripts/ to sys.path before any test collection happens.
"""
import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))
