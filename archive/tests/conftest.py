"""
Tests of the archived scripts — not run by CI (pytest's testpaths only
covers tests/). Run them by hand with:

    uv run pytest archive/tests

archive/scripts/ comes first on sys.path (its config.py is the full
pre-archive version these scripts expect); scripts/ comes second, for
generate_data_entry_file, which integrate_data_entry imports.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "archive" / "scripts"))
