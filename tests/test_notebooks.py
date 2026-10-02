"""The Fabric notebooks in scripts/ must never carry outputs: they contain
real data. Import them with `import-fabric-notebook` (bioline_utils)."""
import json
from pathlib import Path

import pytest

NOTEBOOKS = sorted((Path(__file__).resolve().parents[1] / "scripts").glob("*.ipynb"))


def test_notebooks_are_found():
    assert {p.name for p in NOTEBOOKS} >= {"nb_consolidate_csr_data.ipynb", "nb_consolidate_sap_and_csr_data.ipynb"}


@pytest.mark.parametrize("path", NOTEBOOKS, ids=lambda p: p.name)
def test_notebook_has_no_outputs(path):
    nb = json.loads(path.read_text(encoding="utf-8"))
    for cell in nb["cells"]:
        if cell["cell_type"] == "code":
            assert cell["outputs"] == [], f"{path.name}: a cell still has outputs"
    assert not nb["metadata"].get("synapse_widget", {}).get("state"), f"{path.name}: widget state not empty"
