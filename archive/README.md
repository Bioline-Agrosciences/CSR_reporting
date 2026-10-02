# Archive

Scripts that were needed to build and migrate the pipeline, but are not
used to run the monthly reporting anymore. Kept for reference and in case a
migration ever has to be replayed. The reasons behind each step are in
[DECISIONS.md](DECISIONS.md).

| Script | What it did |
|---|---|
| `extract_reference_and_data.py` | One-time conversion of the original 12-tab workbooks into `Indicator_reference_CSR.xlsx` + `Raw_data_CSR.xlsx` |
| `migrate_2026_09_*.py`, `migrate_2026_10_*.py` | One-time fixes of the real files after each decision (Year column, Env.1 removal, reclassified factors, carbon indicators, `CSR_parameters.xlsx` creation) — all already run |
| `integrate_data_entry.py`, `integrate_safety_data.py`, `csr_calc_engine.py`, `consolidation_report.py`, `csr_parameters.py` | Former local consolidation (monthly cycle steps 3-5, with the Outlook anomaly email), replaced by the Fabric notebook `nb_consolidate_csr_data` on 02/10/2026 |
| `compare_archived_vs_consolidated.py` | Diagnostic: original workbooks vs consolidated results |
| `locate_safety_files.py` | Finds the Working Hours file and the BlueKanGo export in OneDrive |
| `config.py` | The full `config.py` these scripts expect (the one in `scripts/` only keeps what is still used) |

These scripts are frozen: they are not maintained and not run by CI. Their
tests still pass and can be run by hand:

```
uv run pytest archive/tests
```

To run one of them exactly as it was, check out the last commit before the
archive (`git checkout 202ffeb`) and run it from `scripts/`.
