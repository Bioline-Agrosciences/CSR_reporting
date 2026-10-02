# Monthly CSR Reporting — Bioline Agrosciences

Collects and consolidates the monthly CSR indicators (energy, water, waste,
refrigerants, safety...) of Bioline's 6 business units: Viridaxis, BAF, BFR,
BIB, BUK, BUS. Results land in Fabric Delta tables read by Power BI.

## How it works

```
 LOCAL (this repo)                SHAREPOINT "CSR referents" / General/          FABRIC (LH_CSR_Reporting)
                                  (OneLake shortcut Files/sp_csr_general)

 generate_data_entry_file.py ──►  Monthly reporting/
                                    <BU>_data_entry_CSR_<year>.xlsm
                                    (each referent fills in their file)  ──┐
                                  Working Hours <year>.xlsx              ──┤
                                  Monthly reporting/extract_bluekango/   ──┤   nb_consolidate_csr_data
                                    (BlueKanGo accidents export)           ├─►   CSR_raw_data
                                  Reporting_Automation/monthly_reporting/  │     CSR_indicators_report
                                    input_data/                            │     CSR_completion_report
                                      Indicator_reference_CSR.xlsx       ──┤     CSR_parameters
                                      CSR_parameters.xlsx                ──┘     CSR_anomalies
                                                                                      │
                                          commercial semantic model ──► nb_consolidate_data_for_CSR_report
                                          ("Sales | Commercial")          CSR_gold_reporting (+ sales)
                                                                          CSR_gold_tracking
                                                                                      │
                                                                                   Power BI (ratios)
```

1. **Data entry files** — `scripts/generate_data_entry_file.py` (run
   locally) writes one `.xlsm` per BU and per year into the SharePoint
   folder the referents already use. Each file only lists the indicators
   that referent owns, one tab per month; only the value and comment cells
   are editable. A macro keeps the current month highlighted and the completion
   tracker up to date every time the file is opened, so it only needs
   regenerating when the indicator list changes or a new year starts.
2. **Consolidation** — the Fabric notebook `nb_consolidate_csr_data`
   reads the data entry files, the Working Hours file and the latest
   BlueKanGo export straight from SharePoint, cleans the values, computes
   the additive indicators and writes the Delta tables above.
3. **Dashboard tables** — the Fabric notebook
   `nb_consolidate_data_for_CSR_report` reads the "Sales | Commercial"
   measure of the commercial dashboard's semantic model (actuals, per SAP
   entity and month, so the figure matches the commercial report), maps SAP
   entities to BUs, and writes:
   - `CSR_gold_reporting`: the CSR indicators plus the sales, one row per
     SAP entity and month (`Env.1` in k€, `Sales_€` in €). Columns: Entity
     (empty for CSR indicators), BU, Year, Month (number), MonthName, ID,
     Value, Date (1st of the month, for d_Calendar);
   - `CSR_gold_tracking`: `CSR_completion_report` with the same date
     columns.
4. **Ratios** — computed on the fly in the Power BI dashboard.

The two notebooks are versioned in `scripts/` as `.ipynb` files exported
from Fabric **without their outputs** (outputs contain real data). Fabric is
not connected to this repo: after a change in Fabric, export the notebook
and import it with `import-fabric-notebook` (see Setup); after a change
here, import the `.ipynb` back into Fabric.

## Key rules

- **Sums only in Fabric, ratios in the dashboard.** A sum rolls up
  correctly to a group total, a ratio does not (averaging 6 BUs' "%
  renewable energy" gives a meaningless number). Ratios (Wat.2, Ene.10,
  Ene.11, Was.3, Was.4, Saf.6, Saf.7) are computed in Power BI from the
  sums, at the level of aggregation displayed.
- **Every row is keyed by (BU, Year, Month, ID).** Data entry files are one
  per year; a new year never overwrites the previous one.
- **Values are cleaned, never guessed.** A decimal comma or numeric text is
  converted and the correction is logged; unreadable text is left empty and
  flagged for a fix at the source.
- **No real data in git.** All data lives on SharePoint; `input_data/` and
  `output_data/` stay in `.gitignore` as a safety net.

## The indicator reference list

`input_data/Indicator_reference_CSR.xlsx` (sheet "Reference") is the single
source of truth: one row per indicator (46 today), with its unit,
definition and:

- **Kind**: `input` (entered by someone) or `calculated` (never entered).
- **Responsible**: who owns the number. Only `input` + `CSR referent`
  indicators appear in the data entry files. Safety indicators (Saf.*,
  owned by H&S / HR) come from Working Hours and BlueKanGo instead.

Adding or removing a row is picked up by the next data entry file
generation and the next notebook run. A new **calculated** indicator also
needs its formula: in `FORMULAS` (consolidation notebook) if it is a sum,
or as a measure in the Power BI dashboard if it is a ratio.

Calculated indicators are of three kinds:

| Kind | Examples | Where |
|---|---|---|
| Sums of entered values | Ene.9 (total energy), Ref.1, Carb.1-5 (CO2 = energy x factor) | `FORMULAS`, consolidation notebook |
| Context values (depend only on BU and period) | Ene.6.1, Ene.7.1, Ene.12-15 (factors), Saf.4.1 (working days, public holidays of the BU's country) | `CONTEXTUAL_VALUES`, consolidation notebook |
| Ratios | Wat.2, Ene.10, Ene.11, Was.3, Was.4, Saf.6, Saf.7 | Power BI dashboard |

Env.1 (sales) is not in the reference list: it comes from SAP, through the
`nb_consolidate_data_for_CSR_report`.

## Emission and conversion factors

`input_data/CSR_parameters.xlsx`, sheet "Parameters": one row per
(Parameter, BU, Year) with Value, Unit and Source.

- **To publish a new year's factor, add rows with the new Year — never edit
  the previous year's rows**, otherwise already-published emissions change.
- A year without its own row reuses the most recent earlier year's value,
  and the notebook flags it in `CSR_anomalies`. A factor missing altogether
  leaves the dependent indicators empty (and flagged), never computed with 0.

## Safety indicators

| ID | Source | Rule |
|---|---|---|
| Saf.4.2 hours worked | Working Hours `<year>`.xlsx | blank cell = month not declared yet |
| Saf.2 / Saf.3 injuries without / with lost time | BlueKanGo export | agency workers excluded; a month in Working Hours with no accident = 0 |
| Saf.5 days lost | BlueKanGo export | sum of stoppage days |
| Saf.1 days without accident | BlueKanGo export | days from month end to the BU's last accident, across years |

## Anomalies

Each notebook run appends to `CSR_anomalies` (with `Run_date`): corrected or
unreadable values, missing values for every month up to the current one,
month-over-month variations above 20%, and carried-over or missing factors.
The alert email is to be sent from the Fabric Data Pipeline (Office 365
Outlook activity) or Activator, based on this table.

## Setup (local part)

Requirements: [uv](https://docs.astral.sh/uv/), and the "CSR referents"
SharePoint library synced on the computer ("Sync" in SharePoint). Clone the
repo **outside OneDrive** (OneDrive locks files and breaks `uv sync`):

```
git clone https://github.com/Bioline-Agrosciences/CSR_reporting.git C:\Users\<me>\Documents\CSR\monthly_reporting
cd C:\Users\<me>\Documents\CSR\monthly_reporting
uv sync
```

Paths are in `scripts/config.py`. If the library is synced somewhere else
than `C:\Users\<me>\Bioline Agrosciences Group\CSR referents - Documents`,
set the `CSR_SHAREPOINT_GENERAL_DIR` environment variable to its `General`
folder.

Generate the data entry files:

```
uv run scripts/generate_data_entry_file.py          # all 6 BUs
uv run scripts/generate_data_entry_file.py BAF BFR   # a selection
```

## Setup (Fabric part)

- Lakehouse `LH_CSR_Reporting` (workspace BM_F_D - SAP-B1) attached as
  default lakehouse to both notebooks, with the OneLake shortcut
  `Files/sp_csr_general` pointing to `Shared Documents/General` of the
  CSRreferents SharePoint site.
- `holidays==0.105` installed for `nb_consolidate_csr_data` (`%pip` cell,
  or a Fabric Environment when run from a Data Pipeline). Keep it the same
  version as `uv.lock`: holiday dates change between versions, and with
  them Saf.4.1.
- `nb_consolidate_data_for_CSR_report` reads the semantic model
  "Bioline Agrosciences sales - Copy" of the same workspace (`sempy`).

To version a notebook changed in Fabric: export it (File > Export >
.ipynb), then

```
import-fabric-notebook "C:\Users\<me>\Downloads\nb_consolidate_csr_data.ipynb" scripts/
git diff
```

`import-fabric-notebook` comes from the
[bioline_utils](https://github.com/Bioline-Agrosciences/bioline_utils)
package; install it once with
`uv tool install -e C:\Users\<me>\Documents\bioline_utils` (from a local
clone). A test fails if a notebook in `scripts/` still has outputs.

## Layout

| Path | Content |
|---|---|
| `scripts/` | `generate_data_entry_file.py`, `config.py`, and the two Fabric notebooks (`nb_*.ipynb`, without outputs) |
| `templates/` | Empty `.xlsm` template carrying the data entry file's macro |
| `tests/` | pytest tests, run by CI on every push/PR |
| `archive/` | Migration scripts and the former local consolidation, plus the history of decisions (`archive/DECISIONS.md`) |

## Tests

```
uv run pytest
```
