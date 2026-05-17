# Week data pipeline

Python tools that unpack CPQ/NMS exports, rebuild computed workbooks, and paste results into the master **OTS ATP NMS** Excel file and related **Combine**, **All Fiber**, and **OMSP** workbooks.

- **Operators:** read [OPERATOR.md](OPERATOR.md) (weekly checklist, fixed raw file names, `run.bat`).
- **Configuration:** [`config/ingest.yml`](config/ingest.yml) (week labels, data paths, sheet column maps).

## Project layout

```text
final-project-huawei-fadia/
  run.bat                    ← operators: double-click to run
  OPERATOR.md                  ← non-technical weekly guide
  README.md                    ← this file (technical)
  config/
    ingest.yml
    ingest.example.yml
  pipeline/
    run/                       ← orchestration + pipeline_config.py
    cpq/                       ← Fiber, OMS, OCh, SRR, All Fiber, …
    pasting_nms.py
    build_current_performance_output.py
    cur_performance.py
  requirements.txt
```

Week data lives under **`path`** in `ingest.yml` (e.g. `ATP NMS\Week 19\`), not inside this repo.

## Raw input files (fixed names)

Operators must place exports in `{output_base}/input/raw files/` using **exact** names (rename downloads every week; do not edit `ingest.yml` for filenames):

| File in `raw files/` | Config key |
|----------------------|------------|
| `cpq.zip` | `inputs.cpq_zip` |
| `performance.zip` | `inputs.current_performance_file` |
| `card.zip` | `inputs.card_file` |
| `fiber_table.zip` | `inputs.fiber_table_file` |
| `oam.xlsx` | `inputs.oam_file` |
| `sfp.xlsx` | `inputs.sfp_file` |

## Prerequisites

- **Python 3.10+**
- **Microsoft Excel** on **Windows** (xlwings)
- One-time setup:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Run full pipeline

```bash
# Operators
run.bat

# Command line
.venv\Scripts\activate
python pipeline\run\run_week_pipeline.py --config config\ingest.yml
```

Steps: seed artifacts → prepare inputs → validate → Fiber → OMS → OCh → paste CPQ → SRR → update All Fiber → paste NMS → current performance.

### Individual steps

```bat
set PYTHONPATH=%CD%\pipeline\run;%CD%\pipeline\cpq;%CD%\pipeline
python pipeline\cpq\nrr_fiber.py --config config\ingest.yml
```

Many scripts default to `config\ingest.yml` when `--config` is omitted.

### Daily Current Performance

Drop intraweek ZIP(s) in `{output_base}/input/raw files/performance_daily/`, then:

```bash
cur_performance.bat
# or
python pipeline\cur_performance.py --config config\ingest.yml
```

The full weekly pipeline already runs current performance from the prepared weekly NMS files. This manual helper is only for intraweek refreshes: it extracts ZIPs, syncs NMS performance workbooks, overwrites master **OCH Performance** / **OAU**, refreshes master **BU** formulas to the current `dwdm_omsp_output`, and updates OMSP **OSC**. Processed ZIPs move to `performance_daily/processed/`. Use `--no-daily-zip` to use only the weekly prepared NMS files.

### Other manual helpers

```bash
python pipeline\cpq\lib_updater.py --config config\ingest.yml
```

## Week data folder

```text
ATP NMS/Week 19/
  input/raw files/       ← cpq.zip, performance.zip, card.zip, fiber_table.zip, oam.xlsx, sfp.xlsx
  input/CPQ/             ← prepared automatically
  input/NMS/             ← prepared automatically
  computed/
  OTS ATP NMS W19.xlsx
  Combine_....xlsx
  Week 19_All Fiber.xlsx
  Week 19_OMSP_DWDMEvaluation.xlsx
```

## Configure `ingest.yml`

| Key | Purpose |
|-----|--------|
| `path` | Root folder for all weeks (e.g. `ATP NMS`) |
| `week_label` / `previous_week_label` | Current and prior week folder names |
| `output_base` | Usually `{path}/{week_label}` |
| `inputs.*` | Raw files under `input/raw files/` (fixed names above) |
| `pipeline.*` | Master, combine, computed dir, optional overrides |
| `workbooks:` | Column/header mappings (technical; operators do not edit) |

Copy [`config/ingest.example.yml`](config/ingest.example.yml) to `config/ingest.yml` for new installs.

## Troubleshooting

- **Raw file not found:** check all six names in `input/raw files/` ([OPERATOR.md](OPERATOR.md)).
- **Config not found:** `config/ingest.yml` missing — copy from `ingest.example.yml`.
- **Import errors:** use `run.bat` or set `PYTHONPATH` like the batch file.
- **Excel / COM errors:** close workbooks; re-run; `update_all_fiber` / `pasting_nms` exit 1 if Excel recalc fails.
- **Stale OMSP links:** `update_all_fiber` runs before `pasting_nms`; `cur_performance` runs last in `run_week_pipeline.py`.

## License / project

Internal / coursework use per your team policy.
