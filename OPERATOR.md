# Weekly pipeline — operator guide

Use this guide each week. You do **not** need to run Python commands or edit spreadsheet formulas.

For technical setup and troubleshooting, see [README.md](README.md).

---

## Before you start

- [ ] **Microsoft Excel is closed** (no OTS ATP, Combine, All Fiber, or OMSP files open).
- [ ] You know the **current week** (e.g. `Week 20`) and **previous week** (e.g. `Week 19`).
- [ ] The week folder exists under **ATP NMS** (e.g. `ATP NMS\Week 20\`).

---

## Step 1 — Download exports from CPQ / NMS

Export the usual CPQ and NMS files for this week (ZIPs and Excel files).

---

## Step 2 — Rename files and copy to `raw files`

Open this folder (replace `Week 20` with your week):

```text
ATP NMS\Week 20\input\raw files\
```

**Rename every download** to exactly these names, then copy them into `raw files`:

| Save as this name | What it is |
|-------------------|------------|
| `cpq.zip` | CPQ outer export ZIP |
| `performance.zip` | Current Performance Data ZIP |
| `card.zip` | Card Report ZIP |
| `fiber_table.zip` | Fiber Table Data ZIP |
| `oam.xlsx` | Optical Attenuation Management workbook |
| `sfp.xlsx` | SFP Information Report workbook |

All six files must be present. Spelling and extension must match exactly (lowercase names as shown).

**Example:** if you downloaded `Card_2026-05-05_09-55-38.zip`, rename it to `card.zip` before copying.

---

## Step 3 — Update week labels (once per week)

Ask your technical contact to update **`config\ingest.yml`** in the project folder, or edit only these two lines if you are allowed to:

- `week_label: "Week 20"` → current week  
- `previous_week_label: "Week 19"` → last completed week  

Do not change other lines unless instructed.

---

## Step 4 — Run the pipeline

1. Go to the project folder: `final-project-huawei-fadia`
2. Double-click **`run.bat`**
3. Wait until the window shows **Pipeline finished OK**
4. If it says **FAILED**, take a screenshot of the window and send it to support. Do not edit output Excel files by hand to “fix” them.

The run can take several minutes. Do not close the window early.

---

## Step 5 — Check outputs

In your week folder (e.g. `ATP NMS\Week 20\`), open and spot-check:

| File | Purpose |
|------|---------|
| `OTS ATP NMS W20.xlsx` | Master report |
| `Combine_….xlsx` | Combine workbook |
| `Week 20_All Fiber.xlsx` | All Fiber |
| `Week 20_OMSP_DWDMEvaluation.xlsx` | OMSP / DWDM evaluation |

Folders `input\CPQ`, `input\NMS`, and `computed` are created automatically — you normally do not need to open them.

---

## Quick checklist (printable)

```text
[ ] Excel closed
[ ] Six files in raw files\ with correct names (cpq.zip … sfp.xlsx)
[ ] week_label / previous_week_label updated in config\ingest.yml
[ ] Double-clicked run.bat
[ ] Saw "Pipeline finished OK"
[ ] Checked master, Combine, All Fiber, OMSP
```

---

## Common problems

| Problem | What to do |
|---------|------------|
| “Config not found” | Copy `config\ingest.example.yml` to `config\ingest.yml` and set week labels. |
| “cpq_zip not found” (or similar) | A file is missing or not renamed — check all six names in `raw files`. |
| “Cannot save … close Excel” | Close every Excel window and run `run.bat` again. |
| Pipeline FAILED | Screenshot the black window; do not rerun after manually editing outputs. |

---

## What you should never do

- Do not rename files inside `input\CPQ` or `input\NMS` after the pipeline runs.
- Do not delete `computed` while a run is in progress.
- Do not change column layouts in the master template unless your team tells you to.

---

## Need help?

Contact your technical owner with:

1. Week number  
2. Screenshot of the `run.bat` window  
3. Confirmation that all six raw files were renamed correctly  
