"""
Lib Auto-Updater  (xlwings version)
=====================================
Scans OMS_T col L + col R for NE names missing from Lib col N/O/P,
then appends the missing rows. Uses xlwings so slicers, pivots, and
conditional formatting in the master workbook are fully preserved.

Run after ``pasting_cpq2`` / ``prepare_week_inputs`` has pasted fresh OMS Trails into OMS_T.

NE Name format: "{NE_ID}_{prefix}_{SiteName}"
  e.g. "6760_DBALo_Amplapura" -> NE ID=6760, Site="Amplapura"

After this script finishes, open the master in Excel and press Ctrl+Alt+F9
to force-recalculate OMS_T col B (OMS_Name).

CONFIG / USAGE
--------------
All paths come from ingest.yml (same ``pipeline.master_workbook`` resolution as
``pasting_cpq2``). Run::

    python lib_updater.py --config ingest.yml

REQUIREMENTS
------------
    pip install xlwings pyyaml
    Excel must be installed on this machine.
"""

from __future__ import annotations

import argparse
import logging
import shutil
import sys
from pathlib import Path

import xlwings as xw

from pipeline_config import build_pipeline_paths, load_yaml

log = logging.getLogger("lib_updater")


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def parse_ne_name(ne_name: str):
    """
    Parse "{NE_ID}_{prefix}_{SiteName}" into (int ne_id, str site).
    Returns (None, None) if unparseable.
      "6760_DBALo_Amplapura"   -> (6760, "Amplapura")
      "6725_MJ_BumiSetiaMekar" -> (6725, "BumiSetiaMekar")
    """
    if not ne_name or not isinstance(ne_name, str):
        return None, None
    parts = ne_name.split("_", 2)
    if len(parts) < 3:
        return None, None
    try:
        ne_id = int(parts[0])
    except ValueError:
        return None, None
    return ne_id, parts[2]


def to_flat(vals):
    """Flatten xlwings range value into a plain list."""
    if vals is None:
        return []
    if isinstance(vals, list):
        return [v[0] if isinstance(v, list) else v for v in vals]
    return [vals]

def run(config_path: Path) -> int:
    cfg = load_yaml(config_path)
    paths = build_pipeline_paths(cfg, config_path)
    master_path = paths.master_workbook
    if not master_path.is_file():
        log.error("Master file not found: %s", master_path)
        return 1

    log.info("Master workbook path: %s", master_path)

    pipe = cfg.get("pipeline") if isinstance(cfg.get("pipeline"), dict) else {}
    make_backup = pipe.get("make_master_backup", True)
    if make_backup:
        backup = master_path.with_stem(master_path.stem + "_BACKUP")
        shutil.copy2(master_path, backup)
        log.info("Backup → %s", backup.name)

    log.info(f"Opening: {master_path.name}")
    app = xw.App(visible=False, add_book=False)
    app.display_alerts  = False
    app.screen_updating = False

    try:
        wb      = app.books.open(str(master_path))
        ws_omst = wb.sheets["OMS_T"]
        ws_lib  = wb.sheets["Lib"]

        # ── Step 1: Collect NE names from OMS_T col L + col R ────
        log.info("Reading OMS_T col L + col R ...")
        last_row = ws_omst.range("L4").end("down").row
        log.info(f"  OMS_T data: rows 4 to {last_row}")

        src = to_flat(ws_omst.range(f"L4:L{last_row}").value)
        snk = to_flat(ws_omst.range(f"R4:R{last_row}").value)

        all_ne_names = set()
        for name in src + snk:
            if name and isinstance(name, str) and "_" in name:
                all_ne_names.add(name.strip())
        log.info(f"  Unique NE names found: {len(all_ne_names)}")

        # ── Step 2: Parse NE names ────────────────────────────────
        parsed = {}   # ne_id (int) -> (ne_name, site)
        unparseable = []
        for name in all_ne_names:
            ne_id, site = parse_ne_name(name)
            if ne_id is not None:
                if ne_id not in parsed:
                    parsed[ne_id] = (name, site)
            else:
                unparseable.append(name)

        if unparseable:
            log.warning(f"  Could not parse {len(unparseable)} name(s): {unparseable[:5]}")

        # ── Step 3: Read existing Lib col N/O ────────────────────
        log.info("Reading existing Lib col N/O/P ...")
        last_nop_row = ws_lib.range("N1").end("down").row
        nop_data     = ws_lib.range(f"N2:O{last_nop_row}").value

        existing_nop = set()
        if nop_data:
            for row in nop_data:
                ne_name_val, ne_id_val = (row if isinstance(row, list) else [row, None])
                if ne_id_val is not None and str(ne_name_val) != "NEName":
                    try:
                        existing_nop.add(int(ne_id_val))
                    except (ValueError, TypeError):
                        pass

        log.info(f"  Lib col N/O/P: {len(existing_nop)} entries (last row {last_nop_row})")

        # ── Step 4: Find missing ──────────────────────────────────
        missing_nop = {ne_id: v for ne_id, v in parsed.items()
                       if ne_id not in existing_nop}
        log.info(f"  Missing from col N/O/P: {len(missing_nop)}")

        if not missing_nop:
            log.info("  Lib col N/O/P is already up to date — nothing to add.")
            app.calculate()
            wb.save()
            wb.close()
            return 0

        # ── Step 5: Append missing rows to col N/O/P ─────────────
        log.info(f"Appending {len(missing_nop)} new NE(s) to Lib col N/O/P ...")
        for i, (ne_id, (ne_name, site)) in enumerate(sorted(missing_nop.items())):
            target_row = last_nop_row + 1 + i
            ws_lib.range(f"N{target_row}").value = ne_name
            ws_lib.range(f"O{target_row}").value = ne_id
            ws_lib.range(f"P{target_row}").value = site
            log.info(f"  + row {target_row}: {ne_name} | {ne_id} | {site}")

        log.info("Saving …")
        app.calculate()
        wb.save()
        wb.close()

        log.info("Done — %s new NE(s) added to Lib col N/O/P.", len(missing_nop))
        log.info("Saved → %s", master_path.resolve())
        return 0

    except Exception as e:
        log.error("Error: %s", e)
        raise
    finally:
        app.quit()


def main(argv: list[str] | None = None) -> int:
    setup_logging()
    p = argparse.ArgumentParser(description="Update Lib sheet from OMS_T (xlwings)")
    p.add_argument("--config", type=Path, required=True, help="Path to ingest YAML")
    args = p.parse_args(argv)
    cfg_path = args.config.resolve()
    if not cfg_path.is_file():
        log.error("Config not found: %s", cfg_path)
        return 1
    try:
        return run(cfg_path)
    except Exception:
        return 1


if __name__ == "__main__":
    sys.exit(main())
