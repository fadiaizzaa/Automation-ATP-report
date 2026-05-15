"""
Experimental canonical Service Routing Report builder.

This implements the safer flow:
  raw SRR -> validate/find configured headers -> clean internal layout -> formulas

If Huawei inserts or moves a raw column, this script finds the configured header
from ingest.yml and writes it into the stable internal column that formulas use.

It intentionally writes to a separate computed workbook:
  {week_label}_ServiceRouting_canonical.xlsx

Run manually:
  python srr.py --config ingest.yml
"""

from __future__ import annotations

import argparse
import logging
import shutil
import sys
from pathlib import Path
from typing import Any

import try_srr2 as base
from openpyxl.utils import column_index_from_string
from pipeline_config import (
    WB_SERVICE_ROUTING,
    build_pipeline_paths,
    get_sheet_spec,
    idx_map_from_headers,
    load_yaml,
)


log = logging.getLogger("srr_header_safe")


OCH_KEYS_TO_FORMULA_COLUMNS = {
    "och_c": "C",
    "och_e": "E",
    "och_g": "G",
    "och_h": "H",
    "och_m": "M",
    "och_o": "O",
    "och_p": "P",
    "och_q": "Q",
    "och_s": "S",
    "och_ac": "AC",
    "och_ae": "AE",
    "och_af": "AF",
    "och_ag": "AG",
    "och_ai": "AI",
}


E2E_KEYS_TO_FORMULA_COLUMNS = {
    "e2e_c": "C",
    "e2e_d": "D",
    "e2e_e": "E",
    "e2e_h": "H",
    "e2e_j": "J",
    "e2e_m": "M",
    "e2e_o": "O",
    "e2e_s": "S",
    "e2e_u": "U",
    "e2e_x": "X",
    "e2e_z": "Z",
    "e2e_ad": "AD",
    "e2e_ae": "AE",
    "e2e_ag": "AG",
    "e2e_aj": "AJ",
    "e2e_al": "AL",
    "e2e_ao": "AO",
    "e2e_as": "AS",
    "e2e_av": "AV",
    "e2e_ax": "AX",
    "e2e_ba": "BA",
}


RAW_LAST_BY_SHEET = {
    base.OCH_SHEET: base.OCH_RAW_LAST,
    base.E2E_SHEET: base.E2E_RAW_LAST,
}


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def _as_row(values: Any) -> list[Any]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [values]
    if values and isinstance(values[0], list):
        return values[0]
    return values


def _as_column(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and isinstance(values[0], list):
        return values
    return [[v] for v in values]


def _last_row_from_required_columns(raw_ws, source_columns_1based: list[int]) -> int:
    candidates = [1] + source_columns_1based
    last_row = 1
    for col in candidates:
        try:
            row = raw_ws.range((raw_ws.cells.last_cell.row, col)).end("up").row
            last_row = max(last_row, row)
        except Exception:
            pass
    return last_row


def write_canonical_raw_sheet(raw_ws, out_ws, sheet_name: str, columns_yaml: dict[str, str]) -> int:
    key_to_dest = (
        OCH_KEYS_TO_FORMULA_COLUMNS
        if sheet_name == base.OCH_SHEET
        else E2E_KEYS_TO_FORMULA_COLUMNS
    )

    raw_header_row = 4
    raw_last_col = max(raw_ws.used_range.last_cell.column, 1)
    headers = _as_row(raw_ws.range((raw_header_row, 1), (raw_header_row, raw_last_col)).value)
    idx_by_key = idx_map_from_headers(headers, columns_yaml, list(key_to_dest))
    source_columns = [idx + 1 for idx in idx_by_key.values()]
    raw_last_row = _last_row_from_required_columns(raw_ws, source_columns)

    out_last_row = max(1, raw_last_row - raw_header_row + 1)
    raw_last_letter = RAW_LAST_BY_SHEET[sheet_name]
    raw_last_internal_col = column_index_from_string(raw_last_letter)
    raw_copy_last_col = min(raw_last_col, raw_last_internal_col)
    raw_values = raw_ws.range((raw_header_row, 1), (raw_last_row, raw_copy_last_col)).value
    out_ws.range((1, 1), (out_last_row, raw_copy_last_col)).value = raw_values

    for key, dest_col in key_to_dest.items():
        src_col = idx_by_key[key] + 1
        values = _as_column(raw_ws.range((raw_header_row, src_col), (raw_last_row, src_col)).value)
        out_ws.range(f"{dest_col}1:{dest_col}{out_last_row}").value = values

    return out_last_row


def build_output(src, out, master_path, combine_path, workbooks: dict[str, Any]) -> None:
    src = Path(src)
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    if out.exists():
        out.unlink()
    log.info("Creating canonical computed workbook -> %s", out.name)

    lib = base.lib_ref(master_path)
    cmb = base.combine_ref(combine_path)

    app = base.excel_app()
    try:
        raw_wb = app.books.open(str(src), update_links=False)
        wb = app.books.add()
        wb.sheets[0].name = base.OCH_SHEET
        wb.sheets.add(base.E2E_SHEET, after=wb.sheets[0])
        for sh in list(wb.sheets):
            if sh.name not in (base.OCH_SHEET, base.E2E_SHEET):
                sh.delete()

        steps = [
            (
                base.OCH_SHEET,
                base.OCH_RAW_LAST,
                "BE",
                base.OCH_FML_LAST,
                base.och_formulas,
            ),
            (
                base.E2E_SHEET,
                base.E2E_RAW_LAST,
                "BL",
                base.E2E_FML_LAST,
                base.e2e_formulas,
            ),
        ]

        for sheet_name, raw_last, fml_first, fml_last, fml_fn in steps:
            raw_ws = raw_wb.sheets[sheet_name]
            ws = wb.sheets[sheet_name]
            spec = get_sheet_spec(workbooks, WB_SERVICE_ROUTING, sheet_name)

            log.info("Processing %s ...", sheet_name)
            if spec.header_row_1based != 4:
                raise ValueError(
                    f"{sheet_name}: expected raw header row 4 before metadata deletion, "
                    f"got {spec.header_row_1based}"
                )

            n_last = write_canonical_raw_sheet(raw_ws, ws, sheet_name, spec.columns)
            log.info("  Data rows: %s to %s", base.FIRST_DATA_ROW, n_last)
            log.info("  Canonical raw cols written through %s", raw_last)

            ws.range(f"{fml_first}{base.FIRST_DATA_ROW}:{fml_last}{n_last}").clear_contents()
            formulas = fml_fn(base.FIRST_DATA_ROW, lib, cmb)
            for col, formula in formulas.items():
                ws.range(f"{col}{base.FIRST_DATA_ROW}").formula = formula
                if n_last > base.FIRST_DATA_ROW:
                    ws.range(f"{col}{base.FIRST_DATA_ROW}:{col}{n_last}").api.FillDown()
            log.info("  Formula cols %s:%s written and filled down", fml_first, fml_last)

            base.apply_formula_headers_and_header_style(ws, sheet_name)
            log.info("  Formula headers and header-only styling applied")

        wb.save(str(out))
        wb.close()
        raw_wb.close()
        log.info("Saved -> %s", out.name)
    finally:
        app.screen_updating = True
        app.quit()


def run(config_path: Path) -> int:
    cfg = load_yaml(config_path)
    paths = build_pipeline_paths(cfg, config_path)

    workbooks = cfg.get("workbooks")
    if not isinstance(workbooks, dict):
        log.error("Config must define workbooks:")
        return 1

    src = paths.service_routing_raw
    out = paths.computed_dir / f"{paths.week_label}_ServiceRouting_canonical.xlsx"
    master = paths.master_workbook
    combine = paths.combine_workbook

    for p, name in [(src, "Service Routing"), (master, "Master"), (combine, "Combine")]:
        if not p.exists():
            log.error("%s file not found: %s", name, p)
            return 1

    pipe = cfg.get("pipeline") if isinstance(cfg.get("pipeline"), dict) else {}
    make_master_backup = pipe.get("make_master_backup", True)
    full_excel_rebuild = base.config_bool(pipe.get("full_excel_rebuild"), False)
    if make_master_backup:
        backup = master.with_stem(master.stem + "_SRR_CANONICAL_BACKUP")
        shutil.copy2(master, backup)
        log.info("Master backup -> %s", backup.name)

    log.info("Source  : %s", src.name)
    log.info("Output  : %s", out.name)
    log.info("Master  : %s", master.name)
    log.info("Combine : %s", combine.name)

    build_output(src, out, master, combine, workbooks)
    base.update_lib(master, out)
    base.recalculate(master, out, combine, full_rebuild=full_excel_rebuild)
    base.paste_to_combine(out, combine)
    base.paste_to_cpq(master, out)

    log.info("Done.")
    log.info("  SRR output : %s", out.resolve())
    log.info("  Combine    : %s", combine.resolve())
    log.info("  Master     : %s", master.resolve())
    return 0


def main(argv: list[str] | None = None) -> int:
    setup_logging()
    parser = argparse.ArgumentParser(description="Header-safe Service Routing Report pipeline step")
    parser.add_argument("--config", type=Path, required=True)
    args = parser.parse_args(argv)
    cfg_path = args.config.resolve()
    if not cfg_path.is_file():
        log.error("Config not found: %s", cfg_path)
        return 1
    try:
        return run(cfg_path)
    except Exception:
        log.exception("srr failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
