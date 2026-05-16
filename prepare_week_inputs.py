#!/usr/bin/env python3
"""
Prepare normalized week input workbooks from raw exports (ZIP / XLSX).

Reads a YAML config (paths relative to the config file directory), extracts
nested CPQ / NMS archives, and writes files under:

  {output_base}/input/CPQ/  (or {output_base}/CPQ/ if output_base already ends with input)
  {output_base}/input/NMS/  (or {output_base}/NMS/ if output_base already ends with input)

Requirements:
  pip install pyyaml
  pip install xlwings   (optional; used to full-recalc each output Excel so cached values exist for later scripts)

Usage:
  python prepare_week_inputs.py --config ingest.yml
  python prepare_week_inputs.py --config ingest.yml --dry-run
  python prepare_week_inputs.py --config ingest.yml --skip-excel-recalc
"""

from __future__ import annotations

import argparse
import logging
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any

import yaml

from pipeline_config import configured_output_base, expand_config_value, prepared_input_root

_EXCEL_SUFFIXES = (".xlsx", ".xlsm", ".xls")

log = logging.getLogger("prepare_week_inputs")

# --- CPQ inner zip classification (case-insensitive name contains) ---
CPQ_SRR = "serviceroutingreport"
CPQ_TRAIL = "trail integrity report"
CPQ_NETWORK = "network resource report"

TRAIL_PREFIXES: list[tuple[str, str]] = [
    ("OCh Trail Integrity", "OCh Trail Integrity.xlsx"),
    ("OMS Trail Integrity", "OMS Trail Integrity.xlsx"),
    ("Client Trail Integrity", "Client Trail Integrity.xlsx"),
    ("STM Trail Integrity", "STM Trail Integrity.xlsx"),
]

REQUIRED_INPUT_KEYS = (
    "cpq_zip",
    "current_performance_file",
    "card_file",
    "oam_file",
    "sfp_file",
)
OPTIONAL_INPUT_KEYS = (
    "fiber_table_file",
)


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def load_config(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise FileNotFoundError(f"Config not found: {path}")
    with path.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    if not isinstance(data, dict):
        raise ValueError("Config root must be a mapping")
    return data


def resolve_input_path(config_dir: Path, raw: str) -> Path:
    p = Path(raw)
    if p.is_absolute():
        return p.resolve()
    return (config_dir / p).resolve()


def ensure_exists(p: Path, label: str) -> None:
    if not p.is_file():
        raise FileNotFoundError(f"{label} not found or not a file: {p}")


def safe_extract_zip(zip_path: Path, dest: Path) -> None:
    """Extract zip members to dest, blocking path traversal."""
    dest = dest.resolve()
    with zipfile.ZipFile(zip_path) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            member = info.filename
            target = (dest / member).resolve()
            try:
                target.relative_to(dest)
            except ValueError as e:
                raise RuntimeError(
                    f"Unsafe path in archive {zip_path.name}: {member!r}"
                ) from e
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info, "r") as src, target.open("wb") as out:
                shutil.copyfileobj(src, out)


def iter_inner_zips(root: Path) -> list[Path]:
    return sorted(root.rglob("*.zip"), key=lambda p: str(p).lower())


def classify_cpq_inner_zip(zip_path: Path) -> str | None:
    name = zip_path.name.lower()
    if CPQ_SRR in name.replace(" ", ""):
        return "srr"
    if CPQ_TRAIL in name:
        return "trail"
    if CPQ_NETWORK in name:
        return "network"
    return None


def collect_by_class(zips: list[Path]) -> dict[str, list[Path]]:
    buckets: dict[str, list[Path]] = {"srr": [], "trail": [], "network": []}
    unknown: list[Path] = []
    for z in zips:
        cls = classify_cpq_inner_zip(z)
        if cls is None:
            unknown.append(z)
            continue
        buckets[cls].append(z)
    return buckets


def list_workbooks(root: Path, suffixes: tuple[str, ...]) -> list[Path]:
    out: list[Path] = []
    for ext in suffixes:
        out.extend(p for p in root.rglob(f"*{ext}") if p.is_file())
    return sorted(set(out), key=lambda p: str(p).lower())


def pick_single_xlsx(root: Path, context: str) -> Path:
    xs = [p for p in list_workbooks(root, (".xlsx",)) if not p.name.startswith("~$")]
    if not xs:
        raise RuntimeError(f"{context}: no .xlsx found after extract")
    if len(xs) > 1:
        names = "\n  ".join(str(p.name) for p in xs)
        raise RuntimeError(f"{context}: expected one .xlsx, found {len(xs)}:\n  {names}")
    return xs[0]


def pick_single_xlsm(root: Path, context: str) -> Path:
    ms = [p for p in list_workbooks(root, (".xlsm",)) if not p.name.startswith("~$")]
    if not ms:
        raise RuntimeError(f"{context}: no .xlsm found after extract")
    preferred = [p for p in ms if "network resource statistics" in p.name.lower()]
    pool = preferred if preferred else ms
    if len(pool) > 1:
        names = "\n  ".join(str(p.name) for p in pool)
        raise RuntimeError(f"{context}: ambiguous .xlsm choice:\n  {names}")
    return pool[0]


def map_trail_workbooks(root: Path) -> dict[str, Path]:
    """Map logical trail name (prefix) -> source path."""
    xs = [p for p in list_workbooks(root, (".xlsx",)) if not p.name.startswith("~$")]
    found: dict[str, Path] = {}
    for prefix, _out_name in TRAIL_PREFIXES:
        key = prefix.lower()
        matches = [p for p in xs if p.stem.lower().startswith(key)]
        if not matches:
            raise RuntimeError(
                f"Trail Integrity: missing workbook for prefix {prefix!r}. "
                f"Found: {[p.name for p in xs]}"
            )
        if len(matches) > 1:
            names = ", ".join(p.name for p in matches)
            raise RuntimeError(
                f"Trail Integrity: multiple matches for {prefix!r}: {names}"
            )
        found[prefix] = matches[0]
    return found


def copy_or_dry(src: Path, dst: Path, dry_run: bool) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dry_run:
        log.info("DRY-RUN  copy  %s  ->  %s", src, dst)
        return
    shutil.copy2(src, dst)
    log.info("Wrote %s", dst)


def expected_week_workbooks(week_label: str, cpq_dir: Path, nms_dir: Path) -> list[Path]:
    """Final normalized paths under CPQ / NMS (same names as copy_or_dry targets)."""
    out: list[Path] = [cpq_dir / f"{week_label}_ServiceRoutingReport.xlsx"]
    for _prefix, out_suffix in TRAIL_PREFIXES:
        out.append(cpq_dir / f"{week_label}_{out_suffix}")
    out.append(cpq_dir / f"{week_label}_Network Resource Statistics.xlsm")
    out.extend(
        [
            nms_dir / f"{week_label}_Current Performance Data.xlsx",
            nms_dir / f"{week_label}_Current Performance Data_1.xlsx",
            nms_dir / f"{week_label}_Card Report.xlsx",
            nms_dir / f"{week_label}_Optical_Attenuation.xlsx",
            nms_dir / f"{week_label}_SFP.xlsx",
            nms_dir / f"{week_label}_Fiber Table Data.xlsx",
        ]
    )
    return out


def force_recalculate_workbooks(paths: list[Path]) -> None:
    """
    Open each workbook in Excel, full calculate, save — refreshes cached values
    so openpyxl ``data_only`` reads in downstream steps are not empty.
    """
    targets = [
        p.resolve()
        for p in paths
        if p.suffix.lower() in _EXCEL_SUFFIXES and p.is_file() and not p.name.startswith("~$")
    ]
    if not targets:
        log.info("No Excel outputs to recalculate.")
        return
    try:
        import xlwings as xw
    except ImportError:
        log.warning(
            "xlwings not installed — skipping Excel recalc on prepared files. "
            "Install xlwings + Excel or open each file once and save."
        )
        return

    log.info("Recalculating %d workbook(s) in Excel …", len(targets))
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for p in targets:
            log.info("  Recalc: %s", p.name)
            wb = app.books.open(str(p), update_links=False)
            try:
                app.api.CalculateFullRebuild()
            except Exception:
                app.calculate()
            wb.save()
            wb.close()
        log.info("Excel recalculation finished.")
    except Exception as e:
        log.warning("Excel recalc failed (%s). Downstream steps may see empty formula cells.", e)
    finally:
        app.screen_updating = True
        app.quit()


def process_cpq(
    cpq_zip: Path,
    dest_dir: Path,
    week_label: str,
    dry_run: bool,
) -> None:
    with tempfile.TemporaryDirectory(prefix="cpq_outer_") as tmp:
        tmp_root = Path(tmp)
        safe_extract_zip(cpq_zip, tmp_root)
        inner_zips = iter_inner_zips(tmp_root)
        # Exclude the outer archive if re-found (it shouldn't be inside)
        buckets = collect_by_class(inner_zips)
        errs: list[str] = []
        for key, label in (
            ("srr", "ServiceRoutingReport"),
            ("trail", "Trail Integrity"),
            ("network", "Network Resource"),
        ):
            if len(buckets[key]) != 1:
                lst = buckets[key]
                errs.append(
                    f"{label}: expected exactly 1 inner zip, got {len(lst)} {[p.name for p in lst]}"
                )
        if errs:
            details = "\n".join(f"    {z.name} -> {classify_cpq_inner_zip(z) or 'UNKNOWN'}" for z in inner_zips)
            raise RuntimeError(
                "CPQ layout error:\n  "
                + "\n  ".join(errs)
                + f"\n  Inner zips scanned ({len(inner_zips)}):\n{details}"
            )

        leftover = [z for z in inner_zips if classify_cpq_inner_zip(z) is None]
        if leftover:
            log.warning(
                "CPQ outer: %d zip(s) not classified (ignored): %s",
                len(leftover),
                [z.name for z in leftover[:10]],
            )

        srr_zip, trail_zip, net_zip = (
            buckets["srr"][0],
            buckets["trail"][0],
            buckets["network"][0],
        )

        with tempfile.TemporaryDirectory(prefix="cpq_srr_") as t2:
            p = Path(t2)
            safe_extract_zip(srr_zip, p)
            src = pick_single_xlsx(p, "ServiceRoutingReport")
            copy_or_dry(
                src,
                dest_dir / f"{week_label}_ServiceRoutingReport.xlsx",
                dry_run,
            )

        with tempfile.TemporaryDirectory(prefix="cpq_trail_") as t2:
            p = Path(t2)
            safe_extract_zip(trail_zip, p)
            mapped = map_trail_workbooks(p)
            for prefix, out_suffix in TRAIL_PREFIXES:
                src = mapped[prefix]
                copy_or_dry(src, dest_dir / f"{week_label}_{out_suffix}", dry_run)

        with tempfile.TemporaryDirectory(prefix="cpq_net_") as t2:
            p = Path(t2)
            safe_extract_zip(net_zip, p)
            src = pick_single_xlsm(p, "Network Resource Report")
            copy_or_dry(
                src,
                dest_dir / f"{week_label}_Network Resource Statistics.xlsm",
                dry_run,
            )


def is_performance_variant(stem: str) -> bool:
    s = stem.lower()
    if "@1" in s:
        return True
    if s.endswith("_1"):
        return True
    return False


def process_performance_zip(
    zip_path: Path,
    dest_dir: Path,
    week_label: str,
    dry_run: bool,
) -> None:
    with tempfile.TemporaryDirectory(prefix="perf_") as tmp:
        root = Path(tmp)
        safe_extract_zip(zip_path, root)
        xs = [
            p
            for p in list_workbooks(root, (".xlsx",))
            if not p.name.startswith("~$")
        ]
        if len(xs) != 2:
            names = "\n  ".join(p.name for p in xs) if xs else "(none)"
            raise RuntimeError(
                f"Performance ZIP: expected exactly 2 .xlsx files, got {len(xs)}:\n  {names}"
            )
        xs_sorted = sorted(xs, key=lambda p: (is_performance_variant(p.stem), p.name.lower()))
        copy_or_dry(
            xs_sorted[0],
            dest_dir / f"{week_label}_Current Performance Data.xlsx",
            dry_run,
        )
        copy_or_dry(
            xs_sorted[1],
            dest_dir / f"{week_label}_Current Performance Data_1.xlsx",
            dry_run,
        )


def pick_card_xlsx(root: Path) -> Path:
    xs = [p for p in list_workbooks(root, (".xlsx",)) if not p.name.startswith("~$")]
    if not xs:
        raise RuntimeError("Card ZIP: no .xlsx found")
    if len(xs) == 1:
        return xs[0]
    cardish = [p for p in xs if "card" in p.stem.lower()]
    if len(cardish) == 1:
        return cardish[0]
    names = ", ".join(p.name for p in xs)
    raise RuntimeError(f"Card ZIP: multiple .xlsx files, could not disambiguate: {names}")


def process_card_zip(
    zip_path: Path,
    dest_dir: Path,
    week_label: str,
    dry_run: bool,
) -> None:
    with tempfile.TemporaryDirectory(prefix="card_") as tmp:
        root = Path(tmp)
        safe_extract_zip(zip_path, root)
        src = pick_card_xlsx(root)
        copy_or_dry(src, dest_dir / f"{week_label}_Card Report.xlsx", dry_run)


def process_fiber_table_zip(
    zip_path: Path,
    dest_dir: Path,
    week_label: str,
    dry_run: bool,
) -> None:
    with tempfile.TemporaryDirectory(prefix="fiber_table_") as tmp:
        root = Path(tmp)
        safe_extract_zip(zip_path, root)
        xs = [p for p in list_workbooks(root, (".xlsx",)) if not p.name.startswith("~$")]
        if not xs:
            raise RuntimeError("Fiber Table Data ZIP: no .xlsx found")
        for idx, src in enumerate(xs):
            suffix = "" if idx == 0 else f"_{idx}"
            copy_or_dry(src, dest_dir / f"{week_label}_Fiber Table Data{suffix}.xlsx", dry_run)


def process_flat_xlsx(
    src: Path,
    dest_dir: Path,
    dest_name: str,
    dry_run: bool,
) -> None:
    copy_or_dry(src, dest_dir / dest_name, dry_run)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Normalize raw week exports into input/ layout.")
    p.add_argument("--config", required=True, type=Path, help="Path to ingest YAML")
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Log planned copies only; do not write files",
    )
    p.add_argument(
        "--skip-excel-recalc",
        action="store_true",
        help="Do not open outputs in Excel for full recalculation (faster / no Excel)",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    setup_logging()
    args = parse_args(argv)
    cfg_path = args.config.resolve()
    config_dir = cfg_path.parent

    try:
        data = load_config(cfg_path)
    except (OSError, ValueError, yaml.YAMLError) as e:
        log.error("Failed to load config: %s", e)
        return 1

    week_label = data.get("week_label")
    if not week_label or not isinstance(week_label, str):
        log.error("Config must set non-empty string 'week_label'")
        return 1

    try:
        out_root = configured_output_base(data, config_dir)
    except ValueError as e:
        log.error("%s", e)
        return 1

    inputs_block = data.get("inputs") or data.get("paths")
    if not isinstance(inputs_block, dict):
        log.error("Config must contain 'inputs' mapping (or legacy 'paths')")
        return 1

    missing = [k for k in REQUIRED_INPUT_KEYS if k not in inputs_block]
    if missing:
        log.error("Config 'inputs' missing keys: %s", ", ".join(missing))
        return 1

    paths: dict[str, Path] = {}
    for key in REQUIRED_INPUT_KEYS + OPTIONAL_INPUT_KEYS:
        if key in OPTIONAL_INPUT_KEYS and key not in inputs_block:
            continue
        val = inputs_block[key]
        if not isinstance(val, str) or not val.strip():
            log.error("inputs.%s must be a non-empty string path", key)
            return 1
        try:
            paths[key] = resolve_input_path(
                config_dir,
                expand_config_value(val.strip(), data),
            )
        except ValueError as e:
            log.error("inputs.%s: %s", key, e)
            return 1

    prepared_root = prepared_input_root(out_root)
    cpq_out = prepared_root / "CPQ"
    nms_out = prepared_root / "NMS"

    try:
        ensure_exists(paths["cpq_zip"], "cpq_zip")
        ensure_exists(paths["current_performance_file"], "current_performance_file")
        ensure_exists(paths["card_file"], "card_file")
        ensure_exists(paths["oam_file"], "oam_file")
        ensure_exists(paths["sfp_file"], "sfp_file")
        if "fiber_table_file" in paths:
            ensure_exists(paths["fiber_table_file"], "fiber_table_file")

        log.info("Week: %s", week_label)
        log.info("Output root: %s", prepared_root)
        if args.dry_run:
            log.info("DRY-RUN mode (no files written)")

        process_cpq(paths["cpq_zip"], cpq_out, week_label, args.dry_run)
        process_performance_zip(
            paths["current_performance_file"], nms_out, week_label, args.dry_run
        )
        process_card_zip(paths["card_file"], nms_out, week_label, args.dry_run)
        process_flat_xlsx(
            paths["oam_file"],
            nms_out,
            f"{week_label}_Optical_Attenuation.xlsx",
            args.dry_run,
        )
        process_flat_xlsx(
            paths["sfp_file"],
            nms_out,
            f"{week_label}_SFP.xlsx",
            args.dry_run,
        )
        if "fiber_table_file" in paths:
            process_fiber_table_zip(
                paths["fiber_table_file"],
                nms_out,
                week_label,
                args.dry_run,
            )

        if not args.dry_run and not args.skip_excel_recalc:
            force_recalculate_workbooks(
                expected_week_workbooks(week_label, cpq_out, nms_out)
            )
        elif args.skip_excel_recalc and not args.dry_run:
            log.info("Skipped Excel recalc (--skip-excel-recalc).")
    except (OSError, RuntimeError, ValueError) as e:
        log.error("%s", e)
        return 1

    log.info("Done.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
