#!/usr/bin/env python3
"""Run the full week pipeline steps in order (single command)."""

from __future__ import annotations

import argparse
import logging
import os
import shutil
import subprocess
import sys
from pathlib import Path

from pipeline_config import (
    default_config_path,
    expand_config_value,
    expand_config_value_for_week,
    load_yaml,
    resolve_under,
)

log = logging.getLogger("run_week_pipeline")

RUN_DIR = Path(__file__).resolve().parent
PIPELINE_DIR = RUN_DIR.parent
CPQ_DIR = PIPELINE_DIR / "cpq"
PROJECT_ROOT = PIPELINE_DIR.parent

SEED_FROM_PREVIOUS_KEYS = [
    "master_workbook",
    "combine_workbook",
    "all_fiber_output",
]

STEPS: list[tuple[str, Path]] = [
    ("prepare_week_inputs", RUN_DIR / "prepare_week_inputs.py"),
    ("validate_week_inputs", RUN_DIR / "validate_week_inputs.py"),
    ("nrr_fiber", CPQ_DIR / "nrr_fiber.py"),
    ("oms_trail", CPQ_DIR / "oms_trail.py"),
    ("och_trail", CPQ_DIR / "och_trail.py"),
    ("pasting_cpq2", CPQ_DIR / "pasting_cpq2.py"),
    ("srr", CPQ_DIR / "srr.py"),
    ("update_all_fiber", CPQ_DIR / "update_all_fiber.py"),
    ("pasting_nms", PIPELINE_DIR / "pasting_nms.py"),
    ("cur_performance", PIPELINE_DIR / "cur_performance.py"),
]


def pipeline_pythonpath() -> str:
    entries = [str(RUN_DIR), str(CPQ_DIR), str(PIPELINE_DIR)]
    existing = os.environ.get("PYTHONPATH", "").strip()
    if existing:
        entries.append(existing)
    return os.pathsep.join(entries)


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def seed_missing_week_workbooks(config_path: Path) -> int:
    cfg = load_yaml(config_path)
    cfg_dir = config_path.parent.resolve()
    previous_week_label = cfg.get("previous_week_label")
    if not isinstance(previous_week_label, str) or not previous_week_label.strip():
        log.error("previous_week_label is required to seed missing week workbooks.")
        return 1
    previous_week_label = previous_week_label.strip()

    pipe = cfg.get("pipeline")
    if not isinstance(pipe, dict):
        log.error("pipeline mapping is required")
        return 1

    for key in SEED_FROM_PREVIOUS_KEYS:
        raw = pipe.get(key)
        if not isinstance(raw, str) or not raw.strip():
            log.info("Seed skip: pipeline.%s is not configured.", key)
            continue

        target = resolve_under(cfg_dir, expand_config_value(raw.strip(), cfg))
        if target.is_file():
            log.info("Seed exists: %s", target)
            continue

        source = resolve_under(
            cfg_dir,
            expand_config_value_for_week(raw.strip(), cfg, previous_week_label),
        )
        if not source.is_file():
            log.error(
                "Cannot seed pipeline.%s; previous-week source not found: %s",
                key,
                source,
            )
            return 1

        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
        log.info("Seed copied: %s -> %s", source.name, target)

    return 0


def main(argv: list[str] | None = None) -> int:
    setup_logging()
    p = argparse.ArgumentParser(description="Run week data pipeline via ingest.yml")
    p.add_argument(
        "--config",
        type=Path,
        default=None,
        help=f"Path to ingest YAML (default: {default_config_path()})",
    )
    args = p.parse_args(argv)
    cfg_path = (args.config or default_config_path()).resolve()
    if not cfg_path.is_file():
        log.error("Config not found: %s", cfg_path)
        return 1

    py = sys.executable
    env = os.environ.copy()
    env["PYTHONPATH"] = pipeline_pythonpath()

    seed_status = seed_missing_week_workbooks(cfg_path)
    if seed_status != 0:
        return seed_status

    for name, script in STEPS:
        if not script.is_file():
            log.error("Missing script %s", script)
            return 1
        cmd = [py, str(script), "--config", str(cfg_path)]
        if name == "cur_performance":
            cmd.append("--no-daily-zip")
        log.info("=== %s ===", name)
        r = subprocess.run(cmd, cwd=str(script.parent), env=env)
        if r.returncode != 0:
            log.error("Step failed: %s (exit %s)", name, r.returncode)
            return r.returncode

    log.info("=== All steps finished OK ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
