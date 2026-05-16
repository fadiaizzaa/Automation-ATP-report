#!/usr/bin/env python3
"""
Current Performance Data helper.

This is the short/manual entrypoint for the full Current Performance process:
  1. read the capped main Current Performance workbook and any numbered continuations
  2. filter the rows directly in Python
  3. paste filtered Line Board rows into master sheet "OCH Performance"
  4. paste filtered Amplifier Board rows into master sheet "OAU"

Use --build-output only when you also want the huge merged formula workbook.

Paths are read from ingest.yml:
  - prepared NMS files come from output_base/week_label/NMS
  - master workbook comes from pipeline.master_workbook
  - template can come from pipeline.current_performance_template
  - output comes from pipeline.current_performance_output

Usage:
    python cur_performance.py
    python cur_performance.py --config ingest.yml
"""

from __future__ import annotations

import importlib.util
from pathlib import Path


def _bootstrap_import_paths() -> None:
    pipeline_dir = next(p for p in Path(__file__).resolve().parents if p.name == "pipeline")
    spec = importlib.util.spec_from_file_location("_pathsetup", pipeline_dir / "_pathsetup.py")
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    module.setup(__file__)


_bootstrap_import_paths()

import argparse
import sys

from build_current_performance_output import main as run_current_performance_output
from build_current_performance_output import run_fast_master_paste, setup_logging
from pipeline_config import default_config_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build and paste Current Performance Data")
    parser.add_argument(
        "--config",
        type=Path,
        default=None,
        help=f"Path to ingest.yml (default: {default_config_path()})",
    )
    parser.add_argument(
        "--template",
        type=Path,
        help="Override pipeline.current_performance_template from ingest.yml",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Override pipeline.current_performance_output from ingest.yml",
    )
    parser.add_argument(
        "--build-output",
        action="store_true",
        help="Also build the large merged formula workbook before pasting to master.",
    )
    args = parser.parse_args(argv)
    cfg_path = (args.config or default_config_path()).resolve()

    if not args.build_output:
        setup_logging()
        return run_fast_master_paste(cfg_path)

    forwarded = ["--config", str(cfg_path)]
    if args.template:
        forwarded.extend(["--template", str(args.template)])
    if args.output:
        forwarded.extend(["--output", str(args.output)])
    return run_current_performance_output(forwarded)


if __name__ == "__main__":
    sys.exit(main())
