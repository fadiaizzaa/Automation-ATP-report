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

import argparse
import sys
from pathlib import Path

from build_current_performance_output import main as run_current_performance_output
from build_current_performance_output import run_fast_master_paste, setup_logging


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build and paste Current Performance Data")
    parser.add_argument(
        "--config",
        type=Path,
        default=Path("ingest.yml"),
        help="Path to ingest.yml. Default: ingest.yml",
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

    if not args.build_output:
        setup_logging()
        return run_fast_master_paste(args.config.resolve())

    forwarded = ["--config", str(args.config)]
    if args.template:
        forwarded.extend(["--template", str(args.template)])
    if args.output:
        forwarded.extend(["--output", str(args.output)])
    return run_current_performance_output(forwarded)


if __name__ == "__main__":
    sys.exit(main())
