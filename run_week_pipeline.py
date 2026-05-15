#!/usr/bin/env python3
"""Run the full week pipeline steps in order (single command)."""

from __future__ import annotations

import argparse
import logging
import subprocess
import sys
from pathlib import Path

log = logging.getLogger("run_week_pipeline")

STEPS: list[tuple[str, list[str]]] = [
    ("prepare_week_inputs", ["prepare_week_inputs.py", "--config"]),
    ("validate_week_inputs", ["validate_week_inputs.py", "--config"]),
    ("try_nrr_fiber", ["try_nrr_fiber.py", "--config"]),
    ("try_oms_trail", ["try_oms_trail.py", "--config"]),
    ("try_och_trail", ["try_och_trail.py", "--config"]),
    ("try_pasting_cpq2", ["try_pasting_cpq2.py", "--config"]),
    ("srr", ["srr.py", "--config"]),
]


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def main(argv: list[str] | None = None) -> int:
    setup_logging()
    p = argparse.ArgumentParser(description="Run week data pipeline via ingest.yml")
    p.add_argument("--config", type=Path, required=True, help="Path to ingest YAML")
    args = p.parse_args(argv)
    cfg_path = args.config.resolve()
    if not cfg_path.is_file():
        log.error("Config not found: %s", cfg_path)
        return 1

    root = Path(__file__).resolve().parent
    py = sys.executable

    for name, cmd_prefix in STEPS:
        script = root / cmd_prefix[0]
        if not script.is_file():
            log.error("Missing script %s", script)
            return 1
        cmd = [py, str(script), cmd_prefix[1], str(cfg_path)]
        log.info("=== %s ===", name)
        r = subprocess.run(cmd, cwd=str(root))
        if r.returncode != 0:
            log.error("Step failed: %s (exit %s)", name, r.returncode)
            return r.returncode

    log.info("=== All steps finished OK ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
