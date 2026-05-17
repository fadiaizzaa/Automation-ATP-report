"""Add pipeline package folders to sys.path (call setup before importing pipeline_config)."""

from __future__ import annotations

import sys
from pathlib import Path


def setup(caller_file: str | Path) -> Path:
    pipeline_dir = next(p for p in Path(caller_file).resolve().parents if p.name == "pipeline")
    project_root = pipeline_dir.parent
    for sub in ("run", "cpq", ""):
        folder = pipeline_dir / sub if sub else pipeline_dir
        entry = str(folder.resolve())
        if entry not in sys.path:
            sys.path.insert(0, entry)
    root_entry = str(project_root.resolve())
    if root_entry not in sys.path:
        sys.path.insert(0, root_entry)
    return project_root


def default_config_path(caller_file: str | Path) -> Path:
    return setup(caller_file) / "config" / "ingest.yml"
