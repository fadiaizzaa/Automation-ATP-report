"""Shared Excel external-reference and formula fill helpers."""

from __future__ import annotations

import shutil
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

import xlwings as xw


def external_ref(path: Path, sheet_name: str) -> str:
    folder = str(path.parent)
    if not folder.endswith("\\"):
        folder += "\\"
    safe_sheet = sheet_name.replace("'", "''")
    return f"'{folder}[{path.name}]{safe_sheet}'!"


def fill_formula_column(ws: xw.Sheet, col: str, last_row: int, formula: str, start_row: int = 3) -> None:
    ws.range(f"{col}{start_row}").formula = formula
    if last_row > start_row:
        ws.range(f"{col}{start_row}:{col}{last_row}").api.FillDown()


def force_excel_calculate(app: xw.App, wb: xw.Book) -> None:
    try:
        app.api.CalculateFullRebuild()
    except Exception:
        app.calculate()


def restore_workbook_from_backup(backup_path: Path, target_path: Path) -> bool:
    """Copy backup over target when a workbook update fails mid-run."""
    if not backup_path.is_file():
        return False
    shutil.copy2(backup_path, target_path)
    return True


def _check_windows_excel_prereqs() -> None:
    """Raise RuntimeError when pywin32/Excel automation cannot load on Windows."""
    import sys

    if sys.platform != "win32":
        return
    try:
        import pywintypes  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError as e:
        raise RuntimeError(
            "pywin32 could not be loaded (required for Excel on Windows). "
            "In the project venv run: pip install --force-reinstall pywin32 "
            "then: python .venv\\Scripts\\pywin32_postinstall.py -install"
        ) from e
    if xw.engines.active is None:
        raise RuntimeError(
            "xlwings found no Excel engine. Install Microsoft Excel, open it once, "
            "close all Excel windows, then retry."
        )


def _start_excel_app() -> xw.App:
    """Create an Excel automation instance or raise a clear error."""
    _check_windows_excel_prereqs()
    try:
        app = xw.App(visible=False, add_book=False)
    except AttributeError as e:
        if "apps" in str(e) or "active" in str(e):
            raise RuntimeError(
                "Could not start Microsoft Excel via xlwings. "
                "Install Excel on this PC, open Excel once manually, then close it and retry."
            ) from e
        raise
    except Exception as e:
        raise RuntimeError(
            f"Could not start Microsoft Excel via xlwings: {e}"
        ) from e
    app.display_alerts = False
    app.screen_updating = False
    return app


@contextmanager
def excel_session() -> Iterator[xw.App]:
    """Start a headless Excel instance and always close it."""
    app = _start_excel_app()
    try:
        yield app
    finally:
        for book in list(app.books):
            try:
                book.close()
            except Exception:
                pass
        try:
            app.screen_updating = True
        except Exception:
            pass
        try:
            app.quit()
        except Exception:
            pass


def force_recalc_file(path: Path) -> None:
    with excel_session() as app:
        wb = app.books.open(str(path), update_links=False)
        try:
            force_excel_calculate(app, wb)
            wb.save()
        finally:
            wb.close()
