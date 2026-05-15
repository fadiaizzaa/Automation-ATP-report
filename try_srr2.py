"""
Service Routing Report Builder  (xlwings)
==========================================
Processes the raw Service Routing Report and produces a clean computed .xlsx,
updates master Lib, recalculates, pastes into Combine, then pastes OCh Route
into master CPQ_OCH.

PIPELINE ORDER
--------------
1. Copy raw SRR -> output file, delete rows 1-3 from both sheets
2. Set raw data cols as values, write coded formula cols, fill down
3. Update master Lib col C/D with new NEs from Trail E2E col H+S
4. Recalculate: open Combine first, then Master, then SRR output -> save
5. Paste SRR Trail E2E (filtered CF==1) into Combine Sheet1 cols O-AB
   Paste SRR OCh Route (filtered BM!=0) into Combine Sheet1 cols AD-AG
   Save Combine
6. Paste OCh Route A:BM (header + all data) as values -> master CPQ_OCH

PASTE MAP TO COMBINE (Sheet1, data from row 3)
----------------------------------------------
Source: SRR Trail E2E, rows where CF (Filter SL Needed) == 1
  CG -> O   (service name)
  CH -> P   (signal name)
  CI -> Q   (ASON)
  CJ -> R   (route type)
  CK -> S   (och name)
  BL -> T   (Circuit Group)
  BQ -> U   (ClientSrcID)
  BR -> V   (ClientSnkID)
  BS -> W   (OCHSrcID)
  BT -> X   (OCHSnkID)
  BZ -> Y   (W %)
  CA -> Z   (WnR %)
  CD -> AA  (#Sharing Ratio W %)
  CE -> AB  (#Sharing Ratio WnR %)

Source: SRR OCh Route, rows where BM (Filter) != 0
  BI -> AD  (OCH Name)
  BJ -> AE  (Signal Rate)
  BK -> AF  (OCH Name2)
  BL -> AG  (Circuit ID / OCH Name3)

WEEKLY UPDATES IN CONFIG
------------------------
  Paths come from ingest.yml: week_label, pipeline.master_workbook,
  pipeline.combine_workbook, and CPQ week inputs under output_base.

USAGE
-----
    python try_srr2.py --config ingest.yml
"""

from __future__ import annotations

import argparse
import logging
import re
import shutil
import sys
from pathlib import Path

import xlwings as xw

from pipeline_config import build_pipeline_paths, load_yaml

OCH_SHEET = "OCh Route"
E2E_SHEET = "Trail E2E Route"

OCH_RAW_LAST = "BD"
OCH_FML_LAST = "BM"
E2E_RAW_LAST = "BK"
E2E_FML_LAST = "CK"

FIRST_DATA_ROW = 2   # after rows 1-3 deleted: row1=header, row2=first data

LIB_SHEET        = "Lib"
MASTER_CPQ_SHEET = "CPQ_OCH"
COMBINE_SHEET    = "Sheet1"

# Combine paste starts at row 3 (row1=section labels, row2=headers)
COMBINE_DATA_START = 3


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


log = logging.getLogger("srr_builder")


# -- Helpers ---------------------------------------------------------------

def excel_app():
    app = xw.App(visible=False, add_book=False)
    app.display_alerts  = False
    app.screen_updating = False
    return app


def last_row_in_col(ws, col="A"):
    return ws.range(f"{col}{ws.cells.last_cell.row}").end("up").row


def to_flat(val):
    if val is None:
        return []
    if not isinstance(val, list):
        return [val]
    if val and isinstance(val[0], list):
        return [r[0] if r else None for r in val]
    return val


def lib_ref(master_path):
    return f"'[{Path(master_path).name}]{LIB_SHEET}'!"


def combine_ref(combine_path):
    return f"'[{Path(combine_path).name}]Sheet1'!"


def read_sheet_as_values(ws, first_col, last_col, first_row, last_row):
    """Read a rectangular range and return as list of lists."""
    vals = ws.range(f"{first_col}{first_row}:{last_col}{last_row}").value
    if vals is None:
        return []
    if not isinstance(vals[0], list):
        return [[v] for v in vals]
    return vals


def config_bool(value, default=False):
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "y", "on"}
    return bool(value)


# -- Formula headers + header-only styling --------------------------------

# RGB colors used only on header row. Body cells are not styled.
HDR_DARK_BLUE = (31, 78, 121)
HDR_YELLOW    = (255, 255, 0)
HDR_BLUE      = (91, 155, 213)
HDR_ORANGE    = (244, 176, 132)
HDR_GOLD      = (191, 143, 0)
HDR_WHITE     = (255, 255, 255)
HDR_PURPLE    = (112, 48, 160)
FONT_WHITE    = 0xFFFFFF
FONT_BLACK    = 0x000000


def och_formula_headers():
    return {
        "BE": "OCH_Source",
        "BF": "OCH_Sink",
        "BG": "OCHSrcID",
        "BH": "OCHSnkID",
        "BI": "OCH Name",
        "BJ": "Signal Rate",
        "BK": "OCH Name",
        "BL": "Circuit ID",
        "BM": "Filter",
    }


def e2e_formula_headers():
    return {
        "BL": "Circuit Group",
        "BM": "Client_Source",
        "BN": "Client_Sink",
        "BO": "OCH_Source",
        "BP": "OCH_Sink",
        "BQ": "ClientSrcID",
        "BR": "ClientSnkID",
        "BS": "OCHSrcID",
        "BT": "OCHSnkID",
        "BU": "OCH",
        "BV": "Client",
        "BW": "Preset",
        "BX": "Count",
        "BY": "Normalization",
        "BZ": "W %",
        "CA": "WnR %",
        "CB": "#Client",
        "CC": "#Normal",
        "CD": "#Sharing Ratio W %",
        "CE": "#Sharing Ratio WnR %",
        "CF": "Filter SL Needed",
        "CG": "service name",
        "CH": "signal name",
        "CI": "ASON",
        "CJ": "route type",
        "CK": "och name",
    }


def style_header_block(ws, address, fill_color, font_color=FONT_WHITE):
    """Style only header cells. This avoids slow formatting across thousands of rows."""
    rng = ws.range(address)
    rng.color = fill_color
    try:
        rng.api.Font.Bold = True
        rng.api.Font.Name = "Arial"
        rng.api.Font.Size = 9
        rng.api.Font.Color = font_color
        rng.api.HorizontalAlignment = -4108  # xlCenter
        rng.api.VerticalAlignment = -4108    # xlCenter
        rng.api.WrapText = True
        borders = rng.api.Borders
        borders.LineStyle = 1   # xlContinuous
        borders.Weight = 2      # xlThin
        borders.Color = 0xCCCCCC
    except Exception:
        pass


def apply_formula_headers_and_header_style(ws, sheet_name):
    """
    Writes formula header names and applies header-row-only colors.
    Does not style body cells and does not change formulas/data.
    """
    if sheet_name == OCH_SHEET:
        for col, header in och_formula_headers().items():
            ws.range(f"{col}1").value = header

        # Raw headers remain dark blue. Formula headers are custom groups.
        style_header_block(ws, f"A1:{OCH_RAW_LAST}1", HDR_DARK_BLUE, FONT_WHITE)
        style_header_block(ws, "BE1:BF1", HDR_YELLOW, FONT_BLACK)
        style_header_block(ws, "BG1:BH1", HDR_BLUE, FONT_WHITE)
        style_header_block(ws, "BI1:BM1", HDR_ORANGE, FONT_BLACK)

        try:
            ws.range(f"A:{OCH_RAW_LAST}").api.EntireColumn.ColumnWidth = 14
            ws.range("BE:BM").api.EntireColumn.ColumnWidth = 18
        except Exception:
            pass

    elif sheet_name == E2E_SHEET:
        for col, header in e2e_formula_headers().items():
            ws.range(f"{col}1").value = header

        # Raw headers remain dark blue. Formula headers are custom groups.
        style_header_block(ws, f"A1:{E2E_RAW_LAST}1", HDR_DARK_BLUE, FONT_WHITE)
        style_header_block(ws, "BL1:BL1", HDR_GOLD, FONT_WHITE)
        style_header_block(ws, "BM1:BP1", HDR_YELLOW, FONT_BLACK)
        style_header_block(ws, "BQ1:BT1", HDR_BLUE, FONT_WHITE)
        style_header_block(ws, "BU1:CA1", HDR_WHITE, FONT_BLACK)
        style_header_block(ws, "CB1:CF1", HDR_PURPLE, FONT_WHITE)
        style_header_block(ws, "CG1:CK1", HDR_ORANGE, FONT_BLACK)

        try:
            ws.range(f"A:{E2E_RAW_LAST}").api.EntireColumn.ColumnWidth = 14
            ws.range("BL:CK").api.EntireColumn.ColumnWidth = 18
        except Exception:
            pass

    try:
        ws.range("1:1").api.RowHeight = 24
        ws.activate()
        ws.range("A2").select()
        ws.book.app.api.ActiveWindow.FreezePanes = False
        ws.book.app.api.ActiveWindow.FreezePanes = True
    except Exception:
        pass


# -- Formula maps ----------------------------------------------------------

def och_formulas(r, lib, cmb):
    return {
        "BE": f'=LEFT(M{r},4)&"-"&RIGHT(O{r},1)&"-"&TEXT(Q{r},"00")&"-"&TEXT(IF(ISNUMBER(SEARCH("IN/OUT",S{r},1)),"1",RIGHT(S{r},1)),"00")',
        "BF": f'=LEFT(AC{r},4)&"-"&RIGHT(AE{r},1)&"-"&TEXT(AG{r},"00")&"-"&TEXT(IF(ISNUMBER(SEARCH("IN/OUT",AI{r},1)),"1",RIGHT(AI{r},1)),"00")',
        "BG": f'=LEFT(M{r},4)&RIGHT(P{r},1)&TEXT(Q{r},"00")&TEXT(IF(ISNUMBER(SEARCH("IN/OUT",S{r},1)),"1",RIGHT(S{r},1)),"00")',
        "BH": f'=LEFT(AC{r},4)&RIGHT(AF{r},1)&TEXT(AG{r},"00")&TEXT(IF(ISNUMBER(SEARCH("IN/OUT",AI{r},1)),"1",RIGHT(AI{r},1)),"00")',
        "BI": f"=C{r}",
        "BJ": f"=E{r}",
        "BK": f'=IF(AND(G{r}="-",H{r}="-"),"-",IF(H{r}="-",TRIM(TEXT(LEFT(G{r},FIND("THz",G{r})-1)*1,"0.##"))&SUBSTITUTE(MID(G{r},FIND("±",SUBSTITUTE(G{r},"+-","±")),LEN(G{r})),"0GHz","GHz"),TEXT(H{r},"00")))&"_"&VLOOKUP(E{r},{lib}$AC:$AE,2,FALSE)&"_"&VLOOKUP(VALUE(LEFT(M{r},4)),{lib}$C:$D,2,FALSE)&"-"&VLOOKUP(VALUE(LEFT(AC{r},4)),{lib}$C:$D,2,FALSE)',
        "BL": f'=LEFT(M{r},4)&"_"&LEFT(AC{r},4)&"/"&IF(AND(G{r}="-",H{r}="-"),"-",IF(H{r}="-",TRIM(TEXT(LEFT(G{r},FIND("THz",G{r})-1)*1,"0.##"))&SUBSTITUTE(MID(G{r},FIND("±",SUBSTITUTE(G{r},"+-","±")),LEN(G{r})),"0GHz","GHz"),TEXT(H{r},"00")))&"/"&VLOOKUP(E{r},{lib}$AC:$AD,2,FALSE)',
        "BM": f"=COUNTIF({cmb}$S:$S,C{r})",
    }


def e2e_formulas(r, lib, cmb):
    return {
        "BL": f'=INDEX({lib}$D:$D,MATCH(VALUE(LEFT(H{r},4)),{lib}$C:$C,0))&"-"&INDEX({lib}$D:$D,MATCH(VALUE(LEFT(S{r},4)),{lib}$C:$C,0))',
        "BM": f'=LEFT(H{r},4)&"-"&RIGHT(J{r},1)&"-"&TEXT(M{r},"00")&"-"&TEXT(IF(ISNUMBER(SEARCH("RX/TX",O{r},1)),"1",IF(ISNUMBER(SEARCH("/TX",O{r},1)),MID(LEFT(O{r},FIND("/TX",O{r},1)-1),SEARCH("RX",O{r},1)+2,3),IF(SEARCH("/AP",O{r},1),MID(O{r},3,SEARCH("/",O{r},1)-3),"00"))),"00")',
        "BN": f'=LEFT(S{r},4)&"-"&RIGHT(U{r},1)&"-"&TEXT(X{r},"00")&"-"&TEXT(IF(ISNUMBER(SEARCH("RX/TX",Z{r},1)),"1",IF(ISNUMBER(SEARCH("/TX",Z{r},1)),MID(LEFT(Z{r},FIND("/TX",Z{r},1)-1),SEARCH("RX",Z{r},1)+2,3),IF(SEARCH("/AP",Z{r},1),MID(Z{r},3,SEARCH("/",Z{r},1)-3),"00"))),"00")',
        "BO": f'=LEFT(AG{r},4)&"-"&RIGHT(AJ{r},1)&"-"&TEXT(AL{r},"00")&"-"&TEXT(IF(ISNUMBER(SEARCH("IN/OUT",AO{r},1)),"1",RIGHT(AO{r},1)),"00")',
        "BP": f'=LEFT(AS{r},4)&"-"&RIGHT(AV{r},1)&"-"&TEXT(AX{r},"00")&"-"&TEXT(IF(ISNUMBER(SEARCH("IN/OUT",BA{r},1)),"1",RIGHT(BA{r},1)),"00")',
        "BQ": f'=LEFT(H{r},4)&RIGHT(J{r},1)&TEXT(M{r},"00")&TEXT(IF(ISNUMBER(SEARCH("RX/TX",O{r},1)),"1",IF(ISNUMBER(SEARCH("/TX",O{r},1)),MID(LEFT(O{r},FIND("/TX",O{r},1)-1),SEARCH("RX",O{r},1)+2,3),IF(SEARCH("/AP",O{r},1),MID(O{r},3,SEARCH("/",O{r},1)-3),"00"))),"00")',
        "BR": f'=LEFT(S{r},4)&RIGHT(U{r},1)&TEXT(X{r},"00")&TEXT(IF(ISNUMBER(SEARCH("RX/TX",Z{r},1)),"1",IF(ISNUMBER(SEARCH("/TX",Z{r},1)),MID(LEFT(Z{r},FIND("/TX",Z{r},1)-1),SEARCH("RX",Z{r},1)+2,3),IF(SEARCH("/AP",Z{r},1),MID(Z{r},3,SEARCH("/",Z{r},1)-3),"00"))),"00")',
        "BS": f'=LEFT(AG{r},4)&RIGHT(AJ{r},1)&TEXT(AL{r},"00")&TEXT(IF(ISNUMBER(SEARCH("IN/OUT",AO{r},1)),"1",RIGHT(AO{r},1)),"00")',
        "BT": f'=LEFT(AS{r},4)&RIGHT(AV{r},1)&TEXT(AX{r},"00")&TEXT(IF(ISNUMBER(SEARCH("IN/OUT",BA{r},1)),"1",RIGHT(BA{r},1)),"00")',
        "BU": f"=VALUE(LEFT(INDEX('OCh Route'!E:E,MATCH(AE{r},'OCh Route'!C:C,0)),LEN(INDEX('OCh Route'!E:E,MATCH(AE{r},'OCh Route'!C:C,0)))-1))",
        "BV": f"=VLOOKUP(D{r},{lib}$K:$L,2,FALSE)",
        "BW": f'=ISNUMBER(FIND("Preset",AD{r},1))',
        "BX": f"=COUNTIFS(BW:BW,BW{r},C:C,C{r},AE:AE,AE{r})",
        "BY": f"=BV{r}/BX{r}",
        "BZ": f'=SUMIFS(BV:BV,AE:AE,AE{r},AD:AD,"Working")/BU{r}',
        "CA": f"=SUMIFS(BY:BY,AE:AE,AE{r})/BU{r}",
        "CB": f"=IFERROR(VLOOKUP(BO{r}&\";\"&BM{r},{lib}$S:$S,2,FALSE),BV{r})",
        "CC": f"=CB{r}/BX{r}",
        "CD": f'=SUMIFS(CB:CB,AE:AE,AE{r},AD:AD,"Working")/BU{r}',
        "CE": f"=SUMIFS(CC:CC,AE:AE,AE{r})/BU{r}",
        "CF": f"=COUNTIF({cmb}$M:$M,C{r})",
        "CG": f"=C{r}",
        "CH": f"=D{r}",
        "CI": f"=E{r}",
        "CJ": f"=AD{r}",
        "CK": f"=AE{r}",
    }


# -- Step 1-4: Build output workbook ---------------------------------------

def build_output(src, out, master_path, combine_path):
    src = Path(src)
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, out)
    log.info(f"Copied {src.name} -> {out.name}")

    lib = lib_ref(master_path)
    cmb = combine_ref(combine_path)

    app = excel_app()
    try:
        wb = app.books.open(str(out), update_links=False)

        # Remove all sheets except the two we need
        for sh in list(wb.sheets):
            if sh.name not in (OCH_SHEET, E2E_SHEET):
                sh.delete()

        for sheet_name, raw_last, fml_first, fml_last, fml_fn in [
            (OCH_SHEET, OCH_RAW_LAST, "BE", OCH_FML_LAST, och_formulas),
            (E2E_SHEET, E2E_RAW_LAST, "BL", E2E_FML_LAST, e2e_formulas),
        ]:
            ws = wb.sheets[sheet_name]
            log.info(f"Processing {sheet_name} ...")

            # Delete rows 1-3 (metadata/merged cells)
            ws.api.Rows("1:3").Delete()
            log.info(f"  Rows 1-3 deleted")

            n_last = last_row_in_col(ws, "A")
            log.info(f"  Data rows: {FIRST_DATA_ROW} to {n_last}")

            # Convert raw area to plain values
            raw_rng = ws.range(f"A{FIRST_DATA_ROW}:{raw_last}{n_last}")
            raw_rng.value = raw_rng.value
            log.info(f"  Raw cols A:{raw_last} set as values")

            # Clear old formula area, write seed row, fill down
            ws.range(f"{fml_first}{FIRST_DATA_ROW}:{fml_last}{n_last}").clear_contents()
            formulas = fml_fn(FIRST_DATA_ROW, lib, cmb)
            for col, formula in formulas.items():
                ws.range(f"{col}{FIRST_DATA_ROW}").formula = formula
                if n_last > FIRST_DATA_ROW:
                    ws.range(f"{col}{FIRST_DATA_ROW}:{col}{n_last}").api.FillDown()
            log.info(f"  Formula cols {fml_first}:{fml_last} written and filled down")

            # Add formula header names and header-only styling.
            # This keeps the preferred pipeline unchanged while restoring the readable template look.
            apply_formula_headers_and_header_style(ws, sheet_name)
            log.info("  Formula headers and header-only styling applied")

        wb.save()
        wb.close()
        log.info(f"Saved -> {out.name}")
    finally:
        app.screen_updating = True
        app.quit()


# -- Step 5: Update master Lib C/D ----------------------------------------

def parse_ne(ne_name):
    if not ne_name or not isinstance(ne_name, str):
        return None, None
    m = re.match(r"^(\d{4})_.+?_(.+)$", ne_name.strip())
    if not m:
        return None, None
    return int(m.group(1)), m.group(2).strip()


def update_lib(master_path, output_path):
    log.info("Updating master Lib C/D ...")
    app = excel_app()
    try:
        wb_out  = app.books.open(str(output_path), update_links=False)
        ws_e2e  = wb_out.sheets[E2E_SHEET]
        n_last  = last_row_in_col(ws_e2e, "A")

        parsed = {}
        for col in ("H", "S"):
            vals = to_flat(ws_e2e.range(f"{col}{FIRST_DATA_ROW}:{col}{n_last}").value)
            for v in vals:
                ne_id, site = parse_ne(v)
                if ne_id is not None and ne_id not in parsed:
                    parsed[ne_id] = (str(v).strip(), site)
        wb_out.close()
        log.info(f"  Unique NEs from Trail E2E H+S: {len(parsed)}")

        wb_master = app.books.open(str(master_path), update_links=False)
        ws_lib    = wb_master.sheets[LIB_SHEET]
        last_lib  = last_row_in_col(ws_lib, "C")
        existing  = set()
        for v in to_flat(ws_lib.range(f"C2:C{last_lib}").value):
            try:
                existing.add(int(v))
            except (TypeError, ValueError):
                pass

        missing = {nid: v for nid, v in parsed.items() if nid not in existing}
        log.info(f"  Existing Lib NE IDs: {len(existing)}, missing: {len(missing)}")

        if missing:
            for i, (ne_id, (ne_name, site)) in enumerate(sorted(missing.items())):
                row = last_lib + 1 + i
                ws_lib.range(f"B{row}").value = ne_name
                ws_lib.range(f"C{row}").value = ne_id
                ws_lib.range(f"D{row}").value = site
                ws_lib.range(f"E{row}").formula = f"=COUNTIF(B:B,B{row})"
                log.info(f"  + row {row}: {ne_name} | {ne_id} | {site}")
            wb_master.save()
            log.info(f"  {len(missing)} NE(s) added. Master saved.")
        else:
            log.info("  Lib already up to date.")

        wb_master.close()
    finally:
        app.screen_updating = True
        app.quit()


# -- Step 6: Recalculate ---------------------------------------------------

def recalculate(master_path, output_path, combine_path, full_rebuild=True):
    mode = "full rebuild" if full_rebuild else "open dependencies + save only"
    log.info("Opening dependencies and saving SRR output (%s) ...", mode)
    app = excel_app()
    combine_wb = master_wb = service_wb = None
    try:
        combine_wb = app.books.open(str(combine_path), update_links=False)
        master_wb  = app.books.open(str(master_path),  update_links=False)
        service_wb = app.books.open(str(output_path),  update_links=False)
        if full_rebuild:
            try:
                app.api.CalculateFullRebuild()
            except Exception:
                app.calculate()
        service_wb.save()
        log.info("  Saved SRR output.")
    except Exception as e:
        log.warning(f"  SRR dependency open/save failed: {e}")
        log.warning("  Manual fix: open Combine, Master, SRR output -> save.")
    finally:
        for wb in (service_wb, master_wb, combine_wb):
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass
        app.screen_updating = True
        app.quit()


# -- Step 7: Paste into Combine --------------------------------------------

def paste_to_combine(output_path, combine_path):
    """
    Trail E2E (CF==1) -> Combine Sheet1 cols O-AB
    OCh Route  (BM!=0) -> Combine Sheet1 cols AD-AG
    Clears existing data in those cols first, then pastes from row 3.
    """
    log.info("Pasting SRR data into Combine ...")
    app = excel_app()
    srr_wb = combine_wb = None
    try:
        srr_wb     = app.books.open(str(output_path), update_links=False)
        combine_wb = app.books.open(str(combine_path), update_links=False)
        ws_e2e     = srr_wb.sheets[E2E_SHEET]
        ws_och     = srr_wb.sheets[OCH_SHEET]
        ws_cmb     = combine_wb.sheets[COMBINE_SHEET]

        # ── Trail E2E: filter CF==1 ───────────────────────────────
        # CF = col 84 (0-based index 83), after rows 1-3 deleted row2=first data
        n_e2e = last_row_in_col(ws_e2e, "A")
        log.info(f"  Trail E2E rows: {FIRST_DATA_ROW} to {n_e2e}")

        # Read all needed cols in one call for speed
        # CG-CK = cols 85-89, BL=64, BQ-BT=69-72, BZ-CA=78-79, CD-CE=82-83, CF=84
        # Read as full rows CG:CK range + individual cols
        cf_vals  = to_flat(ws_e2e.range(f"CF{FIRST_DATA_ROW}:CF{n_e2e}").value)
        cg_ck    = ws_e2e.range(f"CG{FIRST_DATA_ROW}:CK{n_e2e}").value  # 5 cols -> O-S
        bl_vals  = to_flat(ws_e2e.range(f"BL{FIRST_DATA_ROW}:BL{n_e2e}").value)  # -> T
        bq_bt    = ws_e2e.range(f"BQ{FIRST_DATA_ROW}:BT{n_e2e}").value  # 4 cols -> U-X
        bz_ca    = ws_e2e.range(f"BZ{FIRST_DATA_ROW}:CA{n_e2e}").value  # 2 cols -> Y-Z
        cd_ce    = ws_e2e.range(f"CD{FIRST_DATA_ROW}:CE{n_e2e}").value  # 2 cols -> AA-AB

        # Normalize multi-col ranges (xlwings returns list of lists)
        def norm(v, n_rows):
            if v is None:
                return [[None] * 1 for _ in range(n_rows)]
            if not isinstance(v, list):
                return [[v]]
            if not isinstance(v[0], list):
                return [[x] for x in v]
            return v

        n_rows = len(cf_vals)
        cg_ck = norm(cg_ck, n_rows)
        bq_bt = norm(bq_bt, n_rows)
        bz_ca = norm(bz_ca, n_rows)
        cd_ce = norm(cd_ce, n_rows)

        # Build filtered rows for Combine (CF==1)
        e2e_rows = []
        for i, cf in enumerate(cf_vals):
            try:
                keep = int(cf) == 1
            except (TypeError, ValueError):
                keep = False
            if not keep:
                continue
            row = (
                (cg_ck[i] if i < len(cg_ck) else [None]*5) +   # O-S (5 cols)
                [bl_vals[i] if i < len(bl_vals) else None] +    # T   (1 col)
                (bq_bt[i] if i < len(bq_bt) else [None]*4) +   # U-X (4 cols)
                (bz_ca[i] if i < len(bz_ca) else [None]*2) +   # Y-Z (2 cols)
                (cd_ce[i] if i < len(cd_ce) else [None]*2)     # AA-AB (2 cols)
            )  # total 14 cols -> O through AB
            e2e_rows.append(row)

        log.info(f"  Trail E2E rows with CF==1: {len(e2e_rows)}")

        # ── OCh Route: filter BM != 0 ─────────────────────────────
        n_och = last_row_in_col(ws_och, "A")
        log.info(f"  OCh Route rows: {FIRST_DATA_ROW} to {n_och}")

        bm_vals = to_flat(ws_och.range(f"BM{FIRST_DATA_ROW}:BM{n_och}").value)
        bi_bl   = ws_och.range(f"BI{FIRST_DATA_ROW}:BL{n_och}").value  # 4 cols -> AD-AG
        bi_bl   = norm(bi_bl, len(bm_vals))

        och_rows = []
        for i, bm in enumerate(bm_vals):
            try:
                keep = bm is not None and int(bm) != 0
            except (TypeError, ValueError):
                keep = False
            if not keep:
                continue
            och_rows.append(bi_bl[i] if i < len(bi_bl) else [None]*4)

        log.info(f"  OCh Route rows with BM!=0: {len(och_rows)}")

        # ── Clear and paste into Combine ──────────────────────────
        # Clear cols O-AB (15-28) and AD-AG (30-33) from row 3 down
        used_last = max(ws_cmb.used_range.last_cell.row, COMBINE_DATA_START)
        ws_cmb.range(f"O{COMBINE_DATA_START}:AB{used_last}").clear_contents()
        ws_cmb.range(f"AD{COMBINE_DATA_START}:AG{used_last}").clear_contents()

        if e2e_rows:
            ws_cmb.range(f"O{COMBINE_DATA_START}").value = e2e_rows
            log.info(f"  Pasted {len(e2e_rows)} rows to Combine cols O-AB")

        if och_rows:
            ws_cmb.range(f"AD{COMBINE_DATA_START}").value = och_rows
            log.info(f"  Pasted {len(och_rows)} rows to Combine cols AD-AG")

        combine_wb.save()
        log.info("  Combine saved.")

    finally:
        for wb in (srr_wb, combine_wb):
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass
        app.screen_updating = True
        app.quit()


# -- Step 8: Paste OCh Route into master CPQ_OCH --------------------------

def paste_to_cpq(master_path, output_path):
    log.info(f"Pasting OCh Route -> master {MASTER_CPQ_SHEET} ...")
    app = excel_app()
    srr_wb = master_wb = None
    try:
        srr_wb    = app.books.open(str(output_path), update_links=False)
        ws_src    = srr_wb.sheets[OCH_SHEET]
        n_last    = last_row_in_col(ws_src, "A")
        # Read full A:BM including header row 1
        values    = ws_src.range(f"A1:{OCH_FML_LAST}{n_last}").value
        srr_wb.close()
        srr_wb    = None

        master_wb = app.books.open(str(master_path), update_links=False)
        ws_dst    = master_wb.sheets[MASTER_CPQ_SHEET]
        used_last = max(ws_dst.used_range.last_cell.row, 1)
        ws_dst.range(f"A1:{OCH_FML_LAST}{used_last}").clear_contents()
        ws_dst.range("A1").value = values
        master_wb.save()
        log.info(f"  Pasted {n_last} rows (incl. header) into {MASTER_CPQ_SHEET}. Master saved.")

    finally:
        for wb in (srr_wb, master_wb):
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass
        app.screen_updating = True
        app.quit()


# -- Main ------------------------------------------------------------------

def run(config_path: Path) -> int:
    cfg = load_yaml(config_path)
    paths = build_pipeline_paths(cfg, config_path)

    src = paths.service_routing_raw
    out = paths.service_routing_computed_out
    master = paths.master_workbook
    combine = paths.combine_workbook

    for p, name in [(src, "Service Routing"), (master, "Master"), (combine, "Combine")]:
        if not p.exists():
            log.error("%s file not found: %s", name, p)
            return 1

    pipe = cfg.get("pipeline") if isinstance(cfg.get("pipeline"), dict) else {}
    make_master_backup = pipe.get("make_master_backup", True)
    full_excel_rebuild = config_bool(pipe.get("full_excel_rebuild"), True)
    if make_master_backup:
        backup = master.with_stem(master.stem + "_SRR_BACKUP")
        shutil.copy2(master, backup)
        log.info("Master backup -> %s", backup.name)

    log.info("Source  : %s", src.name)
    log.info("Output  : %s", out.name)
    log.info("Master  : %s", master.name)
    log.info("Combine : %s", combine.name)

    build_output(src, out, master, combine)
    update_lib(master, out)
    recalculate(master, out, combine, full_rebuild=full_excel_rebuild)
    paste_to_combine(out, combine)
    paste_to_cpq(master, out)

    log.info("Done.")
    log.info("  SRR output : %s", out.resolve())
    log.info("  Combine    : %s", combine.resolve())
    log.info("  Master     : %s", master.resolve())
    return 0


def main(argv: list[str] | None = None) -> int:
    setup_logging()
    p = argparse.ArgumentParser(description="Service Routing Report pipeline step")
    p.add_argument("--config", type=Path, required=True)
    args = p.parse_args(argv)
    cfg_path = args.config.resolve()
    if not cfg_path.is_file():
        log.error("Config not found: %s", cfg_path)
        return 1
    try:
        return run(cfg_path)
    except Exception:
        log.exception("try_srr2 failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
