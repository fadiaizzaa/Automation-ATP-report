from __future__ import annotations

import re
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


RAW = Path(r"C:\Users\fadia\Downloads\Huawei-Fadia Izza Nabila\ATP NMS\input\Week 19\NMS\Week 19_Current Performance Data.xlsx")
RAW_1 = Path(r"C:\Users\fadia\Downloads\Huawei-Fadia Izza Nabila\ATP NMS\input\Week 19\NMS\Week 19_Current Performance Data_1.xlsx")
OUTPUT = Path(r"C:\Users\fadia\Downloads\Huawei-Fadia Izza Nabila\final-project-huawei-fadia\outputs\week19_current_performance_data_output.xlsx")

RAW_LAST_ROW = 500_000
RAW_1_HEADER_ROWS = 1
RAW_1_TOTAL_ROWS = 365_389
FINAL_LAST_ROW = RAW_LAST_ROW + (RAW_1_TOTAL_ROWS - RAW_1_HEADER_ROWS)
FINAL_RECORDS = FINAL_LAST_ROW - 8

LINE_PATTERNS = [
    "*LOG*",
    "*N30*",
    "*N40*",
    "*N50*",
    "*N60*",
    "*NS4*",
    "*NS3*",
    "*ND2*",
    "*LSX*",
    "*LDX*",
    "*ELOM*",
    "*LSC*",
    "*LTX*",
    "*LQM*",
    "*LDC*",
]
AMP_PATTERNS = ["*VA*", "*OAU*", "*OBU*", "*RAU*", "*RPC*", "*DAP*", "*MD40*", "*RAPXF*", "*MR4*", "*AFS*", "*OPU*", "*WSMD*"]
OSC_PATTERNS = ["*ST*", "*SC*"]

ROW_RE = re.compile(br'<row r="(\d+)"[^>]*>.*?</row>')
ROW_R_RE = re.compile(br'<row r="\d+"')
CELL_REF_RE = re.compile(br' r="([A-H])\d+"')
VALUE_RE = re.compile(br'<v>(\d+)</v>')


def pattern_array(patterns: list[str]) -> str:
    return "{" + ",".join(f'"{pattern}"' for pattern in patterns) + "}"


def formula_cell(col: str, row_idx: int, patterns: list[str]) -> bytes:
    body = f'SUM(COUNTIF(A{row_idx}, {pattern_array(patterns)})) &gt; 0'
    return f'<c r="{col}{row_idx}"><f>{body}</f></c>'.encode("utf-8")


def filter_cells(row_idx: int) -> bytes:
    return (
        formula_cell("I", row_idx, LINE_PATTERNS)
        + formula_cell("J", row_idx, AMP_PATTERNS)
        + formula_cell("K", row_idx, OSC_PATTERNS)
    )


def inline_header(col: str, text: str) -> bytes:
    return f'<c r="{col}8" t="inlineStr"><is><t>{text}</t></is></c>'.encode("utf-8")


def enhance_raw_row(row_xml: bytes, row_idx: int) -> bytes:
    row_xml = row_xml.replace(b'spans="1:8"', b'spans="1:11"', 1)
    if row_idx == 8:
        insert = (
            inline_header("I", "Filter Line Board")
            + inline_header("J", "Filter Amp Board")
            + inline_header("K", "Filter OSC Board")
        )
        return row_xml.replace(b"</row>", insert + b"</row>", 1)
    if row_idx >= 9:
        return row_xml.replace(b"</row>", filter_cells(row_idx) + b"</row>", 1)
    return row_xml


def shift_raw_1_row(row_xml: bytes, source_row_idx: int, shared_string_offset: int) -> tuple[int, bytes] | None:
    if source_row_idx <= RAW_1_HEADER_ROWS:
        return None
    target_row_idx = RAW_LAST_ROW + source_row_idx - RAW_1_HEADER_ROWS
    row_xml = ROW_R_RE.sub(f'<row r="{target_row_idx}"'.encode("utf-8"), row_xml, count=1)
    row_xml = row_xml.replace(b'spans="1:8"', b'spans="1:11"', 1)
    row_xml = CELL_REF_RE.sub(lambda m: b' r="' + m.group(1) + str(target_row_idx).encode("ascii") + b'"', row_xml)
    row_xml = VALUE_RE.sub(
        lambda m: b"<v>" + str(int(m.group(1)) + shared_string_offset).encode("ascii") + b"</v>",
        row_xml,
    )
    row_xml = row_xml.replace(b"</row>", filter_cells(target_row_idx) + b"</row>", 1)
    return target_row_idx, row_xml


def split_sheet_xml(sheet_xml: bytes) -> tuple[bytes, bytes, bytes]:
    start = sheet_xml.index(b"<sheetData>") + len(b"<sheetData>")
    end = sheet_xml.index(b"</sheetData>")
    prefix = sheet_xml[:start].replace(b'<dimension ref="A1:H500000"/>', f'<dimension ref="A1:K{FINAL_LAST_ROW}"/>'.encode("ascii"))
    body = sheet_xml[start:end]
    suffix = sheet_xml[end:]
    return prefix, body, suffix


def update_shared_strings(raw_data: bytes, raw_1_data: bytes) -> bytes:
    raw_unique = int(re.search(br'uniqueCount="(\d+)"', raw_data).group(1))
    raw_count = int(re.search(br'count="(\d+)"', raw_data).group(1))
    raw_1_unique = int(re.search(br'uniqueCount="(\d+)"', raw_1_data).group(1))
    raw_1_count = int(re.search(br'count="(\d+)"', raw_1_data).group(1))

    raw_inner_end = raw_data.rindex(b"</sst>")
    raw_1_inner_start = raw_1_data.index(b"<si>")
    raw_1_inner_end = raw_1_data.rindex(b"</sst>")

    merged = raw_data[:raw_inner_end] + raw_1_data[raw_1_inner_start:raw_1_inner_end] + b"</sst>"
    merged = merged.replace(b"Total 865,380 Records", f"Total {FINAL_RECORDS:,} Records".encode("utf-8"), 1)
    merged = re.sub(br'count="\d+"', f'count="{raw_count + raw_1_count - 8}"'.encode("ascii"), merged, count=1)
    merged = re.sub(br'uniqueCount="\d+"', f'uniqueCount="{raw_unique + raw_1_unique}"'.encode("ascii"), merged, count=1)
    return merged


def shared_string_count(shared_xml: bytes) -> int:
    return int(re.search(br'uniqueCount="(\d+)"', shared_xml).group(1))


def write_sheet(out_zip: ZipFile, raw_zip: ZipFile, raw_1_zip: ZipFile, shared_string_offset: int) -> None:
    raw_sheet = raw_zip.read("xl/worksheets/sheet1.xml")
    raw_1_sheet = raw_1_zip.read("xl/worksheets/sheet1.xml")
    prefix, raw_body, suffix = split_sheet_xml(raw_sheet)
    _, raw_1_body, _ = split_sheet_xml(raw_1_sheet.replace(b'<dimension ref="A1:H365389"/>', b'<dimension ref="A1:H365389"/>'))

    with out_zip.open("xl/worksheets/sheet1.xml", "w") as dest:
        dest.write(prefix)
        for match in ROW_RE.finditer(raw_body):
            row_idx = int(match.group(1))
            dest.write(enhance_raw_row(match.group(0), row_idx))
        for match in ROW_RE.finditer(raw_1_body):
            shifted = shift_raw_1_row(match.group(0), int(match.group(1)), shared_string_offset)
            if shifted is not None:
                dest.write(shifted[1])
        dest.write(suffix)


def main() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(RAW, "r") as raw_zip, ZipFile(RAW_1, "r") as raw_1_zip, ZipFile(OUTPUT, "w", ZIP_DEFLATED, compresslevel=6, allowZip64=True) as out_zip:
        raw_shared = raw_zip.read("xl/sharedStrings.xml")
        raw_1_shared = raw_1_zip.read("xl/sharedStrings.xml")
        shared_offset = shared_string_count(raw_shared)

        for info in raw_zip.infolist():
            if info.filename == "xl/worksheets/sheet1.xml":
                write_sheet(out_zip, raw_zip, raw_1_zip, shared_offset)
            elif info.filename == "xl/sharedStrings.xml":
                out_zip.writestr(info, update_shared_strings(raw_shared, raw_1_shared))
            else:
                out_zip.writestr(info, raw_zip.read(info.filename))
    print(OUTPUT)


if __name__ == "__main__":
    main()
