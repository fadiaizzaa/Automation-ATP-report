from pathlib import Path
import re
from zipfile import ZipFile
import xml.etree.ElementTree as ET


OUTPUT = Path(r"C:\Users\fadia\Downloads\Huawei-Fadia Izza Nabila\final-project-huawei-fadia\outputs\week19_current_performance_data_output.xlsx")


def cell_refs(row_xml: bytes) -> list[str]:
    return [match.decode("ascii") for match in re.findall(br'<c r="([A-Z]+\d+)"', row_xml)]


def main() -> None:
    with ZipFile(OUTPUT) as zf:
        sheet = zf.read("xl/worksheets/sheet1.xml")
        shared = zf.read("xl/sharedStrings.xml")
        shared_values = [
            "".join(si.itertext())
            for si in ET.fromstring(shared).iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si")
        ]
        dim = re.search(br'<dimension ref="([^"]+)"', sheet).group(1).decode("ascii")
        row_count = len(re.findall(br"<row ", sheet))
        formula_count = len(re.findall(br"<f>", sheet))
        shared_unique = re.search(br'uniqueCount="([^"]+)"', shared).group(1).decode("ascii")
        total_text_ok = b"Total 865,380 Records" in shared
        old_total_present = b"Total 865,380 Records" in shared

        print(f"dimension={dim}")
        print(f"row_count={row_count}")
        print(f"formula_count={formula_count}")
        print(f"shared_unique={shared_unique}")
        print(f"old_total_present={old_total_present}")

        for row_num in [8, 9, 500000, 500001, 865388]:
            match = re.search(fr'<row r="{row_num}"[^>]*>.*?</row>'.encode("ascii"), sheet)
            if not match:
                print(f"row_{row_num}=MISSING")
                continue
            row_xml = match.group(0)
            formulas = re.findall(br"<f>(.*?)</f>", row_xml)
            print(f"row_{row_num}_cells={cell_refs(row_xml)}")
            print(f"row_{row_num}_formula_count={len(formulas)}")
            if formulas:
                print(f"row_{row_num}_first_formula={formulas[0][:120].decode('utf-8')}")
            if row_num in [500000, 500001, 865388]:
                indexes = [int(value) for value in re.findall(br"<v>(\d+)</v>", row_xml)[:8]]
                print(f"row_{row_num}_sample_values={[shared_values[index] for index in indexes[:2]]}")


if __name__ == "__main__":
    main()
