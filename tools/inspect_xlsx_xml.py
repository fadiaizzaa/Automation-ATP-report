from pathlib import Path
from zipfile import ZipFile


FILES = [
    Path(r"C:\Users\fadia\Downloads\Huawei-Fadia Izza Nabila\ATP NMS\input\Week 19\NMS\Week 19_Current Performance Data.xlsx"),
    Path(r"C:\Users\fadia\Downloads\Huawei-Fadia Izza Nabila\ATP NMS\input\Week 19\NMS\Week 19_Current Performance Data_1.xlsx"),
]


def main() -> None:
    for path in FILES:
        with ZipFile(path) as zf:
            print("FILE", path)
            print([name for name in zf.namelist() if "sheet" in name.lower()][:20])
            data = zf.read("xl/worksheets/sheet1.xml")
            print("size", len(data))
            for marker in [b'<row r="8"', b'<row r="9"', b'<row r="499999"', b'<row r="500000"']:
                idx = data.find(marker)
                if idx >= 0:
                    print(marker, data[idx : idx + 800])
            print("tail", data[-1000:])


if __name__ == "__main__":
    main()
