# scripts/convert_uploaded_excels.py
import pandas as pd
from pathlib import Path

UPLOADED_DIR = Path("uploaded_excels")
OUT_DIR = Path("data")
OUT_DIR.mkdir(exist_ok=True)

def safe_sheet_name(s):
    return "".join(c if c.isalnum() else "_" for c in s).lower()

def excel_to_jsons(path):
    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        print(f"Failed to open {path}: {e}")
        return []
    produced = []
    ticker = path.stem.upper()
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet).fillna("")
            safe_sheet = safe_sheet_name(sheet)
            out_file = OUT_DIR / f"{ticker}__{safe_sheet}.json"
            df.to_json(out_file, orient="records", force_ascii=False)
            produced.append(out_file)
        except Exception as e:
            print(f"Failed to convert sheet {sheet} in {path}: {e}")
    return produced

def main():
    files = list(UPLOADED_DIR.glob("*.xls*"))
    if not files:
        print("No excel files found in uploaded_excels/")
        return
    for f in files:
        print("Converting", f)
        produced = excel_to_jsons(f)
        print("Produced:", [p.name for p in produced])

if __name__ == "__main__":
    main()
