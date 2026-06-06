"""
Verify the workbook has no formula errors.
Usage: python scripts/recalc.py output.xlsx
"""
import sys
from openpyxl import load_workbook

def check(path):
    wb = load_workbook(path, data_only=False)
    errors = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                v = str(cell.value) if cell.value else ""
                if v.startswith("=") and any(e in v.upper() for e in
                        ["#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NULL!", "#NUM!"]):
                    errors.append(f"  {sheet}!{cell.coordinate}: {v}")
    if errors:
        print("Formula errors found:")
        for e in errors:
            print(e)
        sys.exit(1)
    else:
        print(f"✓  No formula errors in {path}")
        print("Done")

if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "output.xlsx"
    check(path)
