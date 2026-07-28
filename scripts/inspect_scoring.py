import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from openpyxl import load_workbook

PATH = "output/enriched_Chanmeet_Allocators_MASTER_Scored.xlsx"
wbf = load_workbook(PATH)
wsf = wbf.active
cfg = wbf["Scoring Config"]

print("=== Scoring Config around row 116-117 (HOT/WARM thresholds) ===")
for r in range(110, 120):
    vals = [cfg[f"{c}{r}"].value for c in "ABC"]
    if any(v is not None for v in vals):
        print(f"  row {r}: {vals}")

print("\n=== L-P column formulas (row 2, as template) ===")
for c in "LMNOP":
    print(f"  {c}2: {wsf[f'{c}2'].value}")

print("\n=== Rows I RECOVERED (company filled by us) - checking L-P formulas exist ===")
import json
from pathlib import Path
cp = json.loads(Path("output/chanmeet_checkpoint.json").read_text(encoding="utf-8"))
recovered_rows = [int(r) for r, d in cp["done"].items() if d["status"] == "recovered"]
print(f"total recovered rows: {len(recovered_rows)}")
for r in recovered_rows[:10]:
    lp = [wsf[f"{c}{r}"].value for c in "LMNOPQR"]
    print(f"  row {r}: company={wsf[f'B{r}'].value!r} title={wsf[f'D{r}'].value!r} L-R={lp}")
