import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from openpyxl import load_workbook

PATH = "output/enriched_Chanmeet_Allocators_MASTER_Scored.xlsx"
wbf = load_workbook(PATH)
wsf = wbf.active
wbv = load_workbook(PATH, data_only=True)
wsv = wbv.active

print("Q2 formula:", wsf["Q2"].value)
print("Q2 cached value (data_only):", wsv["Q2"].value)
print("R2 formula:", str(wsf["R2"].value)[:90])
print("R2 cached value (data_only):", wsv["R2"].value)
print()

count = 0
for r in range(2, wsv.max_row + 1):
    temp = wsv[f"R{r}"].value
    if temp == "HOT":
        count += 1
        if count <= 15:
            pts = [wsv[f"{c}{r}"].value for c in "LMNOP"]
            print(f"row {r}: person={wsf[f'C{r}'].value!r} | title={wsf[f'D{r}'].value!r} | "
                  f"company={wsf[f'B{r}'].value!r} | score(Q)={wsv[f'Q{r}'].value!r} | "
                  f"pts L-P={pts}")
print(f"\ntotal HOT rows (cached values): {count}")
