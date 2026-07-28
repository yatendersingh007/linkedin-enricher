import sys, re, json
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from pathlib import Path
from openpyxl import load_workbook

PATH = "output/enriched_Chanmeet_Allocators_MASTER_Scored.xlsx"
wbf = load_workbook(PATH)
wsf = wbf.active
cfg = wbf["Scoring Config"]

# Scoring Config sector keyword table (rows 38-79 per the M-column formula)
sector_kw = []
for r in range(38, 80):
    kw, pts = cfg[f"A{r}"].value, cfg[f"B{r}"].value
    if kw:
        sector_kw.append((str(kw).lower(), r, pts))
edu_rows = [(kw, r, pts) for kw, r, pts in sector_kw
            if re.search(r"univers|college|school|student|professor|academ|educat", kw)]
print("Education-flavored sector keywords in Scoring Config:")
for kw, r, pts in edu_rows:
    print(f"  row {r}: {kw!r} -> {pts} pts")

LOCATION_RX = re.compile(
    r"^[A-Za-z .'-]+,\s*[A-Za-z .'-]+,\s*(United States|Canada)$|"
    r"^[A-Za-z .'-]+,\s*(United States|Canada)$", re.I)

cp = json.loads(Path("output/chanmeet_checkpoint.json").read_text(encoding="utf-8"))
bad = []
for row_s, d in cp["done"].items():
    r = int(row_s)
    title = wsf[f"D{r}"].value or ""
    company = wsf[f"B{r}"].value or ""
    if LOCATION_RX.match(str(title).strip()):
        bad.append((r, d["bucket"], d["status"], company, title))

print(f"\n{len(bad)} rows where Title (D) got overwritten with a LOCATION string:")
for r, bucket, status, company, title in bad:
    print(f"  row {r} ({bucket}/{status}): D={title!r} | B={company!r}")
