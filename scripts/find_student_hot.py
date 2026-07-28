import sys, re, json
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from pathlib import Path
from openpyxl import load_workbook

PATH = "output/enriched_Chanmeet_Allocators_MASTER_Scored.xlsx"
wbf = load_workbook(PATH)
wsf = wbf.active
cfg = wbf["Scoring Config"]

# Title-points table rows 3-33 (per L-column formula)
title_kw = []
for r in range(3, 34):
    kw, pts = cfg[f"A{r}"].value, cfg[f"B{r}"].value
    if kw:
        title_kw.append((str(kw).lower(), pts))
print("Title keyword table (search order matters - first match wins):")
for kw, pts in title_kw:
    print(f"  {kw!r} -> {pts}")

def title_points(title):
    t = str(title or "").lower()
    for kw, pts in title_kw:
        if kw in t:
            return pts, kw
    return 0, None

# scan every row our script touched OR any row at all with student/educator language
cp = json.loads(Path("output/chanmeet_checkpoint.json").read_text(encoding="utf-8"))
STUDENT_RX = re.compile(r"\bstudent\b|\bprofessor\b|\bteacher\b|\bfaculty\b|\bacademic\b|\bph\.?d\.? candidate\b|\bintern\b", re.I)

print("\nScanning ALL 3000 rows for student/educator-sounding titles with points > 0...")
hits = 0
for r in range(2, wsf.max_row + 1):
    title = wsf[f"D{r}"].value or ""
    if STUDENT_RX.search(str(title)):
        pts, matched_kw = title_points(title)
        touched = r_s in cp["done"] if (r_s := str(r)) else False
        if pts > 0 or touched:
            hits += 1
            print(f"  row {r} [{'TOUCHED-'+cp['done'][str(r)]['status'] if touched else 'untouched'}]: "
                  f"title={title!r} -> title_pts={pts} (matched kw {matched_kw!r})")
print(f"\ntotal flagged: {hits}")
