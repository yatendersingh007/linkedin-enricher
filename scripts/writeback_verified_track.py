"""Write the researched Service Track ledger back into the allocators xlsx,
replacing the keyword-guessed values in column S (Service Track) and T
(Track Primary). Reports counts and changed-vs-keyword stats.

Ledger line: row_id | Company | Verified Track | Evidence | Track Source
Track format from research: RE / FUND ADMIN-GP / FUND ADMIN-LP / UNVERIFIED /
"BOTH (primary: X)".
"""
import json
import re
import sys
from collections import Counter
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from openpyxl import load_workbook

PROJ = Path(__file__).resolve().parent.parent
XLSX = PROJ / "output" / "Chanmeet_Allocators_MASTER_Scored.xlsx"
LEDGER = PROJ / "output" / "service_track_ledger.csv"
SNAPSHOT = PROJ / "output" / "keyword_track_snapshot.json"

BOTH_RX = re.compile(r"^BOTH\s*\(primary:\s*([^)]+)\)$", re.I)


def parse_track(raw):
    """Return (service_track, track_primary) for writing to S/T."""
    t = raw.strip()
    m = BOTH_RX.match(t)
    if m:
        return "BOTH", m.group(1).strip().upper()
    return t, t


def norm_for_compare(t):
    """Collapse to a base category for the changed-vs-keyword comparison."""
    t = (t or "").strip().upper()
    if t.startswith("BOTH"):
        return "BOTH"
    if t in ("", "NONE"):
        return "(blank)"
    return t


lines = [l for l in LEDGER.read_text(encoding="utf-8").splitlines()[1:] if l.strip()]
ledger = {}
for line in lines:
    parts = [p.strip() for p in line.split(" | ")]
    if len(parts) < 3:
        continue
    row_id = int(parts[0])
    company, track = parts[1], parts[2]
    ledger[row_id] = (company, track)

keyword_snap = json.loads(SNAPSHOT.read_text(encoding="utf-8")) if SNAPSHOT.exists() else {}

wb = load_workbook(XLSX)
ws = wb.active

hot_counts, warm_counts = Counter(), Counter()
changed, unverified_rows = [], []
missing = []

for row_id, (company, raw_track) in ledger.items():
    temp = (str(ws[f"R{row_id}"].value or "").strip().upper())
    service_track, track_primary = parse_track(raw_track)

    ws[f"S{row_id}"].value = service_track
    ws[f"T{row_id}"].value = track_primary

    bucket = hot_counts if temp == "HOT" else (warm_counts if temp == "WARM" else None)
    if bucket is not None:
        bucket[norm_for_compare(service_track)] += 1

    if service_track.upper() == "UNVERIFIED":
        unverified_rows.append((row_id, company))

    old = norm_for_compare(keyword_snap.get(str(row_id), ""))
    new = norm_for_compare(service_track)
    if old != new:
        changed.append((row_id, old, new))

# rows that were supposed to be in the ledger but aren't (sanity check)
expected_ids = {int(k) for k in keyword_snap.keys()} if keyword_snap else set()
missing = sorted(expected_ids - set(ledger.keys()))

wb.save(XLSX)

print(f"wrote {len(ledger)} rows to columns S (Service Track) / T (Track Primary)")
print(f"saved: {XLSX}\n")

print("=== HOT tracks (researched) ===")
for t, n in hot_counts.most_common():
    print(f"  {t}: {n}")
print(f"  TOTAL HOT: {sum(hot_counts.values())}\n")

print("=== WARM tracks (researched) ===")
for t, n in warm_counts.most_common():
    print(f"  {t}: {n}")
print(f"  TOTAL WARM: {sum(warm_counts.values())}\n")

print(f"=== Changed vs keyword pass: {len(changed)} / {len(ledger)} rows ===")
change_matrix = Counter((old, new) for _, old, new in changed)
for (old, new), n in change_matrix.most_common(15):
    print(f"  {old} -> {new}: {n}")

print(f"\n=== UNVERIFIED: {len(unverified_rows)} rows ===")
for row_id, company in unverified_rows[:15]:
    print(f"  row {row_id}: {company}")
if len(unverified_rows) > 15:
    print(f"  ... and {len(unverified_rows) - 15} more")

if missing:
    print(f"\n*** WARNING: {len(missing)} rows expected but missing from ledger: {missing[:20]}")
else:
    print(f"\nintegrity check: all {len(keyword_snap)} queued rows accounted for in ledger")
