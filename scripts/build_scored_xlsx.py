"""Build the sorted scored-leads Excel from an append-only ledger.

Usage: python scripts/build_scored_xlsx.py --ledger <ledger.csv> --out <scored.xlsx>

The ledger is the source of truth; this file can be rebuilt anytime.
Sheet 1 "Scored Leads": all rows sorted HOT, HOT ACCOUNT, WARM, COLD,
REVIEW, UNVERIFIED, DISQUALIFIED (score descending within a tier).
Sheet 2 "Summary": tier counts, HOT deep-research list, HOT ACCOUNT
contact-discovery list. Prints a JSON summary to stdout.
"""
import argparse
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

TIER_ORDER = ["HOT", "HOT ACCOUNT", "WARM", "COLD", "REVIEW", "UNVERIFIED",
              "DISQUALIFIED"]
COLS = ["row_id", "Company", "Contact", "Title", "HQ Location", "Sector",
        "Est. Size", "Score", "Tier", "Top Reasons", "Red Flags",
        "Key Evidence Sources", "Service Track"]


def norm_tier(raw):
    t = (raw or "").strip().upper()
    # 'HOT ACCOUNT' must match before 'HOT'
    for k in sorted(TIER_ORDER, key=len, reverse=True):
        if t.startswith(k):
            return k
    return "UNVERIFIED"


def parse_line(line):
    parts = [p.strip() for p in line.split(" | ")]
    if len(parts) > 13:  # stray pipes: fold extras into Top Reasons
        parts = parts[:9] + [" / ".join(parts[9:-3])] + parts[-3:]
    while len(parts) < 13:  # old 12-column rows get a blank Service Track
        parts.append("")
    try:
        parts[7] = int(float(parts[7]))
    except ValueError:
        parts[7] = 0
    return parts


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ledger", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()

    lines = [l for l in Path(a.ledger).read_text(encoding="utf-8").splitlines()[1:]
             if l.strip()]
    rows = [parse_line(l) for l in lines]
    rows.sort(key=lambda r: (TIER_ORDER.index(norm_tier(r[8])), -r[7]))

    wb = Workbook()
    ws = wb.active
    ws.title = "Scored Leads"
    ws.append(COLS)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append(r)
    ws.freeze_panes = "A2"

    counts = {}
    for r in rows:
        counts[norm_tier(r[8])] = counts.get(norm_tier(r[8]), 0) + 1

    s = wb.create_sheet("Summary")
    s.append(["ICP scoring summary"])
    s["A1"].font = Font(bold=True, size=14)
    s.append([])
    s.append(["Tier", "Count"])
    s["A3"].font = s["B3"].font = Font(bold=True)
    for t in TIER_ORDER:
        if counts.get(t):
            s.append([t, counts[t]])
    s.append([])
    s.append(["HOT - deepest pre-outreach research first:"])
    s[f"A{s.max_row}"].font = Font(bold=True)
    for r in rows:
        if norm_tier(r[8]) == "HOT":
            s.append([r[1], r[2], r[3], f"score {r[7]}"])
    s.append([])
    s.append(["HOT ACCOUNT - need contact discovery (find CFO/Controller/MP):"])
    s[f"A{s.max_row}"].font = Font(bold=True)
    for r in rows:
        if norm_tier(r[8]) == "HOT ACCOUNT":
            s.append([r[1], f"current contact: {r[2]} ({r[3]})", f"score {r[7]}"])

    out = Path(a.out)
    out.parent.mkdir(exist_ok=True)
    wb.save(out)
    print(json.dumps({"total": len(rows), "counts": counts, "xlsx": str(out)}))


if __name__ == "__main__":
    main()
