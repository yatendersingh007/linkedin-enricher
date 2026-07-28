"""Recompute SCORE (Q) and TEMP (R) as real numbers/text in Python, replicating
the exact Scoring Config rubric — so the deliverable is correct regardless of
whether Excel recalculates formulas on open.

Why this exists: openpyxl cannot evaluate formulas. After writing new company
names into column B, the old cached Q/R values (from BEFORE our edits) stay
displayed unless Excel does a full recalc. This script removes that
dependency entirely by computing L-R as plain values, using the same
keyword tables as the workbook's own formulas (read live from Scoring Config,
so tuning the config there and re-running this script picks up any changes).

Also enforces: title cells that are actually "City, State, Country" strings
(a data-quality bug from an earlier enrichment pass) are cleared before
scoring, so they can never masquerade as a finance title.

Usage: python scripts/rescore_allocators.py [--in PATH] [--out PATH]
Default in/out: output/enriched_Chanmeet_Allocators_MASTER_Scored.xlsx (in place)
"""
import argparse
import re
import sys
from pathlib import Path

PROJ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJ))

from openpyxl import load_workbook
from enrich import is_location_string

DEFAULT = PROJ / "output" / "enriched_Chanmeet_Allocators_MASTER_Scored.xlsx"


def load_keyword_table(cfg, row_start, row_end):
    table = []
    for r in range(row_start, row_end + 1):
        kw, pts = cfg[f"A{r}"].value, cfg[f"B{r}"].value
        if kw is not None and str(kw).strip():
            table.append((str(kw).lower(), pts if pts is not None else 0))
    return table


def first_match(table, haystack):
    """Excel SEARCH+MATCH(TRUE) semantics: first keyword (in table order)
    that appears as a substring of haystack, case-insensitive."""
    h = str(haystack or "").lower()
    for kw, pts in table:
        if kw in h:
            return pts, kw
    return None, None


def cfg_scalar(cfg, row):
    return cfg[f"B{row}"].value


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", default=str(DEFAULT))
    ap.add_argument("--out", default=None, help="default: overwrite --in")
    args = ap.parse_args()
    inp = Path(args.inp)
    out = Path(args.out) if args.out else inp

    wb = load_workbook(inp)
    ws = wb.active
    cfg = wb["Scoring Config"]

    title_table = load_keyword_table(cfg, 3, 33)
    sector_table = load_keyword_table(cfg, 38, 79)
    geo_table = load_keyword_table(cfg, 84, 97)
    size_20_1000 = cfg_scalar(cfg, 102)
    size_edge = cfg_scalar(cfg, 103)
    size_huge = cfg_scalar(cfg, 104)
    size_tiny = cfg_scalar(cfg, 105)
    pts_event = cfg_scalar(cfg, 110)
    pts_linkedin = cfg_scalar(cfg, 111)
    pts_email = cfg_scalar(cfg, 112)
    hot_thresh = cfg_scalar(cfg, 116)
    warm_thresh = cfg_scalar(cfg, 117)

    fixed_titles = 0
    tier_counts = {}
    for r in range(2, ws.max_row + 1):
        B = ws[f"B{r}"].value
        company = str(B).strip() if B else ""
        if not company:
            for c in "LMNOPQR":
                ws[f"{c}{r}"].value = None
            continue

        D = ws[f"D{r}"].value
        if D and is_location_string(D):
            ws[f"D{r}"].value = None
            D = None
            fixed_titles += 1

        E = ws[f"E{r}"].value or ""
        F = ws[f"F{r}"].value or ""
        G = ws[f"G{r}"].value
        H = str(ws[f"H{r}"].value or "").lower()
        I = str(ws[f"I{r}"].value or "").lower()
        J = str(ws[f"J{r}"].value or "").lower()

        title_pts, _ = first_match(title_table, D) if D else (0, None)
        title_pts = title_pts if D else ""  # L: "" when D blank, mirroring the formula
        sector_haystack = f"{F} {company} {D or ''}"
        sector_pts, _ = first_match(sector_table, sector_haystack)
        sector_pts = sector_pts if sector_pts is not None else 5  # IFERROR default
        geo_pts, _ = first_match(geo_table, E) if E else (0, None)
        geo_pts = geo_pts if E else ""

        if G in (None, ""):
            size_pts = 0
        else:
            try:
                g = float(G)
            except (TypeError, ValueError):
                g = None
            if g is None:
                size_pts = 0
            elif 20 <= g <= 1000:
                size_pts = size_20_1000
            elif (10 <= g < 20) or (1000 < g <= 5000):
                size_pts = size_edge
            elif g > 5000:
                size_pts = size_huge
            else:
                size_pts = size_tiny

        engage_pts = (pts_event if H == "yes" else 0) \
            + (pts_linkedin if I in ("connected", "replied") else 0) \
            + (pts_email if J in ("opened", "replied") else 0)

        ws[f"L{r}"].value = title_pts
        ws[f"M{r}"].value = sector_pts
        ws[f"N{r}"].value = geo_pts
        ws[f"O{r}"].value = size_pts
        ws[f"P{r}"].value = engage_pts

        numeric_parts = [v for v in (title_pts, sector_pts, geo_pts, size_pts, engage_pts)
                         if isinstance(v, (int, float))]
        score = sum(numeric_parts)
        ws[f"Q{r}"].value = score

        if sector_pts == 0 or geo_pts == 0:
            tier = "CHECK"
        elif score >= hot_thresh:
            tier = "HOT"
        elif score >= warm_thresh:
            tier = "WARM"
        else:
            tier = "COLD"
        ws[f"R{r}"].value = tier
        tier_counts[tier] = tier_counts.get(tier, 0) + 1

    wb.save(out)
    print(f"title cells cleared (were location strings): {fixed_titles}")
    print(f"tier counts (Python-computed, authoritative): {tier_counts}")
    print(f"saved: {out}")


if __name__ == "__main__":
    main()
