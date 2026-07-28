"""SERVICE TRACK classifier — routes each HOT/WARM lead to the right
Xanalyst offer so outreach messages never collide.

Adds two columns to output/Chanmeet_Allocators_MASTER_Scored.xlsx right
after TEMP (R) and before Notes: "Service Track" (S) and "Track Primary"
(T). Notes and LinkedIn URL shift right to U/V; nothing else moves.
Safe to re-run: if the columns already exist they are refreshed in place.

Rule-based (free, no API), keyword hits over Company (B) + Title (D) +
Company Desc (F) + segment tags in Notes. Distinct-keyword counting,
word-boundary matching. Precedence and special cases per spec:
- "asset management" counts as RE only alongside other RE context
- "family office" counts as LP only when NO direct-property language
- BOTH when RE and fund signals coexist; Track Primary = side with more
  hits, tie goes to RE
- No signals -> Service Track left blank (never guess)
Only TEMP = HOT or WARM rows are classified.
"""
import re
import sys
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

PATH = Path(__file__).resolve().parent.parent / "output" / "Chanmeet_Allocators_MASTER_Scored.xlsx"

RE_KW = ["property", "properties", "realty", "real estate", "multifamily",
         "reit", "developer", "development", "housing", "acquisitions",
         "underwriting", "asset management", "yardi", "mri", "appfolio",
         "cam", "noi"]
GP_KW = ["fund manager", "private equity", "private credit",
         "capital partners", "fund accounting", "gp", "general partner",
         "waterfall", "carry", "nav", "investor reporting", "venture",
         "growth equity"]
LP_KW = ["endowment", "foundation", "pension", "ocio", "allocator",
         "lp positions", "manager selection", "family office",
         "wealth management", "private wealth", "private bank",
         "advisor to families"]

TRACK_RE = "RE"
TRACK_GP = "FUND ADMIN-GP"
TRACK_LP = "FUND ADMIN-LP"


def kw_rx(kw):
    return re.compile(r"(?<![a-z0-9])" + re.escape(kw) + r"(?![a-z0-9])", re.I)


RX = {kw: kw_rx(kw) for kw in RE_KW + GP_KW + LP_KW}
# "Multifamily Office" is a multi-family (wealth) office — an LP concept,
# not multifamily real estate. Don't let it fire the RE signal.
RX["multifamily"] = re.compile(
    r"(?<![a-z0-9])multifamily(?!\s+offices?)(?![a-z0-9])", re.I)


def matched(text, keywords):
    return {kw for kw in keywords if RX[kw].search(text)}


def classify(text):
    """Returns (track, primary, detail) — track may be '' when no signals."""
    re_hits = matched(text, RE_KW)
    # 'asset management' needs real-estate context to count as RE
    if "asset management" in re_hits and len(re_hits) == 1:
        re_hits = set()
    gp_hits = matched(text, GP_KW)
    lp_hits = matched(text, LP_KW)
    # family office rides LP only WITHOUT direct-property language
    if "family office" in lp_hits and re_hits:
        lp_hits.discard("family office")

    n_re, n_gp, n_lp = len(re_hits), len(gp_hits), len(lp_hits)
    n_fund = n_gp + n_lp
    detail = {"re": sorted(re_hits), "gp": sorted(gp_hits), "lp": sorted(lp_hits)}

    if n_re and n_fund:
        fund_track = TRACK_GP if n_gp >= n_lp else TRACK_LP
        primary = TRACK_RE if n_re >= n_fund else fund_track  # tie -> RE
        return "BOTH", primary, detail
    if n_re:
        return TRACK_RE, TRACK_RE, detail
    if n_gp or n_lp:
        track = TRACK_GP if n_gp >= n_lp else TRACK_LP
        return track, track, detail
    return "", "", detail


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    wb = load_workbook(PATH)
    ws = wb.active

    already = str(ws["S1"].value or "").strip() == "Service Track"
    notes_col = "U" if already else "S"

    # snapshot the inputs BEFORE any column shifting
    rows = {}
    for r in range(2, ws.max_row + 1):
        temp = str(ws[f"R{r}"].value or "").strip().upper()
        if temp not in ("HOT", "WARM"):
            continue
        text = " ".join(str(ws[f"{c}{r}"].value or "")
                        for c in ("B", "D", "F", notes_col))
        rows[r] = (temp, text)

    if not already:
        ws.insert_cols(19, 2)  # new S and T; Notes -> U, LinkedIn URL -> V
        ws["S1"].value = "Service Track"
        ws["T1"].value = "Track Primary"

    summary = {"HOT": Counter(), "WARM": Counter()}
    samples = []
    for r, (temp, text) in rows.items():
        track, primary, detail = classify(text)
        ws[f"S{r}"].value = track or None
        ws[f"T{r}"].value = primary or None
        summary[temp][track or "(blank)"] += 1
        if track:
            samples.append((r, temp, track, primary, detail,
                            str(ws[f"C{r}"].value or ""),
                            str(ws[f"B{r}"].value or "")[:45],
                            str(ws[f"D{r}"].value or "")[:45]))

    wb.save(PATH)

    print(f"classified rows (HOT+WARM): {len(rows)}\n")
    for tier in ("HOT", "WARM"):
        print(f"{tier}:")
        for track, n in summary[tier].most_common():
            print(f"  {track}: {n}")
        print()

    # sample spread: up to 3 per track, then fill to 10
    picked, seen = [], Counter()
    for s in samples:
        if seen[s[2]] < 3:
            picked.append(s)
            seen[s[2]] += 1
    for s in samples:
        if len(picked) >= 10:
            break
        if s not in picked:
            picked.append(s)
    print("SAMPLES:")
    for r, temp, track, primary, detail, person, comp, title in picked[:10]:
        hits = "; ".join(f"{k}:{','.join(v)}" for k, v in detail.items() if v)
        print(f"  row {r} [{temp}] {person} | {comp} | {title}")
        print(f"    -> {track}" + (f" (primary: {primary})" if track == "BOTH" else "")
              + f"  [{hits}]")


if __name__ == "__main__":
    main()
