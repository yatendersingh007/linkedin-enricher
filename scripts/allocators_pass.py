"""Selective fetch+search pass for Chanmeet_Allocators_MASTER_Scored.xlsx.

Buckets (selection computed from the ORIGINAL file's cached formula values):
  RECOVER - Company (B) blank + LinkedIn URL (T) + US/Canada location (E):
            fill B with the extracted employer, refine D with the verified
            title, put the company LinkedIn URL in F.
  VERIFY  - TEMP (R) = HOT + Company (B) present: never overwrite B; on
            material disagreement (or no employer visible) append
            "MISMATCH: <found>" to Notes (S). Segment tags already in S are
            preserved - notes are appended, never overwritten.
  Blocked/failed rows get "NEEDS_RESEARCH" appended to S. No Agent endpoint.

Output: output/enriched_Chanmeet_Allocators_MASTER_Scored.xlsx (formulas
intact - Excel re-scores on open). Checkpoint: output/chanmeet_checkpoint.json
saved after every row; re-running skips finished rows.

Usage:
  python scripts/allocators_pass.py --test 10     # 5 recover + 5 verify
  python scripts/allocators_pass.py               # full selection
  python scripts/allocators_pass.py --workers 8
"""
import argparse
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

PROJ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJ))

from openpyxl import load_workbook

from enrich import (fetch_one, extract_from_fetch, get_tinyfish_client,
                    recover_via_search, is_location_string)

INPUT = PROJ / "input" / "Chanmeet_Allocators_MASTER_Scored.xlsx"
OUT = PROJ / "output" / "enriched_Chanmeet_Allocators_MASTER_Scored.xlsx"
CKPT = PROJ / "output" / "chanmeet_checkpoint.json"

COL = {"company": "B", "person": "C", "title": "D", "location": "E",
       "note_f": "F", "temp": "R", "notes": "S", "url": "T"}

GENERIC_TOKENS = {"llc", "inc", "lp", "llp", "ltd", "co", "corp", "corporation",
                  "company", "the", "group", "partners", "capital", "management",
                  "holdings", "of", "and", "a", "an", "firm", "fund", "funds",
                  "advisors", "advisers", "associates"}

TRANSIENT = ("unavailable", "timeout", "timed out", "rate limited", "unreachable")


def sval(ws, coord):
    v = ws[coord].value
    return str(v).strip() if v is not None and str(v).strip() else ""


def company_tokens(name):
    return {t for t in re.findall(r"[a-z0-9]+", str(name).lower())
            if len(t) > 1 and t not in GENERIC_TOKENS}


def materially_same(existing, found):
    a, b = company_tokens(existing), company_tokens(found)
    return bool(a and b and (a & b))


def append_note(ws, row, note):
    cur = sval(ws, f"{COL['notes']}{row}")
    if note in cur:
        return
    ws[f"{COL['notes']}{row}"].value = (cur + "; " + note) if cur else note


def clear_needs_research(ws, row):
    """A retried row that now succeeded must not keep its stale flag."""
    cur = sval(ws, f"{COL['notes']}{row}")
    if "NEEDS_RESEARCH" not in cur:
        return
    parts = [p.strip() for p in cur.split(";")]
    ws[f"{COL['notes']}{row}"].value = "; ".join(
        p for p in parts if p and p != "NEEDS_RESEARCH") or None


def load_checkpoint():
    if CKPT.exists():
        return json.loads(CKPT.read_text(encoding="utf-8"))
    return {"done": {}}


def save_checkpoint(cp):
    CKPT.parent.mkdir(exist_ok=True)
    tmp = CKPT.with_suffix(".tmp")
    tmp.write_text(json.dumps(cp, ensure_ascii=False), encoding="utf-8")
    tmp.replace(CKPT)


def select_rows(wsv):
    verify, recover = [], []
    for r in range(2, wsv.max_row + 1):
        B = sval(wsv, f"B{r}")
        T = sval(wsv, f"T{r}")
        E = sval(wsv, f"E{r}").lower()
        R = sval(wsv, f"R{r}").upper()
        has_url = "linkedin.com" in T.lower()
        uscan = "united states" in E or "canada" in E
        if R == "HOT" and B and has_url:
            verify.append(r)
        elif not B and has_url and uscan:
            recover.append(r)
    return verify, recover


def work_one(client, ctx):
    """Fetch, then search fallback. Returns (data|None, calls, last_error)."""
    fr, attempts, err = fetch_one(client, ctx["url"])
    data = extract_from_fetch(fr) if fr else None
    calls = attempts
    blocked = fr is not None and data is None  # page loaded but authwalled
    if not data or not data.get("company"):
        sdata, sattempts, serr = recover_via_search(client, {
            "name": ctx["person"], "title": ctx["title"],
            "location": ctx["location"], "linkedin_url": ctx["url"],
            "raw": ctx["title"],
        })
        calls += sattempts
        if sdata and sdata.get("company"):
            data = dict(data or {})
            data["company"] = sdata["company"]
        elif err is None:
            err = serr
    return data, calls, err, blocked


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("--test", type=int, default=0,
                    help="Process only N rows (split evenly across buckets)")
    ap.add_argument("--workers", type=int, default=8)
    ap.add_argument("--retry-failed", action="store_true",
                    help="Re-attempt rows whose checkpoint status is 'failed'")
    args = ap.parse_args()

    # Cached formula values ALWAYS from the original master file
    wsv = load_workbook(INPUT, data_only=True).active
    verify, recover = select_rows(wsv)
    print(f"selected: VERIFY={len(verify)}  RECOVER={len(recover)}")

    cp = load_checkpoint()
    done = cp["done"]

    def is_pending(r):
        d = done.get(str(r))
        if d is None:
            return True
        return args.retry_failed and d["status"] == "failed"

    verify = [r for r in verify if is_pending(r)]
    recover = [r for r in recover if is_pending(r)]
    if args.test:
        half = args.test // 2
        recover = recover[:args.test - half]
        verify = verify[:half]
    todo = [("verify", r) for r in verify] + [("recover", r) for r in recover]
    print(f"after checkpoint/test filter: VERIFY={len(verify)} RECOVER={len(recover)}")
    if not todo:
        print("Nothing to do.")
        return

    # Write into the output copy (formulas preserved); resume from it if present
    wbf = load_workbook(OUT if OUT.exists() else INPUT)
    wsf = wbf.active

    client = get_tinyfish_client()
    total_calls = 0
    counts = {"recovered": 0, "verified": 0, "mismatched": 0, "failed": 0}

    def ctx_for(r):
        return {"row": r,
                "person": sval(wsv, f"C{r}"), "title": sval(wsv, f"D{r}"),
                "location": sval(wsv, f"E{r}"), "url": sval(wsv, f"T{r}"),
                "company": sval(wsv, f"B{r}")}

    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {}
        for bucket, r in todo:
            futures[pool.submit(work_one, client, ctx_for(r))] = (bucket, r)
            time.sleep(0.2)

        done_n = 0
        for fut in as_completed(futures):
            bucket, r = futures[fut]
            data, calls, err, blocked = fut.result()
            total_calls += calls
            found = (data or {}).get("company") or ""
            person = sval(wsv, f"C{r}")

            if bucket == "recover":
                if found:
                    clear_needs_research(wsf, r)
                    wsf[f"B{r}"].value = found
                    new_title = (data or {}).get("title")
                    if new_title and not is_location_string(new_title):
                        wsf[f"D{r}"].value = new_title
                    if (data or {}).get("company_url"):
                        wsf[f"F{r}"].value = f"company page: {data['company_url']}"
                    status = "recovered"
                else:
                    append_note(wsf, r, "NEEDS_RESEARCH")
                    status = "failed"
            else:  # verify — never overwrite B
                existing = sval(wsv, f"B{r}")
                if found:
                    clear_needs_research(wsf, r)
                    if materially_same(existing, found):
                        status = "verified"
                    else:
                        append_note(wsf, r, f"MISMATCH: {found}")
                        status = "mismatched"
                elif blocked or (err and any(t in str(err).lower() for t in TRANSIENT)):
                    append_note(wsf, r, "NEEDS_RESEARCH")
                    status = "failed"
                else:
                    append_note(wsf, r, "MISMATCH: no employer visible on public profile")
                    status = "mismatched"

            counts[status] += 1
            done[str(r)] = {"bucket": bucket, "status": status,
                            "found": found, "calls": calls,
                            "error": str(err)[:80] if err else None}
            wbf.save(OUT)          # checkpoint: workbook after every row
            save_checkpoint(cp)
            done_n += 1
            print(f"[{done_n}/{len(todo)}] row {r} ({bucket}) {person}: "
                  f"{status.upper()}" + (f" -> {found}" if found else "") +
                  (f" [{calls} calls]"))

    print("\n===== PASS COMPLETE =====")
    print(f"counts this run: {counts}")
    print(f"TinyFish calls this run: {total_calls}")
    print(f"output: {OUT}")


if __name__ == "__main__":
    main()
