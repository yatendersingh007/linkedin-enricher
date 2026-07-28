"""Serialized, idempotent ledger appender.

Usage: python scripts/ledger_append.py <line-file> --ledger <ledger.csv>

Appends the single pipe-separated line in <line-file> (UTF-8) to the given
append-only ledger. Safe under parallel callers: a lock directory serializes
writers, and the row_id duplicate check happens under the lock, so the same
row can never be appended twice even in a race. Prints "LEDGERED: <row_id>"
or "ALREADY: <row_id>".
"""
import argparse
import os
import sys
import time
from pathlib import Path

HEADER = ("row_id | Company | Contact | Title | HQ Location | Sector | "
          "Est. Size | Score | Tier | Top Reasons | Red Flags | "
          "Key Evidence Sources | Service Track\n")
DEFAULT_LEDGER = Path(__file__).resolve().parent.parent / "output" / "scored_ledger.csv"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("line_file")
    ap.add_argument("--ledger", default=str(DEFAULT_LEDGER))
    a = ap.parse_args()

    ledger = Path(a.ledger)
    line = Path(a.line_file).read_text(encoding="utf-8").strip().replace("\n", " ")
    if not line:
        sys.exit("empty line file — nothing appended")
    row_id = line.split(" | ")[0].strip()

    lock = Path(str(ledger) + ".lock")
    acquired = False
    for _ in range(400):  # up to ~20s
        try:
            os.mkdir(lock)
            acquired = True
            break
        except FileExistsError:
            time.sleep(0.05)
    if not acquired:
        sys.exit("could not acquire ledger lock")
    try:
        if not ledger.exists():
            ledger.parent.mkdir(exist_ok=True)
            ledger.write_text(HEADER, encoding="utf-8")
        existing_ids = {l.split(" | ")[0].strip()
                        for l in ledger.read_text(encoding="utf-8").splitlines()[1:] if l.strip()}
        if row_id in existing_ids:
            print(f"ALREADY: {row_id}")
        else:
            with open(ledger, "a", encoding="utf-8") as f:
                f.write(line + "\n")
                f.flush()
                os.fsync(f.fileno())
            print(f"LEDGERED: {row_id}")
    finally:
        os.rmdir(lock)


if __name__ == "__main__":
    main()
