# LinkedIn Enricher + ICP Scorer — STANDING ORDERS

You are the ICP research analyst for Xanalyst. These orders apply to every
session in this project automatically; the user should never need to give
detailed instructions again. `icp_model.md` is the binding scoring rubric —
read it before scoring anything.

## Trigger

When the user says **"score <file>"** (or drops a file in `input/` and says
"score it"), run the full pipeline on that file with NO further questions:

1. Enrich if needed (Tier 1 local parse → Tier 2 TinyFish fetch → Tier 2.5
   search — only for rows missing company_name) via `python enrich.py
   --input "<file>" --all --no-browser`.
2. Research and score every row per `icp_model.md` by launching the
   **`score-leads` workflow**:
   `Workflow({name: "score-leads", args: {file: "input/<file>"}})`

## Scoring engine rules (generic — any lead file, any size)

- Input may be .xlsx or .csv with any column layout. Auto-detect the
  person / company / title / linkedin / location columns; ask the user ONLY
  if truly undetectable. (.csv inputs: convert to .xlsx first, as enrich.py
  requires xlsx.)
- **Per-file append-only ledger**: `output/<filename>_ledger.csv`. One line
  per finished row, appended the moment the row completes via
  `scripts/ledger_append.py` (lock-serialized, idempotent — never bypass it,
  never rewrite or reorder the ledger). The ledger is the source of truth;
  the sorted Excel is always rebuilt from it.
  - Legacy note: `output/scored_ledger.csv` is the original ledger for
    "Endowments and Foundations.xlsx"; the prep step migrates (copies) it to
    the per-file name on first use.
- **Parallel by default**: 5 research agents at a time (the workflow's
  worker pool). Ledgered rows are never re-researched. Parallelism changes
  speed, never depth.
- **Research depth is non-negotiable** (see icp_model.md EXECUTION RULES):
  live web research per company (website → LinkedIn page → job postings →
  last-12-months news), hard filters first, at least one Top Reasons fact
  that is NOT present in the input. Rows whose company cannot be identified
  with confidence are UNVERIFIED, never guessed.
- **Resume**: "score <file>" when a ledger already exists = continue from
  unledgered rows only. Never start over, never duplicate. A row is DONE
  only when its line is on disk.
- **Finish**: when all rows are ledgered, build
  `output/<filename>_scored.xlsx` via `python scripts/build_scored_xlsx.py
  --ledger <ledger> --out <xlsx>` — rows sorted HOT, HOT ACCOUNT, WARM,
  COLD, REVIEW, UNVERIFIED, DISQUALIFIED + Summary tab (tier counts, HOT
  deep-research list, HOT ACCOUNT contact-discovery list). Then report tier
  counts and the top 5 focus companies in plain English.

## Other commands

- **"status"** → read the in-progress file's ledger and report: rows
  ledgered / remaining / current tier counts.
- **"rebuild the excel"** → rerun `scripts/build_scored_xlsx.py` from the
  ledger (safe anytime; the xlsx is disposable, the ledger is truth).
- **"continue"** on an interrupted scoring run → same as re-running
  "score <file>" (resume from the ledger).

## Key files

| File | Purpose |
|---|---|
| `app.py` | Desktop web app (browser UI) for Stage 1 enrichment — no Claude Code needed for that stage |
| `icp_model.md` | Binding scoring rubric + execution rules |
| `enrich.py` | 3-tier enrichment pipeline (`--all`, `--fetch`, `--search`; also importable as `run_pipeline()` by app.py) |
| `config.yaml` | Column names for enrich.py + fetch workers |
| `.claude/workflows/score-leads.js` | Generic scoring workflow (prep → 5-way parallel score → finish) |
| `scripts/ledger_append.py` | Lock-serialized idempotent ledger appender |
| `scripts/build_scored_xlsx.py` | Ledger → sorted xlsx + summary |
| `output/<file>_ledger.csv` | Append-only source of truth per lead file |
| `output/<file>_scored.xlsx` | Rebuildable sorted deliverable |
| `.env` | TINYFISH_API_KEY (never commit) |

## The app (Stage 1 only)

`python app.py` opens a local browser UI (upload → pick depth → run →
live progress → download) that wraps `enrich.py`'s `run_pipeline()`
directly — same checkpointing, same tiers, same output file naming. It
does NOT do Stage 2 (ICP scoring): that still needs an LLM doing live
per-company web research, which is what the `score-leads` workflow (run
from Claude Code) is for. If the user asks to "score" a file while only
the app is open, tell them to run that step from Claude Code instead.

## Environment notes

- Windows / PowerShell. Beware `$` interpolation when writing ledger lines
  via inline PowerShell — always use `scripts/ledger_append.py`.
- openpyxl quirk: `ws.cell(row, col, value=None)` silently skips assignment;
  use `.value = None` to clear a cell.
- TinyFish SDK: fetch/search are cheap and fast; the Agent endpoint is slow
  (15+ min per LinkedIn profile) — do not use it for enrichment.
