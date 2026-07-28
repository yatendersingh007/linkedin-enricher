# LinkedIn Enricher + ICP Scorer — Xanalyst lead pipeline

Turns a raw event-attendee spreadsheet into a scored, prioritized lead
list. Two stages:

1. **Enrichment** — fills in each person's company, title, and location
   from their LinkedIn profile, using a free local parse first and
   TinyFish (fetch + search) only for the gaps.
2. **ICP scoring** — researches each company live and scores it against
   Xanalyst's rubric (`icp_model.md`), producing a sorted Excel with
   HOT/WARM/COLD tiers and a Service Track (RE / Fund Admin-GP / Fund
   Admin-LP) for outreach routing.

## Running it — desktop app (recommended)

```
python app.py
```

It picks the first free port starting at 5000 and opens that exact URL in
your browser — **the console prints the real link** (e.g.
`http://127.0.0.1:5001` if 5000 is busy), so always use the URL it prints
rather than assuming 5000. Drop your file, pick a depth (free parse only /
+ fetch / + full search), click Run, watch the live progress bar, and
download the enriched spreadsheet when it's done. Covers Stage 1
(enrichment) end-to-end without touching a terminal or Claude Code.

**Any column layout works.** The app auto-detects the name / title /
LinkedIn-URL / location columns from common header spellings (and by
sniffing for linkedin.com links), so files like `Person Name` + `Title` +
`LinkedIn URL` enrich without editing `config.yaml`. If it can't find a
usable column it tells you which headers it saw, instead of failing
silently.

**Stage 2 (ICP scoring)** still runs through Claude Code's `score-leads`
workflow — it needs an LLM doing live web research per company, which is
outside what a plain Flask app can drive on its own. In Claude Code, just
say **"score `<file>`"** and the standing orders in `CLAUDE.md` take it
from there (see that file for the full behavior contract).

## Running it — command line

```
python enrich.py --input "input/yourfile.xlsx" --all
```

- Free parse only: drop `--all`
- Add TinyFish fetch: `--fetch`
- Full pipeline (parse → fetch → retry → search → retry): `--all`
- Resume a partial run: just re-run the same command — the enriched
  output file **is** the checkpoint, finished rows are never redone
- `--reset` retries rows that ended up `NEEDS_RESEARCH`
- `--workers N` to change TinyFish parallelism (default from `config.yaml`)

## Setup

```
pip install -r requirements.txt
cp .env.example .env    # then paste your TinyFish key into .env
```

Set the column names for your sheet in `config.yaml` if they differ from
the defaults (`Linkedin`, `Positions/Company`, `Location`, `Name`).

## Output columns (enrichment)

| Column | Meaning |
|---|---|
| person_name | Name (from sheet or recovered) |
| person_title | Current job title |
| company_name | Current company — this is what Stage 2 needs |
| company_url | LinkedIn company page, when found |
| linkedin_url | The profile URL used |
| hq_location | Location |
| TF Status | PARSED / FETCHED / SEARCHED / NEEDS_RESEARCH / SKIPPED |
| TF Steps | TinyFish calls spent on that row |

## Cost control

TinyFish fetch/search are cheap (a few cents per call); the old Agent
endpoint was not (15+ min and real cost per profile) and is deliberately
unused. Still, test on a subset first with `--limit N` before running a
1000-row file.

## Key files

| File | Purpose |
|---|---|
| `app.py` | Desktop web app for enrichment (Stage 1) |
| `enrich.py` | The enrichment pipeline itself (CLI + importable `run_pipeline`) |
| `CLAUDE.md` | Standing orders for the ICP scoring stage (Stage 2) |
| `icp_model.md` | The binding scoring rubric |
| `config.yaml` | Column names + fetch worker count |
| `scripts/` | Ledger, xlsx-building, and one-off analysis helpers |
