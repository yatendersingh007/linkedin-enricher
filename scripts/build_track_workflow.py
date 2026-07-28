"""Extract HOT/WARM rows from the scored allocators xlsx and generate the
verify-service-track workflow script with the queue embedded (args proved
unreliable for large payloads). Also pre-creates the service-track ledger
with its own header so the shared appender never writes the scoring header
into it."""
import json
import sys
from pathlib import Path

from openpyxl import load_workbook

PROJ = Path(__file__).resolve().parent.parent
XLSX = PROJ / "output" / "Chanmeet_Allocators_MASTER_Scored.xlsx"
LEDGER = PROJ / "output" / "service_track_ledger.csv"
WF = PROJ / ".claude" / "workflows" / "verify-service-track.js"

ws = load_workbook(XLSX).active


import re
CTRL_RX = re.compile(r"[\x00-\x1f\x7f-\x9f]")


def s(coord):
    v = ws[coord].value
    if v is None:
        return None
    text = CTRL_RX.sub("", str(v)).strip()
    return text or None


queue = []
for r in range(2, ws.max_row + 1):
    temp = (s(f"R{r}") or "").upper()
    if temp not in ("HOT", "WARM"):
        continue
    queue.append({
        "row_id": r,
        "temp": temp,
        "company": s(f"B{r}"),
        "person": s(f"C{r}"),
        "title": s(f"D{r}"),
        "location": s(f"E{r}"),
        "linkedin_url": s(f"V{r}"),   # after the S/T insertion, URL lives in V
        # keyword_track kept OUT of the agent prompt to avoid anchoring;
        # stored separately for the changed-vs-keyword comparison later
    })
print(f"queue: {len(queue)} rows "
      f"(HOT={sum(1 for q in queue if q['temp']=='HOT')}, "
      f"WARM={sum(1 for q in queue if q['temp']=='WARM')})")

if not LEDGER.exists():
    LEDGER.write_text(
        "row_id | Company | Verified Track | Evidence | Track Source\n",
        encoding="utf-8")
    print("created service_track_ledger.csv")

TEMPLATE = r"""export const meta = {
  name: 'verify-service-track',
  description: 'Research-based Service Track verification for HOT/WARM allocator rows: light web research per row, append-only ledger, 5 parallel workers',
  whenToUse: 'When the user asks to verify/replace keyword service tracks with researched ones',
  phases: [
    { title: 'Prep', detail: 'read ledger, filter done rows' },
    { title: 'Classify', detail: 'light research per row, 5 workers' },
  ],
}

const PROJ = 'c:/Users/DELL/Documents/linkedin-enricher/linkedin-enricher'
const LEDGER = PROJ + '/output/service_track_ledger.csv'

const A = (() => {
  if (typeof args === 'string') { try { return JSON.parse(args) } catch (e) { return {} } }
  return args || {}
})()
const LIMIT = A.limit || 0
const WORKERS = A.workers || 5

const QUEUE = __QUEUE_JSON__

const PREP_SCHEMA = {
  type: 'object',
  properties: { done_ids: { type: 'array', items: { type: 'integer' } } },
  required: ['done_ids'],
}
const ROW_SCHEMA = {
  type: 'object',
  properties: {
    row_id: { type: 'integer' },
    company: { type: 'string' },
    track: { type: 'string' },
    evidence: { type: 'string' },
    searches: { type: 'integer' },
    appended: { type: 'string', enum: ['true', 'already', 'failed'] },
  },
  required: ['row_id', 'company', 'track', 'evidence', 'searches', 'appended'],
}

function prompt(row) {
  return `You are verifying the SERVICE TRACK for one lead of Xanalyst (embeds India-based finance/accounting professionals into US/Canada real estate and private equity firms). LIGHT, focused research only: 2-4 web searches answering ONE question - what does this person's CURRENT company actually operate?

LEAD (row_id ${row.row_id}):
- person: ${row.person}
- company on sheet (may be STALE or a former employer - e.g. "Ex GE Capital" is not a current employer): ${row.company}
- title on sheet: ${row.title}
- location: ${row.location}
- linkedin: ${row.linkedin_url}

STEP 1 - verify the CURRENT company. Search the person (name + company / LinkedIn slug). If the sheet company is wrong or stale, use the real current firm. If you cannot confidently identify the current firm, track = UNVERIFIED (evidence = what is ambiguous) - NEVER guess.

STEP 2 - classify what the firm operates (from research findings, never from the sheet's words):
- RE = operates/owns/manages property directly (owner-operator, property mgmt, developer, direct CRE)
- FUND ADMIN-GP = runs its OWN private funds (PE / private credit / VC / RE funds): capital calls, NAV, investor reporting
- FUND ADMIN-LP = allocator, wealth manager, private bank, or family office investing in OTHERS' funds, with no direct property
- BOTH (primary: X) = genuine direct-RE AND own-fund workload; tie goes to RE
Note: a "multifamily office" is a wealth office (LP side), NOT multifamily real estate.

Use WebSearch (and WebFetch when a specific page settles it); load via ToolSearch if needed. Count your searches.

STEP 3 - compose ONE ledger line (pipes ONLY as delimiters, use '/' inside cells, no newlines):
${row.row_id} | <Current Company> | <track> | <1-2 SPECIFIC researched facts that decided it - never a keyword echo> | research

STEP 4 - append it: Write the line to ${PROJ}/output/tmp_track_${row.row_id}.txt (UTF-8), run:
python "${PROJ}/scripts/ledger_append.py" "${PROJ}/output/tmp_track_${row.row_id}.txt" --ledger "${LEDGER}"
Expect "LEDGERED: ${row.row_id}" or "ALREADY: ${row.row_id}" (both = success), then delete the tmp file.

Return structured output: row_id, company (current, as verified), track, evidence, searches (int), appended.`
}

phase('Prep')
const prep = await agent(`Read ${LEDGER} (pipe-separated, first field of each data line is row_id). Return structured output: done_ids = the list of row_id integers already present (empty list if only the header exists).`,
  { label: 'read track ledger', schema: PREP_SCHEMA, phase: 'Prep' })
const done = new Set((prep && prep.done_ids) || [])
let todo = QUEUE.filter(q => !done.has(q.row_id))
if (LIMIT) todo = todo.slice(0, LIMIT)
log(`queue ${QUEUE.length}, already done ${done.size}, this run ${todo.length}`)

phase('Classify')
const results = []
let failures = 0
let next = 0
async function worker() {
  while (next < todo.length) {
    const row = todo[next++]
    const label = `row ${row.row_id}: ${row.company || row.person}`
    let r = await agent(prompt(row), { label, schema: ROW_SCHEMA, phase: 'Classify', model: 'sonnet' })
    if (!r) r = await agent(prompt(row), { label: label + ' (retry)', schema: ROW_SCHEMA, phase: 'Classify', model: 'sonnet' })
    if (!r) { failures++; log(`row ${row.row_id}: FAILED twice`); continue }
    if (r.appended === 'failed' && r.evidence) {
      await agent(`Write this line to ${PROJ}/output/tmp_track_${r.row_id}.txt then run python "${PROJ}/scripts/ledger_append.py" "${PROJ}/output/tmp_track_${r.row_id}.txt" --ledger "${LEDGER}" and confirm LEDGERED/ALREADY, then delete the tmp file. Line:\n${r.row_id} | ${r.company} | ${r.track} | ${r.evidence} | research`,
        { label: `scribe ${r.row_id}`, phase: 'Classify' })
    }
    results.push(r)
    log(`row ${r.row_id}: ${r.track} (${r.searches} searches) - ${r.company}`)
  }
}
if (todo.length) await Promise.all(Array.from({ length: Math.min(WORKERS, todo.length) }, () => worker()))
return { attempted: todo.length, completed: results.length, failures }
"""

WF.write_text(TEMPLATE.replace("__QUEUE_JSON__",
                               json.dumps(queue, ensure_ascii=True)),
              encoding="utf-8", newline="\n")
print(f"workflow written: {WF} ({WF.stat().st_size // 1024} KB)")

# stash keyword tracks for the later changed-vs-keyword comparison
snap = {str(q["row_id"]): (s(f"S{q['row_id']}") or "") for q in queue}
(PROJ / "output" / "keyword_track_snapshot.json").write_text(
    json.dumps(snap, ensure_ascii=False), encoding="utf-8")
print("keyword snapshot saved")
