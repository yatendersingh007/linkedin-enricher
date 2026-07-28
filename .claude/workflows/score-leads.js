export const meta = {
  name: 'score-leads',
  description: 'Generic ICP scoring pipeline: enrich if needed, then research + score every unledgered row per icp_model.md with 5 parallel agents and a per-file append-only ledger',
  whenToUse: 'When the user says "score <file>" or "score it" for a lead spreadsheet',
  phases: [
    { title: 'Prep', detail: 'detect columns, enrich if needed, build queue of unledgered rows' },
    { title: 'Score', detail: '5 parallel research agents, serialized ledger appends' },
    { title: 'Finish', detail: 'rebuild sorted Excel + summary from the ledger' },
  ],
}

const PROJ = 'c:/Users/DELL/Documents/linkedin-enricher/linkedin-enricher'

// args may arrive as an object, a JSON string, or a bare file path string
const A = (() => {
  if (typeof args === 'string') {
    try { return JSON.parse(args) } catch (e) { return { file: args.trim() } }
  }
  return args || {}
})()
if (!A.file) throw new Error('args.file required, e.g. {"file": "input/leads.xlsx"}')
const WORKERS = A.workers || 5

const PREP_SCHEMA = {
  type: 'object',
  properties: {
    ledger_path: { type: 'string' },
    source_file: { type: 'string' },
    total_rows: { type: 'integer' },
    done_count: { type: 'integer' },
    notes: { type: 'string' },
    queue: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          row_id: { type: 'integer' },
          person_name: { type: ['string', 'null'] },
          person_title: { type: ['string', 'null'] },
          company_name: { type: ['string', 'null'] },
          linkedin_url: { type: ['string', 'null'] },
          hq_location: { type: ['string', 'null'] },
          raw_headline: { type: ['string', 'null'] },
        },
        required: ['row_id'],
      },
    },
  },
  required: ['ledger_path', 'source_file', 'total_rows', 'done_count', 'queue'],
}

const ROW_SCHEMA = {
  type: 'object',
  properties: {
    row_id: { type: 'integer' },
    company: { type: 'string' },
    tier: { type: 'string' },
    score: { type: 'integer' },
    ledger_line: { type: 'string' },
    appended: { type: 'string', enum: ['true', 'already', 'failed'] },
  },
  required: ['row_id', 'company', 'tier', 'score', 'ledger_line', 'appended'],
}

const FIN_SCHEMA = {
  type: 'object',
  properties: {
    total: { type: 'integer' },
    counts: { type: 'object' },
    xlsx: { type: 'string' },
  },
  required: ['total', 'counts', 'xlsx'],
}

function scorePrompt(row, ledgerPath) {
  return `You are an ICP research analyst for Xanalyst scoring ONE lead. Be rigorous; a wrong score poisons outreach.

STEP 0 (mandatory): Read ${PROJ}/icp_model.md IN FULL. It is your binding rubric and operating manual. Its EXECUTION RULES override everything: hard filters first, live web research mandatory, no proxy scores from input text, at least one Top Reasons fact must come from research (not from the row below).

YOUR LEAD (spreadsheet row_id ${row.row_id}):
- person_name: ${row.person_name}
- person_title: ${row.person_title}
- company_name: ${row.company_name}
- linkedin_url: ${row.linkedin_url}
- hq_location: ${row.hq_location}
- raw LinkedIn headline: ${row.raw_headline}

${row.company_name ? "The company_name came from automated enrichment - verify it is really this person's current company while researching." : 'company_name is MISSING: first identify the person\'s current company via web search (their LinkedIn URL slug, name + title + location, news). If you cannot identify it with confidence, Tier = UNVERIFIED (state what is ambiguous) - never guess.'}

RESEARCH (live, mandatory - use WebSearch and WebFetch; load them via ToolSearch if not yet available): (1) company website, (2) LinkedIn company page for headcount and finance-team titles, (3) open job postings (careers page, LinkedIn Jobs, Indeed), (4) news from the last 12 months. Typically 3-8 searches. Apply the rubric: hard filters (geography US/Canada, allowed sector, not a competitor) THEN sector fit /30, F&A workload evidence /25, size /15 (family-office and allocator exceptions!), active need signals /15, designation fit /10 (score the person above as the contact; if the row is effectively company-only, find the best finance contact or use 'no contact identified' = 3), geography /5, minus red flags. Tier per the model: HOT needs 70+ AND designation 8 or 10; 70+ otherwise = HOT ACCOUNT; UK/Western Europe 70+ = REVIEW; WARM 45-69; COLD 1-44; DISQUALIFIED 0; UNVERIFIED if company unidentifiable.

COMPOSE the ledger line - EXACTLY one line, pipe-separated, this order:
${row.row_id} | Company | Contact | Title | HQ Location | Sector | Est. Size | Score | Tier | Top Reasons | Red Flags | Key Evidence Sources | Service Track
Rules: never use the pipe character inside a cell (use '/'); no newlines; Top Reasons = 2-3 SPECIFIC facts with at least one from research; Red Flags = from the model's list or 'none found'; Sources = 1-2 URLs; Service Track per the SERVICE TRACK section of icp_model.md (RE / FUND ADMIN-GP / FUND ADMIN-LP / BOTH (primary: X), from researched evidence).

APPEND (mandatory final step): Write your single line to ${PROJ}/output/tmp_row_${row.row_id}.txt (UTF-8), then run:
python "${PROJ}/scripts/ledger_append.py" "${PROJ}/output/tmp_row_${row.row_id}.txt" --ledger "${ledgerPath}"
It must print LEDGERED: ${row.row_id} (or ALREADY: ${row.row_id} - both count as success; the appender is lock-serialized and idempotent). Then delete the tmp file. Set appended="true" (or "already"). If the command errors twice, set appended="failed" but still return your ledger_line.

Return structured output: row_id, company (what you scored), tier, score, ledger_line, appended.`
}

// ---------------------------------------------------------------- Prep
phase('Prep')
const prep = await agent(`You are preparing a lead file for ICP scoring in the project at ${PROJ}.

TARGET FILE (relative to project unless absolute): ${A.file}

Do these steps precisely:
1. Locate the file. Let STEM = its basename without extension.
2. LEDGER: the ledger for this file is ${PROJ}/output/<STEM>_ledger.csv. If it does not exist but ${PROJ}/output/scored_ledger.csv exists AND clearly belongs to this same lead file (spot-check a few company/person names), COPY scored_ledger.csv to <STEM>_ledger.csv (do not delete the original). If neither exists, the ledger starts empty (do not create it; the appender will).
3. SOURCE ROWS: if ${PROJ}/output/enriched_<STEM>.xlsx exists, use it as the row source. Otherwise inspect the raw file (.xlsx via openpyxl, .csv via csv module): auto-detect which columns hold person name, person title, company name, LinkedIn profile URL, and location (case-insensitive header match plus sniffing values like linkedin.com). If a large share of rows lack a company but have LinkedIn URLs, first run: python enrich.py --input "<file>" --all --no-browser (update config.yaml column names first if this file's headers differ), then use its enriched output as the row source.
4. QUEUE: row_id = the spreadsheet row number in the SOURCE file (header = row 1, so data starts at row_id 2). Read the ledger (if present); a row is DONE if its row_id appears as the first pipe-separated field of any ledger line. Build the queue of NOT-done rows in ascending row order with fields: row_id, person_name, person_title, company_name, linkedin_url, hq_location, raw_headline (the original raw position/headline text if such a column exists, else null).
5. Return structured output: ledger_path (absolute), source_file (absolute), total_rows (data rows in source), done_count, queue, notes (anything noteworthy, e.g. 'migrated legacy ledger', 'ran enrichment').

Use python scripts via Bash for reading the spreadsheet (openpyxl is installed). Do NOT modify the source file or the ledger (except the one legacy COPY in step 2).`,
  { label: 'prep: queue + ledger', schema: PREP_SCHEMA, phase: 'Prep' })

if (!prep) throw new Error('prep agent failed')
log(`${prep.notes || ''} ledger=${prep.ledger_path} total=${prep.total_rows} done=${prep.done_count} queue=${prep.queue.length}`)

// ---------------------------------------------------------------- Score
phase('Score')
const queue = prep.queue
const results = []
let failures = 0
let next = 0

async function workerLoop(wid) {
  while (next < queue.length) {
    const row = queue[next++]
    const label = `row ${row.row_id}: ${row.company_name || row.person_name}`
    let r = await agent(scorePrompt(row, prep.ledger_path), { label, schema: ROW_SCHEMA, phase: 'Score' })
    if (!r) {
      log(`row ${row.row_id}: agent died, retrying once`)
      r = await agent(scorePrompt(row, prep.ledger_path), { label: label + ' (retry)', schema: ROW_SCHEMA, phase: 'Score' })
    }
    if (!r) {
      failures++
      log(`row ${row.row_id}: FAILED twice - needs manual reconciliation`)
      results.push({ row_id: row.row_id, company: row.company_name || row.person_name, tier: 'AGENT_FAILED', score: -1, ledger_line: '', appended: 'failed' })
      continue
    }
    if (r.appended === 'failed' && r.ledger_line) {
      await agent(`Run this from ${PROJ}: Write the line below to ${PROJ}/output/tmp_row_${r.row_id}.txt then execute python "${PROJ}/scripts/ledger_append.py" "${PROJ}/output/tmp_row_${r.row_id}.txt" --ledger "${prep.ledger_path}" and confirm it prints LEDGERED or ALREADY, then delete the tmp file. The line:\n${r.ledger_line}\nReturn "done" or the error.`,
        { label: `scribe row ${r.row_id}`, phase: 'Score' })
    }
    results.push(r)
    log(`row ${r.row_id} ledgered: ${r.tier} (${r.score}) - ${r.company}`)
  }
}

if (queue.length) {
  await Promise.all(Array.from({ length: Math.min(WORKERS, queue.length) }, (_, i) => workerLoop(i)))
} else {
  log('Queue empty - all rows already ledgered.')
}
log(`Scoring done: ${results.length} rows this run, ${failures} hard failures.`)

// ---------------------------------------------------------------- Finish
phase('Finish')
const stem = prep.ledger_path.split(/[\\/]/).pop().replace(/_ledger\.csv$/i, '')
const fin = await agent(`Run this command from ${PROJ} and return its JSON output as structured output (total, counts, xlsx):
python "${PROJ}/scripts/build_scored_xlsx.py" --ledger "${prep.ledger_path}" --out "${PROJ}/output/${stem}_scored.xlsx"
If it errors, fix nothing - just report the error text in place of counts (as {"error": "..."}) with total=0 and xlsx="".`,
  { label: 'build sorted excel', schema: FIN_SCHEMA, phase: 'Finish' })

return { ledger: prep.ledger_path, total_rows: prep.total_rows, scored_this_run: results.length, failures, finish: fin }
