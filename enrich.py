"""
LinkedIn Enricher — Stage 1 of the ICP pipeline (one-command, 3.5-tier).

TIER 1 (free): parse the sheet's "Positions/Company" text locally into
person_title + company_name. Clean company -> PARSED, no API call.

TIER 2 (--fetch): rows missing a company but having a LinkedIn URL go
through TinyFish Fetch (one page render, NOT the slow Agent endpoint),
in parallel workers. Fields extracted from the markdown locally -> FETCHED.

TIER 2.5 (--search): rows Fetch couldn't resolve get one TinyFish Search
pass: query built from the person's name + a hint. A company is accepted
only if the result URL matches the person's LinkedIn slug, OR the result
title contains their name AND the snippet overlaps something we already
know about them. Wrong company is worse than missing -> abstain. SEARCHED.

--all runs Tier 1 -> Tier 2 -> one low-concurrency retry of transient
fetch errors -> Tier 2.5 -> one low-concurrency retry of transient search
errors. Authwall / page_not_found rows are never retried. Whatever
survives stays NEEDS_RESEARCH for the next (name-research) stage.

The output workbook IS the checkpoint: statuses are saved after every
row, so re-running the same command skips finished rows automatically.

Progress: live console display (rich) + a browser dashboard.
output/progress.json is rewritten after every row; output/progress.html
(auto-opened, self-contained) polls it every 2 seconds.

Usage:
    python enrich.py --input "input/file.xlsx" --all
    python enrich.py --input "input/file.xlsx"            # Tier 1 only
    python enrich.py --input "input/file.xlsx" --fetch    # Tier 2 only
    python enrich.py --input "input/file.xlsx" --search   # Tier 2.5 only
    extra: --limit N  --workers N  --reset  --no-browser  --samples N
"""

import argparse
import json
import os
import re
import sys
import threading
import time
import webbrowser
from collections import deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

import yaml
from dotenv import load_dotenv
from openpyxl import load_workbook

load_dotenv()

CONFIG_PATH = Path(__file__).parent / "config.yaml"
OUTPUT_DIR = Path(__file__).parent / "output"

STATUS_PARSED = "PARSED"
STATUS_PENDING = "PENDING_FETCH"
STATUS_FETCHED = "FETCHED"
STATUS_SEARCHED = "SEARCHED"
STATUS_NEEDS_RESEARCH = "NEEDS_RESEARCH"
STATUS_SKIPPED = "SKIPPED"
TERMINAL_STATUSES = {STATUS_FETCHED, STATUS_SEARCHED, STATUS_NEEDS_RESEARCH}
STATUS_ORDER = [STATUS_PARSED, STATUS_PENDING, STATUS_FETCHED, STATUS_SEARCHED,
                STATUS_NEEDS_RESEARCH, STATUS_SKIPPED, "EMPTY"]

TRANSIENT_MARKERS = ("service unavailable", "timeout", "timed out",
                     "rate limited", "unreachable")

NEW_COLUMNS = ["person_name", "person_title", "company_name", "company_url",
               "linkedin_url", "hq_location", "TF Status", "TF Steps"]

COVERAGE_FIELDS = ["person_name", "person_title", "company_name",
                   "company_url", "linkedin_url", "hq_location"]


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


# ---------------------------------------------------------------------------
# Tier 1 — local parsing of "Positions/Company"
# ---------------------------------------------------------------------------

TITLE_RX = re.compile(
    r"\b(ceo|cio|cfo|coo|cgo|cto|cmo|evp|svp|avp|vp|vice president|founder|"
    r"co[- ]?founder|president|chairman|chairwoman|chair|director|manager|"
    r"managing|partner|principal|head|officer|advisor|adviser|strategist|"
    r"consultant|executive|dean|professor|analyst|associate|lead|trustee|"
    r"board member|senator|owner|investor|banker|philanthropist)\b", re.I)

COMPANY_RX = re.compile(
    r"\b(llc|inc|ltd|llp|lp|corp|corporation|company|co\.|capital|management|"
    r"partners|group|holdings|foundation|fund|funds|ventures|advisors|"
    r"associates|university|school|college|federation|bank|trust|institute|"
    r"center|centre|forum|firm|wealth|solutions|accelerator)\b", re.I)

# Phrases that describe a market segment, not an actual organization.
GENERIC_COMPANIES = {
    "family office", "family offices", "single family office",
    "endowments & foundations", "endowments and foundations",
    "foundations & endowments", "foundations and endowments",
    "philanthropy", "finance", "institutional", "private equity",
    "wealth management", "growth", "fundraising & partnerships",
    "board member", "family-owned businesses & family offices",
    "family office, truffle pig, philanthropy, travel guru",
    "private foundation finance",
}


def clean_company(candidate):
    """Validate a company-name candidate; return it cleaned or None."""
    if not candidate:
        return None
    c = re.sub(r"[\U0001F000-\U0001FAFF☀-➿⬀-⯿️]", "",
               str(candidate)).strip().strip(".,;:").strip()
    if not c or len(c) < 2 or len(c) > 70 or not re.search(r"[A-Za-zÀ-ÿ]", c):
        return None
    if "..." in c or "…" in c:
        return None
    # Two or more commas means a tagline/list, not an organization name
    if c.count(",") >= 2:
        return None
    if c.lower() in GENERIC_COMPANIES:
        return None
    # A candidate that is nothing but title words is not a company
    if TITLE_RX.search(c) and not COMPANY_RX.search(c):
        return None
    return c


def clean_title(candidate):
    if not candidate:
        return None
    t = str(candidate).strip().strip(".,;:").strip()
    if len(t) < 3 or len(t) > 90 or "[" in t or "]" in t:
        return None
    return t


def parse_positions(raw):
    """Split a raw 'Positions/Company' cell into (title, company).

    Handles separators ' at ', ' @ ', ' - ', ',', ' | ' and both
    'title SEP company' and 'company SEP title' orderings.
    Either element may come back None."""
    if not raw or not str(raw).strip():
        return None, None
    text = str(raw).strip()
    # Some headlines use a decorative " I " instead of " | " — treat it as a
    # separator only when it appears repeatedly (a styling pattern).
    sep = r"\s*\|\s*"
    if len(re.findall(r"\s+I\s+(?=[A-Z])", text)) >= 2:
        sep = r"\s*\|\s*|\s+I\s+(?=[A-Z])"
    segments = [s.strip() for s in re.split(sep, text) if s.strip()]
    if not segments:
        return None, None
    title, company = _parse_segment(segments[0])
    # Primary segment may hold only one of the two — scan the rest for the
    # other, demanding stronger evidence (COMPANY_RX) from secondary segments.
    for seg in segments[1:]:
        if title and company:
            break
        t, c = _parse_segment(seg)
        if not title and t and not c:
            title = t
        if not company and c and COMPANY_RX.search(str(c)):
            company = c
            if not title and t:
                title = t
    return clean_title(title), clean_company(company)


def _parse_segment(seg):
    """Parse one pipe-free segment into (title, company) candidates."""
    # 1) "Founder @ Astra Filantropía"
    parts = re.split(r"\s+@\s*", seg)
    if len(parts) >= 2:
        return parts[0], parts[-1]
    # 2) "Investment Manager at The Weinberg Foundation" (last ' at ' wins)
    parts = re.split(r"\s+at\s+", seg, flags=re.I)
    if len(parts) >= 2:
        return " at ".join(parts[:-1]), parts[-1]
    # 3) "Managing Director - Amchor" / "US Equestrian Federation - President"
    parts = re.split(r"\s+[-–—]\s+", seg)
    if len(parts) >= 2:
        left, right = parts[0], " - ".join(parts[1:])
        left_titley = bool(TITLE_RX.search(left))
        right_titley = bool(TITLE_RX.search(right))
        if right_titley and not left_titley:
            # reversed order "Company - Title" needs real company evidence,
            # else "Politics and Democracy - Donor Advisor" invents a company
            return right, (left if COMPANY_RX.search(left) else None)
        return left, right
    # 4) "Chairman & CEO, Bolt-Hole Capital Management,LLC"
    m = re.match(r"^(?P<t>[^,]{2,60}?),\s*(?P<c>.+)$", seg)
    if m and TITLE_RX.search(m.group("t")) and not TITLE_RX.search(m.group("c")) \
            and COMPANY_RX.search(m.group("c")):
        return m.group("t"), m.group("c")
    # 4b) "CEO, Philanthropist, Elizabeth Elting Foundation" — company after
    # the LAST comma when it reads like a proper multi-word organization
    left, _, right = seg.rpartition(",")
    right = right.strip()
    if left and COMPANY_RX.search(right) and not TITLE_RX.search(right) \
            and len(right.split()) >= 2:
        return left.strip(), right
    # 5) no separator: decide if the whole thing is a title or a company
    if COMPANY_RX.search(seg) and not TITLE_RX.search(seg):
        return None, seg
    return seg, None


# ---------------------------------------------------------------------------
# TinyFish client + Tier 2 fetch
# ---------------------------------------------------------------------------

def get_tinyfish_client():
    api_key = os.getenv("TINYFISH_API_KEY")
    if not api_key or "paste-your" in api_key:
        sys.exit(
            "ERROR: TINYFISH_API_KEY not found. Copy .env.example to .env "
            "and paste your TinyFish API key into it."
        )
    try:
        from tinyfish import TinyFish
    except ImportError:
        sys.exit("ERROR: tinyfish SDK not installed. Run: pip install -r requirements.txt")
    return TinyFish(api_key=api_key)


def fetch_one(client, url, max_attempts=3):
    """Fetch one profile page. Returns (FetchResult|None, attempts, error|None).

    Respects Retry-After on 429s, exponential backoff otherwise."""
    from tinyfish import RateLimitError

    attempts = 0
    while True:
        attempts += 1
        try:
            resp = client.fetch.get_contents([url], links=True)
            results = resp.results or []
            if results:
                return results[0], attempts, None
            err = (resp.errors or [None])[0]
            msg = getattr(err, "message", None) or getattr(err, "error", None) \
                or "no content returned"
            return None, attempts, str(msg)
        except RateLimitError as e:
            if attempts >= max_attempts:
                return None, attempts, f"rate limited: {e}"
            wait = getattr(e, "retry_after", None)
            time.sleep(float(wait) if wait else 2 ** attempts)
        except Exception as e:  # noqa: BLE001 - mark & move on, never stop the run
            if attempts >= max_attempts:
                return None, attempts, str(e)
            time.sleep(2 ** attempts)


LINKEDIN_TITLE_RX = re.compile(
    r"^(?P<name>.+?)\s*[-–]\s*(?P<headline>.+?)\s*\|\s*LinkedIn\s*$", re.S)
COMPANY_URL_RX = re.compile(
    r"https?://[a-z0-9.]*linkedin\.com/company/[A-Za-z0-9\-_%.]+", re.I)
# "City, State, Country" or "City, Country" — guards against writing a
# location string into a title field when parsing goes wrong upstream
LOCATION_STRING_RX = re.compile(
    r"^[A-Za-z .'-]+,\s*[A-Za-z .'-]+,\s*(United States|Canada)$|"
    r"^[A-Za-z .'-]+,\s*(United States|Canada)$", re.I)


def is_location_string(text):
    return bool(text) and bool(LOCATION_STRING_RX.match(str(text).strip()))


def _is_authwall(final_url, page_title, text):
    if any(k in final_url for k in ("authwall", "/login", "/signup",
                                    "/checkpoint", "/uas/")):
        return True
    tl = page_title.lower()
    if any(k in tl for k in ("sign in", "sign up", "join linkedin", "log in")):
        return True
    return len(text.strip()) < 150


def extract_from_fetch(fr):
    """Pull full_name / title / company / location / company_url out of a
    fetched LinkedIn page. Returns dict or None if the page is unusable."""
    page_title = (fr.title or "").strip()
    desc = (getattr(fr, "description", None) or "").strip()
    text = fr.text or ""
    final_url = (fr.final_url or fr.url or "").lower()

    if _is_authwall(final_url, page_title, text):
        return None

    out = {"full_name": None, "title": None, "company": None,
           "location": None, "company_url": None}

    # <title> tag: "Name - Headline | LinkedIn"
    m = LINKEDIN_TITLE_RX.match(page_title)
    if m:
        out["full_name"] = m.group("name").strip()
        t, c = parse_positions(m.group("headline"))
        # guard: never let a "City, State, Country" string end up as a title
        out["title"] = t if not is_location_string(t) else None
        out["company"] = c

    # meta description: "... · Experience: Company · ... · Location: City ..."
    if desc:
        em = re.search(r"Experience:\s*([^·|]+)", desc)
        if em and not out["company"]:
            out["company"] = clean_company(em.group(1))
        lm = re.search(r"Location:\s*([^·|]+)", desc)
        if lm:
            out["location"] = lm.group(1).strip()
        if not out["title"]:
            # guest-page descriptions sometimes open with the location, not
            # the headline — only trust this chunk if it reads like a title
            t, c = parse_positions(desc.split("·")[0])
            if t and TITLE_RX.search(t) and not is_location_string(t):
                out["title"] = t
            if c and not out["company"]:
                out["company"] = c

    # page body: first entry under an "Experience" heading
    if not out["company"] and text:
        xm = re.search(r"(?im)^#{0,6}\s*Experience\s*$\n+\s*\[?([^\n\]]{2,70})", text)
        if xm:
            out["company"] = clean_company(xm.group(1))

    # first linkedin.com/company/... link anywhere in links or body
    blob = " ".join(str(l) for l in (getattr(fr, "links", None) or [])) + " " + text
    cm = COMPANY_URL_RX.search(blob)
    if cm:
        out["company_url"] = cm.group(0).split("?")[0]

    return out if any(out.values()) else None


# ---------------------------------------------------------------------------
# Tier 2.5 — search snippet recovery
# ---------------------------------------------------------------------------

SERP_NOISE_TOKENS = {"linkedin", "profile", "view", "professional",
                     "community", "connections", "experience", "education",
                     "location", "members", "billion"}


def search_one(client, query, max_attempts=3):
    """Run one TinyFish search. Returns (results|None, attempts, error|None)."""
    from tinyfish import RateLimitError

    attempts = 0
    while True:
        attempts += 1
        try:
            resp = client.search.query(query)
            return resp.results or [], attempts, None
        except RateLimitError as e:
            if attempts >= max_attempts:
                return None, attempts, f"rate limited: {e}"
            wait = getattr(e, "retry_after", None)
            time.sleep(float(wait) if wait else 2 ** attempts)
        except Exception as e:  # noqa: BLE001
            if attempts >= max_attempts:
                return None, attempts, str(e)
            time.sleep(2 ** attempts)


def _profile_slug(url):
    m = re.search(r"linkedin\.com/in/([^/?#]+)", unquote(str(url or "")), re.I)
    return m.group(1).strip("/").lower() if m else None


def _name_tokens(name):
    if not name:
        return []
    base = str(name).split(",")[0]  # drop ", CFA" style credentials
    return [t.lower().strip(".") for t in re.split(r"\s+", base)
            if len(t.strip(".")) > 1]


def _known_tokens(*texts):
    toks = set()
    for t in texts:
        for w in re.findall(r"[A-Za-zÀ-ÿ]{4,}", str(t or "").lower()):
            toks.add(w)
    return toks - SERP_NOISE_TOKENS


def _company_from_serp(title, snippet):
    m = LINKEDIN_TITLE_RX.match(str(title or "").strip())
    if m:
        _, c = parse_positions(m.group("headline"))
        if c:
            return c
    em = re.search(r"Experience:\s*([^·|]+)", str(snippet or ""))
    if em:
        c = clean_company(em.group(1))
        if c:
            return c
    am = re.search(r"\b(?:at|@)\s+([A-Z][^.·|,;]{2,60})", str(snippet or ""))
    if am:
        c = clean_company(am.group(1))
        if c:
            return c
    return None


def recover_via_search(client, ctx):
    """Try to recover a company via one search. Returns (data|None, attempts,
    error|None). Abstains (returns None) unless a result passes the
    confidence guard — a wrong company is worse than a missing one."""
    name = str(ctx["name"]).split(",")[0].strip()
    hint = " ".join(str(ctx.get("title") or "").split()[:4]) \
        or str(ctx.get("location") or "").split(",")[0]
    query = f'"{name}" LinkedIn {hint}'.strip()

    results, attempts, error = search_one(client, query)
    if results is None:
        return None, attempts, error
    if not results:
        return None, attempts, "no search results"

    slug = _profile_slug(ctx.get("linkedin_url"))
    toks = _name_tokens(name)
    known = _known_tokens(ctx.get("title"), ctx.get("raw"), ctx.get("location"))

    for r in results:
        r_url = unquote(str(getattr(r, "url", "") or "")).lower()
        r_title = str(getattr(r, "title", "") or "")
        snippet = str(getattr(r, "snippet", "") or "")

        trusted = bool(slug) and f"/in/{slug}" in r_url
        if not trusted and len(toks) >= 2:
            tl = r_title.lower()
            name_ok = toks[0] in tl and toks[-1] in tl
            overlap = known & _known_tokens(r_title, snippet)
            trusted = name_ok and bool(overlap)
        if not trusted:
            continue

        company = _company_from_serp(r_title, snippet)
        if company:
            return {"company": company}, attempts, None

    return None, attempts, "no confident match in results"


# ---------------------------------------------------------------------------
# Progress tracking — console (rich) + browser dashboard
# ---------------------------------------------------------------------------

DASHBOARD_HTML = """<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>LinkedIn Enricher — live</title>
<style>
 body{font-family:Segoe UI,system-ui,sans-serif;background:#111827;color:#e5e7eb;
      margin:0;padding:24px;max-width:860px;margin:auto}
 h1{font-size:20px;margin:12px 0 4px} .muted{color:#9ca3af;font-size:13px}
 .state{font-size:28px;font-weight:700;margin-top:16px}
 .state.done{color:#34d399}.state.run{color:#60a5fa}.state.off{color:#f87171}
 .barwrap{background:#1f2937;border-radius:8px;height:26px;margin:14px 0;overflow:hidden}
 .bar{background:linear-gradient(90deg,#3b82f6,#34d399);height:100%;width:0%;
      transition:width .5s;border-radius:8px}
 .chips{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}
 .chip{background:#1f2937;border-radius:16px;padding:5px 12px;font-size:13px}
 .chip b{color:#fff}
 table{width:100%;border-collapse:collapse;margin-top:10px;font-size:13px}
 td,th{padding:6px 8px;border-bottom:1px solid #1f2937;text-align:left}
 .ok{color:#34d399}.warn{color:#fbbf24}.bad{color:#f87171}
 #summary{background:#064e3b;border-radius:8px;padding:12px 16px;margin-top:14px;display:none}
</style></head><body>
<h1>LinkedIn Enricher</h1><div class="muted" id="meta">waiting for data…</div>
<div class="state run" id="state">…</div><div class="muted" id="tier"></div>
<div class="barwrap"><div class="bar" id="bar"></div></div>
<div class="muted" id="barlabel"></div>
<div class="chips" id="chips"></div>
<div id="summary"></div>
<h1 style="font-size:16px">Last processed rows</h1>
<table id="recent"><tr><th>Name</th><th>Found</th><th>Status</th></tr></table>
<h1 style="font-size:16px">Errors</h1><div class="chips" id="errors"></div>
<script>
function fmt(s){if(s==null)return "—";s=Math.round(s);
 return (s>=60? Math.floor(s/60)+"m "+(s%60)+"s" : s+"s");}
async function tick(){
 try{
  const r=await fetch("progress.json?ts="+Date.now(),{cache:"no-store"});
  const d=await r.json();
  const st=document.getElementById("state");
  st.textContent=d.state + (d.state==="COMPLETE"?" ✅":"");
  st.className="state "+(d.state==="COMPLETE"?"done":"run");
  document.getElementById("tier").textContent=d.tier||"";
  const pct=d.total? Math.round(100*d.done/d.total):0;
  document.getElementById("bar").style.width=pct+"%";
  document.getElementById("barlabel").textContent=
    `${d.done}/${d.total} rows (${pct}%) — now: ${d.current||"—"}`;
  document.getElementById("meta").textContent=
    `TinyFish requests: ${d.requests} · elapsed ${fmt(d.elapsed_s)} · ETA ${fmt(d.eta_s)} · updated ${d.updated}`;
  const chips=document.getElementById("chips");chips.innerHTML="";
  for(const [k,v] of Object.entries(d.counts||{}))
    chips.insertAdjacentHTML("beforeend",`<span class="chip">${k}: <b>${v}</b></span>`);
  const tbl=document.getElementById("recent");
  tbl.innerHTML="<tr><th>Name</th><th>Found</th><th>Status</th></tr>";
  for(const row of (d.recent||[])){
    const cls=row.status==="NEEDS_RESEARCH"?"warn":(row.status==="SKIPPED"?"bad":"ok");
    tbl.insertAdjacentHTML("beforeend",
     `<tr><td>${row.name||""}</td><td>${row.found||"—"}</td><td class="${cls}">${row.status}</td></tr>`);}
  const err=document.getElementById("errors");err.innerHTML="";
  const es=Object.entries(d.errors||{});
  if(!es.length)err.innerHTML='<span class="chip">none 🎉</span>';
  for(const [k,v] of es)
    err.insertAdjacentHTML("beforeend",`<span class="chip warn">${k}: <b>${v}</b></span>`);
  const sm=document.getElementById("summary");
  if(d.summary){sm.style.display="block";
    sm.innerHTML="<b>FINAL SUMMARY</b><br>"+d.summary.replaceAll("\\n","<br>");}
  if(d.state==="COMPLETE")clearInterval(timer);
 }catch(e){
  const st=document.getElementById("state");
  if(st.textContent.indexOf("COMPLETE")<0){st.textContent="waiting / offline";st.className="state off";}
 }}
const timer=setInterval(tick,2000);tick();
</script></body></html>
"""


class _QuietHandler(SimpleHTTPRequestHandler):
    def log_message(self, *args):
        pass


class Tracker:
    """Owns progress state. In CLI mode it also runs a standalone dashboard
    server + rich console; in embedded mode (on_update callback supplied,
    e.g. from the web app) it stays silent and just calls back on every
    change, since the host application owns its own UI."""

    def __init__(self, output_dir, open_browser=True, embedded=False,
                on_update=None):
        self.lock = threading.Lock()
        self.state = "STARTING"
        self.tier = "starting"
        self.done = 0
        self.total = 0
        self.current = ""
        self.counts = {}
        self.requests = 0
        self.errors = {}
        self.recent = deque(maxlen=10)
        self.t0 = time.time()
        self.tier_t0 = time.time()
        self.summary = None
        self.embedded = embedded
        self.on_update = on_update
        self.json_path = output_dir / "progress.json"
        self.html_path = output_dir / "progress.html"
        self.live = None
        self.url = None

        if embedded:
            self._flush()
            return

        self.html_path.write_text(DASHBOARD_HTML, encoding="utf-8")
        try:
            handler = partial(_QuietHandler, directory=str(output_dir))
            self.server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
            threading.Thread(target=self.server.serve_forever, daemon=True).start()
            self.url = f"http://127.0.0.1:{self.server.server_address[1]}/progress.html"
        except Exception:
            self.server = None

        self._flush()
        print(f"Dashboard file: {self.html_path}")
        if self.url:
            print(f"Live dashboard: {self.url}")
            if open_browser:
                try:
                    webbrowser.open(self.url)
                except Exception:
                    pass

        try:
            from rich.console import Console
            from rich.live import Live
            from rich.panel import Panel
            self._Panel = Panel
            console = Console()
            if console.is_terminal:
                self.live = Live(self._render(), console=console,
                                 refresh_per_second=4)
                self.live.start()
        except Exception:
            self.live = None

    # ---- events -----------------------------------------------------------
    def set_tier(self, tier, total, counts):
        with self.lock:
            self.state = "RUNNING"
            self.tier = tier
            self.total = total
            self.done = 0
            self.counts = counts
            self.tier_t0 = time.time()
        self._flush(f"\n=== {tier} — {total} rows ===")

    def row_done(self, name, found, status, counts, requests_add=0):
        with self.lock:
            self.done += 1
            self.current = name or ""
            self.counts = counts
            self.requests += requests_add
            base = str(status).split(":")[0]
            if base == STATUS_NEEDS_RESEARCH and ":" in str(status):
                reason = str(status).split(":", 1)[1].strip()[:60]
                self.errors[reason] = self.errors.get(reason, 0) + 1
            self.recent.appendleft(
                {"name": name, "found": (found or "")[:60], "status": base})
        self._flush(f"  [{self.done}/{self.total}] {name}: {status}")

    def finish(self, summary_text, counts):
        with self.lock:
            self.state = "COMPLETE"
            self.tier = "complete"
            self.counts = counts
            self.summary = summary_text
        self._flush()
        if self.live:
            self.live.stop()

    # ---- output -----------------------------------------------------------
    def _eta(self):
        if not self.done or not self.total or self.done >= self.total:
            return None
        rate = (time.time() - self.tier_t0) / self.done
        return rate * (self.total - self.done)

    def _payload(self):
        eta = self._eta()
        return {
            "state": self.state, "tier": self.tier,
            "done": self.done, "total": self.total, "current": self.current,
            "counts": self.counts, "requests": self.requests,
            "elapsed_s": round(time.time() - self.t0),
            "eta_s": round(eta) if eta else None,
            "recent": list(self.recent), "errors": self.errors,
            "summary": self.summary,
            "updated": time.strftime("%H:%M:%S"),
        }

    def _flush(self, plain_line=None):
        data = self._payload()
        if self.on_update:
            try:
                self.on_update(data)
            except Exception:
                pass
        if self.embedded:
            return
        try:
            tmp = self.json_path.with_suffix(".tmp")
            tmp.write_text(json.dumps(data, ensure_ascii=False),
                           encoding="utf-8")
            os.replace(tmp, self.json_path)
        except OSError:
            pass  # dashboard read collision — next flush wins
        if self.live:
            self.live.update(self._render())
        elif plain_line:
            print(plain_line, flush=True)

    def _render(self):
        d = self._payload()
        width = 34
        frac = d["done"] / d["total"] if d["total"] else 0
        bar = "█" * int(frac * width) + "░" * (width - int(frac * width))
        eta = f"{d['eta_s']}s" if d["eta_s"] else "—"
        counts = "  ".join(f"{k}:{v}" for k, v in d["counts"].items()) or "—"
        body = (
            f"[bold cyan]{d['state']}[/] — {d['tier']}\n"
            f"[{bar}] {d['done']}/{d['total']}\n"
            f"now: {d['current'] or '—'}\n"
            f"{counts}\n"
            f"requests: {d['requests']}  elapsed: {d['elapsed_s']}s  ETA: {eta}"
        )
        if d["summary"]:
            body += f"\n\n[green]{d['summary']}[/]"
        return self._Panel(body, title="LinkedIn Enricher")


# ---------------------------------------------------------------------------
# Excel handling
# ---------------------------------------------------------------------------

class ColumnError(ValueError):
    """A required source column could not be located. Raised (not sys.exit)
    so callers like the web app can report it instead of crashing the
    process."""


def find_col(ws, configured_name, required=True):
    headers = {str(c.value).strip().lower(): i
               for i, c in enumerate(ws[1], start=1) if c.value}
    key = (configured_name or "").strip().lower()
    if key in headers:
        return headers[key]
    if required:
        raise ColumnError(
            f"column '{configured_name}' not found in sheet. "
            f"Available columns: {[c.value for c in ws[1] if c.value]}")
    return None


# Common header spellings for each logical field, tried in order. Lets the
# app accept arbitrary uploads (e.g. 'Person Name'/'LinkedIn URL') without
# the user editing config.yaml first. Matched case-insensitively; exact
# match beats substring.
COLUMN_ALIASES = {
    "name": ["name", "person name", "full name", "person", "contact",
             "contact name", "lead name", "attendee"],
    "positions": ["positions/company", "position", "positions", "title",
                  "job title", "headline", "role", "designation",
                  "position/company"],
    "linkedin": ["linkedin", "linkedin url", "linkedin profile",
                 "profile url", "profile link", "linkedin link", "li url"],
    "location": ["location", "hq location", "city", "region", "geography",
                 "based in", "area"],
}


def _sniff_linkedin_col(ws):
    """Find a column whose data cells contain linkedin.com, as a fallback
    when the header doesn't match any known alias."""
    max_scan = min(ws.max_row, 25)
    for col in range(1, ws.max_column + 1):
        for row in range(2, max_scan + 1):
            v = ws.cell(row=row, column=col).value
            if v and "linkedin.com" in str(v).lower():
                return col
    return None


def autodetect_columns(ws, config=None):
    """Resolve the four source columns for a sheet, trying (1) the exact
    names in config.yaml, then (2) a table of common aliases, then (3) for
    the LinkedIn column, sniffing cell values. Returns a dict with keys
    name/positions/linkedin/location (values are 1-based col indexes or
    None). Only 'positions' or 'linkedin' being findable is required for
    the pipeline to do anything useful."""
    config = config or {}
    headers = {str(c.value).strip().lower(): i
               for i, c in enumerate(ws[1], start=1) if c.value}
    configured = {
        "name": (config.get("name_column") or "").strip().lower(),
        "positions": (config.get("positions_column") or "").strip().lower(),
        "linkedin": (config.get("linkedin_url_column") or "").strip().lower(),
        "location": (config.get("location_column") or "").strip().lower(),
    }
    resolved = {}
    for field, aliases in COLUMN_ALIASES.items():
        col = None
        # 1) exact configured name
        cfg_key = configured.get(field)
        if cfg_key and cfg_key in headers:
            col = headers[cfg_key]
        # 2) alias exact match
        if col is None:
            for a in aliases:
                if a in headers:
                    col = headers[a]
                    break
        # 3) alias substring match (e.g. 'linkedin profile url')
        if col is None:
            for a in aliases:
                hit = next((idx for h, idx in headers.items() if a in h), None)
                if hit:
                    col = hit
                    break
        resolved[field] = col
    if resolved["linkedin"] is None:
        resolved["linkedin"] = _sniff_linkedin_col(ws)
    return resolved


def ensure_output_columns(ws):
    existing = {str(c.value): i for i, c in enumerate(ws[1], start=1) if c.value}
    positions = {}
    next_col = ws.max_column + 1
    for name in NEW_COLUMNS:
        if name in existing:
            positions[name] = existing[name]
        else:
            ws.cell(row=1, column=next_col, value=name)
            positions[name] = next_col
            next_col += 1
    return positions


def cell_get(ws, row, col):
    if col is None:  # source column wasn't detected in this sheet
        return None
    v = ws.cell(row=row, column=col).value
    return str(v).strip() if v is not None and str(v).strip() else None


def base_status(ws, row, cols):
    return (cell_get(ws, row, cols["TF Status"]) or "EMPTY").split(":")[0]


def count_statuses(ws, cols):
    counts = {}
    for row in range(2, ws.max_row + 1):
        counts[base_status(ws, row, cols)] = \
            counts.get(base_status(ws, row, cols), 0) + 1
    return counts


def add_steps(ws, row, cols, n):
    prev = cell_get(ws, row, cols["TF Steps"])
    try:
        prev = int(float(prev)) if prev else 0
    except ValueError:
        prev = 0
    ws.cell(row=row, column=cols["TF Steps"], value=prev + n)


# ---------------------------------------------------------------------------
# Tier runners
# ---------------------------------------------------------------------------

def run_tier1(ws, cols, src, tracker):
    rows = [r for r in range(2, ws.max_row + 1)
            if base_status(ws, r, cols) not in TERMINAL_STATUSES]
    tracker.set_tier("TIER 1 — local parse (free)", len(rows),
                     count_statuses(ws, cols))
    samples = []
    for row in rows:
        raw = cell_get(ws, row, src["positions"])
        name = cell_get(ws, row, src["name"])
        location = cell_get(ws, row, src["location"])
        url = cell_get(ws, row, src["linkedin"])
        if url and "linkedin.com" not in url.lower():
            url = None

        title, company = parse_positions(raw)

        # assign via .value: ws.cell(value=None) silently skips assignment,
        # which would leave stale data behind on re-parsed rows
        ws.cell(row=row, column=cols["person_name"]).value = name
        ws.cell(row=row, column=cols["person_title"]).value = title
        ws.cell(row=row, column=cols["company_name"]).value = company
        ws.cell(row=row, column=cols["linkedin_url"]).value = url
        ws.cell(row=row, column=cols["hq_location"]).value = location
        if not cell_get(ws, row, cols["TF Steps"]):
            ws.cell(row=row, column=cols["TF Steps"], value=0)

        if company:
            status = STATUS_PARSED
        elif url:
            status = STATUS_PENDING
        else:
            status = STATUS_SKIPPED
        ws.cell(row=row, column=cols["TF Status"], value=status)
        samples.append((row, name, raw, title, company, status))
        tracker.row_done(name, company, status, count_statuses(ws, cols))
    return samples


def run_fetch_pass(client, ws, cols, output_path, targets, workers,
                   tracker, label):
    if not targets:
        return
    tracker.set_tier(f"{label} — {workers} workers", len(targets),
                     count_statuses(ws, cols))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {}
        for row in targets:
            url = cell_get(ws, row, cols["linkedin_url"])
            futures[pool.submit(fetch_one, client, url)] = row
            time.sleep(0.2)  # gentle stagger per TinyFish rate-limit guidance

        for fut in as_completed(futures):
            row = futures[fut]
            result, attempts, error = fut.result()
            data = extract_from_fetch(result) if result else None

            if data and data.get("company"):
                status = STATUS_FETCHED
            else:
                reason = error or "authwall or no usable fields"
                status = f"{STATUS_NEEDS_RESEARCH}: {str(reason)[:80]}"

            if data:
                if data.get("title") and not cell_get(ws, row, cols["person_title"]):
                    ws.cell(row=row, column=cols["person_title"], value=data["title"])
                if data.get("company"):
                    ws.cell(row=row, column=cols["company_name"], value=data["company"])
                if data.get("company_url"):
                    ws.cell(row=row, column=cols["company_url"], value=data["company_url"])
                if data.get("full_name") and not cell_get(ws, row, cols["person_name"]):
                    ws.cell(row=row, column=cols["person_name"], value=data["full_name"])
                if data.get("location") and not cell_get(ws, row, cols["hq_location"]):
                    ws.cell(row=row, column=cols["hq_location"], value=data["location"])
            ws.cell(row=row, column=cols["TF Status"], value=status)
            add_steps(ws, row, cols, attempts)

            ws.parent.save(output_path)  # checkpoint: every finished row is safe
            tracker.row_done(cell_get(ws, row, cols["person_name"]),
                             (data or {}).get("company"), status,
                             count_statuses(ws, cols), requests_add=attempts)


def run_search_pass(client, ws, cols, src, output_path, targets, workers,
                    tracker, label):
    if not targets:
        return
    tracker.set_tier(f"{label} — {workers} workers", len(targets),
                     count_statuses(ws, cols))

    def ctx_for(row):
        return {
            "name": cell_get(ws, row, cols["person_name"]),
            "title": cell_get(ws, row, cols["person_title"]),
            "location": cell_get(ws, row, cols["hq_location"]),
            "linkedin_url": cell_get(ws, row, cols["linkedin_url"]),
            "raw": cell_get(ws, row, src["positions"]),
        }

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {}
        for row in targets:
            futures[pool.submit(recover_via_search, client, ctx_for(row))] = row
            time.sleep(0.2)

        for fut in as_completed(futures):
            row = futures[fut]
            data, attempts, error = fut.result()
            if data and data.get("company"):
                ws.cell(row=row, column=cols["company_name"], value=data["company"])
                status = STATUS_SEARCHED
            else:
                status = f"{STATUS_NEEDS_RESEARCH}: search: {str(error)[:70]}"
            ws.cell(row=row, column=cols["TF Status"], value=status)
            add_steps(ws, row, cols, attempts)

            ws.parent.save(output_path)
            tracker.row_done(cell_get(ws, row, cols["person_name"]),
                             (data or {}).get("company"), status,
                             count_statuses(ws, cols), requests_add=attempts)


def reset_statuses(ws, cols, predicate, to_status):
    changed = 0
    for row in range(2, ws.max_row + 1):
        status = cell_get(ws, row, cols["TF Status"]) or ""
        if predicate(status):
            ws.cell(row=row, column=cols["TF Status"], value=to_status)
            changed += 1
    return changed


def is_transient(status):
    return status.startswith(STATUS_NEEDS_RESEARCH) and \
        any(t in status.lower() for t in TRANSIENT_MARKERS)


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------

def print_samples(samples, n):
    parsed = [s for s in samples if s[5] == STATUS_PARSED][:n]
    print(f"\n===== {len(parsed)} SAMPLE PARSED ROWS (raw -> title / company) =====")
    for row, name, raw, title, company, _ in parsed:
        print(f"  row {row} | {name}")
        print(f"    raw:     {(raw or '')[:90]}")
        print(f"    title:   {title}")
        print(f"    company: {company}")
    pending = [s for s in samples if s[5] == STATUS_PENDING]
    if pending:
        print(f"\n===== ROWS NEEDING TIER 2 FETCH ({len(pending)}) =====")
        for row, name, raw, *_ in pending:
            print(f"  row {row} | {name} | raw: {(raw or '')[:60]}")


def build_summary(ws, cols, output_path):
    total = ws.max_row - 1
    counts = count_statuses(ws, cols)
    lines = [f"{total} rows — " + " | ".join(
        f"{s}: {counts[s]}" for s in STATUS_ORDER if counts.get(s))]
    for field in COVERAGE_FIELDS:
        have = sum(1 for r in range(2, ws.max_row + 1)
                   if cell_get(ws, r, cols[field]))
        lines.append(f"{field}: {have}/{total} ({round(100 * have / total)}%)")
    lines.append(f"saved to: {output_path}")
    return "\n".join(lines), counts


# ---------------------------------------------------------------------------
# Pipeline entry point — used by both the CLI and the web app
# ---------------------------------------------------------------------------

def run_pipeline(input_path, do_fetch=False, do_search=False, limit=0,
                 workers=0, reset=False, tracker=None, config=None,
                 verbose=False):
    """Run the enrichment pipeline against input_path and return
    (output_path, summary_text, counts). Reusable by the CLI (main()) and
    by the Flask app (app.py), which supplies its own embedded Tracker.

    verbose=True prints sample rows to stdout (CLI only). The app leaves it
    False — printing lead names with emoji to a Windows cp1252 console
    raises UnicodeEncodeError, and the app has its own UI anyway."""
    config = config or load_config()
    workers = workers or int(config.get("fetch_workers", 8))
    input_path = Path(input_path)
    if not input_path.exists():
        raise FileNotFoundError(f"input file not found: {input_path}")

    OUTPUT_DIR.mkdir(exist_ok=True)
    output_path = OUTPUT_DIR / f"enriched_{input_path.name}"

    # The enriched output IS the checkpoint — always continue from it
    source = output_path if output_path.exists() else input_path
    wb = load_workbook(source)
    ws = wb.active

    src = autodetect_columns(ws, config)
    if src["positions"] is None and src["linkedin"] is None:
        raise ColumnError(
            "could not find a title/headline column or a LinkedIn URL column "
            "in this sheet — nothing to enrich. Detected headers: "
            f"{[c.value for c in ws[1] if c.value]}")
    cols = ensure_output_columns(ws)

    if reset:
        reset_statuses(ws, cols, lambda s: s.startswith(STATUS_NEEDS_RESEARCH),
                       STATUS_PENDING)

    tracker = tracker or Tracker(OUTPUT_DIR, open_browser=False, embedded=True)

    samples = run_tier1(ws, cols, src, tracker)
    wb.save(output_path)

    def pending_rows():
        rows = [r for r in range(2, ws.max_row + 1)
                if base_status(ws, r, cols) == STATUS_PENDING
                and cell_get(ws, r, cols["linkedin_url"])]
        return rows[:limit] if limit else rows

    def searchable_rows(transient_only=False):
        rows = []
        for r in range(2, ws.max_row + 1):
            status = cell_get(ws, r, cols["TF Status"]) or ""
            if cell_get(ws, r, cols["company_name"]):
                continue
            if not cell_get(ws, r, cols["person_name"]):
                continue
            if transient_only:
                if "search:" in status and is_transient(status):
                    rows.append(r)
            elif status.split(":")[0] in (STATUS_NEEDS_RESEARCH, STATUS_PENDING):
                rows.append(r)
        return rows[:limit] if limit else rows

    client = get_tinyfish_client() if (do_fetch or do_search) else None
    full_pipeline = do_fetch and do_search  # mirrors --all in the CLI

    if do_fetch:
        run_fetch_pass(client, ws, cols, output_path, pending_rows(),
                       workers, tracker, "TIER 2 — fetch")
        if full_pipeline:
            n = reset_statuses(ws, cols,
                               lambda s: is_transient(s) and "search:" not in s,
                               STATUS_PENDING)
            if n:
                run_fetch_pass(client, ws, cols, output_path, pending_rows(),
                               max(2, workers // 3), tracker,
                               "TIER 2R — transient retry")

    if do_search:
        run_search_pass(client, ws, cols, src, output_path, searchable_rows(),
                        min(workers, 4), tracker, "TIER 2.5 — search recovery")
        if full_pipeline:
            retry = searchable_rows(transient_only=True)
            if retry:
                run_search_pass(client, ws, cols, src, output_path, retry,
                                2, tracker, "TIER 2.5R — search retry")

    wb.save(output_path)

    if verbose and not (do_fetch or do_search) and samples:
        print_samples(samples, 10)

    summary, counts = build_summary(ws, cols, output_path)
    tracker.finish(summary, counts)
    return output_path, summary, counts


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="Path to input .xlsx")
    ap.add_argument("--all", action="store_true",
                    help="Full pipeline: parse -> fetch -> retry -> search -> retry")
    ap.add_argument("--fetch", action="store_true", help="Tier 2 fetch pass")
    ap.add_argument("--search", action="store_true", help="Tier 2.5 search pass")
    ap.add_argument("--limit", type=int, default=0,
                    help="Max rows per network pass (0 = all)")
    ap.add_argument("--workers", type=int, default=0,
                    help="Parallel workers (default from config.yaml)")
    ap.add_argument("--reset", action="store_true",
                    help="Reset all NEEDS_RESEARCH rows for another attempt")
    ap.add_argument("--no-browser", action="store_true",
                    help="Don't auto-open the dashboard in a browser")
    args = ap.parse_args()

    config = load_config()
    do_fetch = args.all or args.fetch
    do_search = args.all or args.search
    input_path = Path(args.input)
    if not input_path.exists():
        sys.exit(f"ERROR: input file not found: {input_path}")

    output_path = OUTPUT_DIR / f"enriched_{input_path.name}"
    if output_path.exists():
        print(f"Continuing from existing output: {output_path.name}")

    tracker = Tracker(OUTPUT_DIR,
                      open_browser=(do_fetch or do_search) and not args.no_browser)

    try:
        output_path, summary, counts = run_pipeline(
            input_path, do_fetch=do_fetch, do_search=do_search, limit=args.limit,
            workers=args.workers, reset=args.reset, tracker=tracker,
            config=config, verbose=True)
    except ColumnError as e:
        sys.exit(f"ERROR: {e}")

    print("\n===== RUN COMPLETE =====")
    print(summary)

    if tracker.url and (do_fetch or do_search):
        time.sleep(6)  # let the dashboard poll the final COMPLETE state


if __name__ == "__main__":
    main()
