"""
LinkedIn Enricher — desktop web application.

Wraps the existing enrich.py pipeline and the scoring ledger tools behind
a local browser UI: upload a lead file, run Tier 1/2/2.5 enrichment with
a live progress bar, then download the enriched spreadsheet. Scoring
(the icp_model.md research pass) still runs through Claude Code's
score-leads workflow — this app covers the parts that don't need an LLM
in the loop, run/rerun without opening a chat session.

Usage:
    python app.py
    -> picks the first free port (5000, then 5001, ...) and opens that
       exact URL in your browser. The console prints the real URL too.

Each run happens in a background thread; progress is exposed over
Server-Sent Events so the page updates live without polling a file.
"""
import io
import json
import queue
import socket
import threading
import time
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file, Response
from openpyxl import load_workbook

import enrich

BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)

# One run at a time, tracked in memory. Simpler than a job queue/DB for a
# single-user local tool, and matches how enrich.py already works (the
# output workbook is the real checkpoint, this is just live UI state).
RUNS = {}
RUNS_LOCK = threading.Lock()


class Job:
    def __init__(self, job_id, filename):
        self.id = job_id
        self.filename = filename
        self.status = "queued"  # queued -> running -> done -> error
        self.error = None
        self.result = None
        self.listeners = []  # SSE subscriber queues
        self.last_payload = None
        self.lock = threading.Lock()

    def publish(self, payload):
        with self.lock:
            self.last_payload = payload
            for q in list(self.listeners):
                q.put(payload)

    def subscribe(self):
        q = queue.Queue()
        with self.lock:
            if self.last_payload:
                q.put(self.last_payload)
            self.listeners.append(q)
        return q

    def unsubscribe(self, q):
        with self.lock:
            if q in self.listeners:
                self.listeners.remove(q)


def make_tracker(job):
    """An enrich.Tracker whose on_update callback republishes to the job's
    SSE subscribers, so the browser sees the same live counts/ETA/recent
    rows that the CLI's rich console shows."""
    return enrich.Tracker(OUTPUT_DIR, open_browser=False, embedded=True,
                          on_update=job.publish)


def run_job(job, input_path, do_fetch, do_search, reset):
    job.status = "running"
    try:
        tracker = make_tracker(job)
        output_path, summary, counts = enrich.run_pipeline(
            input_path, do_fetch=do_fetch, do_search=do_search,
            reset=reset, tracker=tracker)
        job.result = {
            "output_path": str(output_path),
            "output_name": output_path.name,
            "summary": summary,
            "counts": counts,
        }
        job.status = "done"
    except enrich.ColumnError as e:
        # expected, user-fixable: the sheet's columns weren't recognized
        msg = (f"This file's columns weren't recognized. {e} "
               "Rename a column to something standard (e.g. 'LinkedIn URL', "
               "'Title', 'Name') or set the names in config.yaml.")
        job.error = msg
        job.status = "error"
        job.publish({"state": "ERROR", "error": msg})
    except Exception as e:  # noqa: BLE001 - surface any failure to the UI
        job.error = str(e)
        job.status = "error"
        job.publish({"state": "ERROR", "error": str(e)})


@app.route("/")
def index():
    return render_template("index.html")


def describe_columns(ws, config):
    """Report which logical fields were auto-detected in a sheet, so the UI
    can preview the mapping and warn if the file isn't enrichable."""
    src = enrich.autodetect_columns(ws, config)
    headers = [c.value for c in ws[1] if c.value]

    def hname(idx):
        return ws.cell(row=1, column=idx).value if idx else None

    detected = {field: hname(idx) for field, idx in src.items()}
    enrichable = bool(src["positions"] or src["linkedin"])
    return {"headers": headers, "detected": detected,
            "enrichable": enrichable}


@app.route("/api/files")
def list_files():
    """Files already sitting in input/ or output/ that can be re-run or
    downloaded, so the user isn't limited to a fresh upload each time."""
    inputs = sorted(p.name for p in INPUT_DIR.glob("*.xlsx"))
    enriched = sorted(p.name for p in OUTPUT_DIR.glob("enriched_*.xlsx"))
    scored = sorted(p.name for p in OUTPUT_DIR.glob("*_scored.xlsx"))
    return jsonify({"inputs": inputs, "enriched": enriched, "scored": scored})


@app.route("/api/upload", methods=["POST"])
def upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "no file provided"}), 400
    if not f.filename.lower().endswith((".xlsx", ".csv")):
        return jsonify({"error": "only .xlsx or .csv files are supported"}), 400

    dest = INPUT_DIR / f.filename
    if f.filename.lower().endswith(".csv"):
        # enrich.py requires xlsx — convert on the way in, mirroring the
        # manual CSV->xlsx step this project has always needed.
        import csv as csv_module
        text = f.stream.read().decode("utf-8-sig")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        for row in csv_module.reader(io.StringIO(text)):
            if any(cell.strip() for cell in row):
                ws.append(row)
        dest = INPUT_DIR / (Path(f.filename).stem + ".xlsx")
        wb.save(dest)
    else:
        f.save(dest)

    try:
        ws = load_workbook(dest).active
    except Exception as e:  # noqa: BLE001 - malformed spreadsheet
        return jsonify({"error": f"could not read spreadsheet: {e}"}), 400
    info = describe_columns(ws, enrich.load_config())
    return jsonify({"filename": dest.name, "rows": ws.max_row - 1, **info})


@app.route("/api/preview")
def preview():
    """Column headers + row count + detected mapping for a file already in
    input/, so the UI can show what it's about to process before running."""
    filename = request.args.get("filename", "")
    path = INPUT_DIR / filename
    if not path.exists():
        return jsonify({"error": "file not found"}), 404
    ws = load_workbook(path).active
    info = describe_columns(ws, enrich.load_config())
    return jsonify({"filename": filename, "rows": ws.max_row - 1, **info})


@app.route("/api/run", methods=["POST"])
def start_run():
    data = request.get_json(force=True)
    filename = data.get("filename")
    mode = data.get("mode", "all")  # tier1 | fetch | all
    reset = bool(data.get("reset"))

    input_path = INPUT_DIR / filename if filename else None
    if not input_path or not input_path.exists():
        return jsonify({"error": "file not found in input/"}), 404

    with RUNS_LOCK:
        # refuse a second concurrent run — TinyFish workers + the single
        # checkpoint workbook aren't safe to drive from two jobs at once
        for j in RUNS.values():
            if j.status == "running":
                return jsonify({"error": "a run is already in progress",
                                "job_id": j.id}), 409

        job_id = str(int(time.time() * 1000))
        job = Job(job_id, filename)
        RUNS[job_id] = job

    do_fetch = mode in ("fetch", "all")
    do_search = mode == "all"
    threading.Thread(target=run_job, args=(job, input_path, do_fetch,
                                           do_search, reset),
                     daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/api/status/<job_id>")
def job_status(job_id):
    job = RUNS.get(job_id)
    if not job:
        return jsonify({"error": "unknown job"}), 404
    payload = {"status": job.status, "error": job.error, "result": job.result}
    if job.last_payload:
        payload["progress"] = job.last_payload
    return jsonify(payload)


@app.route("/api/stream/<job_id>")
def stream(job_id):
    job = RUNS.get(job_id)
    if not job:
        return jsonify({"error": "unknown job"}), 404

    def gen():
        q = job.subscribe()
        try:
            while True:
                try:
                    payload = q.get(timeout=25)
                    yield f"data: {json.dumps(payload)}\n\n"
                    if payload.get("state") in ("COMPLETE", "ERROR"):
                        break
                except queue.Empty:
                    yield ": keepalive\n\n"
        finally:
            job.unsubscribe(q)

    return Response(gen(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache",
                            "X-Accel-Buffering": "no"})


@app.route("/api/download/<path:filename>")
def download(filename):
    path = OUTPUT_DIR / filename
    if not path.exists() or path.parent != OUTPUT_DIR:
        return jsonify({"error": "file not found"}), 404
    return send_file(path, as_attachment=True, download_name=filename)


def find_free_port(preferred=5000, span=20):
    """Return the first genuinely-free port at or after `preferred`. Avoids
    the silent 'address already in use' trap where the browser opens a
    stale instance on 5000 instead of this fresh one.

    Note: no SO_REUSEADDR on the probe — with it, bind() can succeed on a
    port another process is already listening on, which would defeat the
    whole point of detecting the collision."""
    for port in range(preferred, preferred + span):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError(f"no free port in {preferred}..{preferred + span}")


def main():
    port = find_free_port(5000)
    url = f"http://127.0.0.1:{port}"
    # open the browser to the ACTUAL bound port, not a hardcoded guess
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    print("=" * 60)
    print(f"  LinkedIn Enricher is running at:  {url}")
    print("  (Ctrl+C here to stop it)")
    print("=" * 60)
    app.run(host="127.0.0.1", port=port, threaded=True)


if __name__ == "__main__":
    main()
