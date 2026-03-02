"""
CB Application Form Stamper — Central Web Server (FIXED)
=========================================================
BUGS FIXED:
  1. init_db() now called at module level — works with gunicorn (was only
     in __main__ before, so tables were never created → 500 on every request)
  2. PDF returned as base64 in /generate response → JS creates a Blob URL
     for download. Eliminates the separate /download route and the
     in-memory token cache that was wiped on every server sleep/restart
     (was causing HTML page to be "downloaded" instead of the PDF)
  3. atomic_next_pair() fixed — no longer nests a manual BEGIN IMMEDIATE
     inside a sqlite3 context manager (which also commits), preventing a
     double-commit error under load
"""

import os
import io
import csv
import base64
import sqlite3
import threading
import datetime
import secrets
from functools import wraps
from zoneinfo import ZoneInfo

from flask import Flask, request, jsonify, send_file, render_template_string, abort
from apscheduler.schedulers.background import BackgroundScheduler
from reportlab.pdfgen import canvas as rl_canvas
from pypdf import PdfReader, PdfWriter

# ─────────────────────────── CONFIG ────────────────────────────
DB_PATH      = "stamper.db"
TEMPLATE_PDF = "template.pdf"        # Your application form PDF
ADMIN_KEY    = os.environ.get("ADMIN_KEY", "admin1234")
TIMEZONE     = "Asia/Kolkata"
MIS_FOLDER   = "mis_reports"

# Stamp positions (points from bottom-left corner of the page)
# A4 = 595 × 842 pt  |  Letter = 612 × 792 pt
# 1 inch = 72 points. Adjust X/Y until the numbers land on your form.
NUM1_X, NUM1_Y = 60, 10   # First  CB number
NUM2_X, NUM2_Y = 100, 10   # Second CB number (right next to first)
FONT_NAME      = "Helvetica-Bold"
FONT_SIZE      = 10
TARGET_PAGE    = 0           # 0 = first page
# ───────────────────────────────────────────────────────────────

app = Flask(__name__)
os.makedirs(MIS_FOLDER, exist_ok=True)

_db_lock = threading.Lock()


# ═══════════════════════════════════════════════════════════════
#  DATABASE
# ═══════════════════════════════════════════════════════════════

def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=20)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_db():
    """Create tables if they don't exist. Safe to call multiple times."""
    conn = get_db()
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS counter (
                id    INTEGER PRIMARY KEY CHECK (id = 1),
                value INTEGER NOT NULL DEFAULT 0
            )
        """)
        conn.execute("""
            INSERT OR IGNORE INTO counter (id, value) VALUES (1, 0)
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS logs (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id  TEXT NOT NULL,
                cb_num_1     TEXT NOT NULL,
                cb_num_2     TEXT NOT NULL,
                generated_at TEXT NOT NULL,
                ip_address   TEXT
            )
        """)
        conn.commit()
    finally:
        conn.close()


# ── FIX 1: call init_db() at import time so gunicorn initialises the DB ──
init_db()


def atomic_next_pair():
    """
    FIX 3: Use a raw connection (no context-manager auto-commit) so we
    can do a single BEGIN IMMEDIATE → UPDATE → COMMIT without double-commit.
    Thread-level lock gives extra safety under concurrent requests.
    """
    with _db_lock:
        conn = get_db()
        try:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute("SELECT value FROM counter WHERE id=1").fetchone()
            current = row["value"]
            conn.execute("UPDATE counter SET value=? WHERE id=1", (current + 2,))
            conn.execute("COMMIT")
        except Exception:
            conn.execute("ROLLBACK")
            conn.close()
            raise
        finally:
            conn.close()

    return (f"CD{current + 1:08d}", f"CD{current + 2:08d}")


def log_generation(employee_id, cb1, cb2, ip):
    tz = ZoneInfo(TIMEZONE)
    ts = datetime.datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    try:
        conn.execute(
            "INSERT INTO logs (employee_id, cb_num_1, cb_num_2, generated_at, ip_address) "
            "VALUES (?, ?, ?, ?, ?)",
            (employee_id.strip().upper(), cb1, cb2, ts, ip)
        )
        conn.commit()
    finally:
        conn.close()


# ═══════════════════════════════════════════════════════════════
#  PDF STAMPING
# ═══════════════════════════════════════════════════════════════

def stamp_pdf(cb1: str, cb2: str) -> bytes:
    if not os.path.exists(TEMPLATE_PDF):
        raise FileNotFoundError(
            "template.pdf was not found on the server. "
            "Please upload your application form PDF and name it 'template.pdf'."
        )

    reader = PdfReader(TEMPLATE_PDF)
    page   = reader.pages[TARGET_PAGE]
    pw     = float(page.mediabox.width)
    ph     = float(page.mediabox.height)

    # Build a transparent overlay page containing only the two CB numbers
    overlay_buf = io.BytesIO()
    c = rl_canvas.Canvas(overlay_buf, pagesize=(pw, ph))
    c.setFont(FONT_NAME, FONT_SIZE)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(NUM1_X, NUM1_Y, cb1)
    c.drawString(NUM2_X, NUM2_Y, cb2)
    c.save()
    overlay_buf.seek(0)

    overlay_page = PdfReader(overlay_buf).pages[0]
    page.merge_page(overlay_page)

    writer = PdfWriter()
    writer.add_page(page)
    for i, p in enumerate(reader.pages):
        if i != TARGET_PAGE:
            writer.add_page(p)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


# ═══════════════════════════════════════════════════════════════
#  MIS REPORT
# ═══════════════════════════════════════════════════════════════

def generate_mis(target_date: datetime.date = None) -> str:
    tz = ZoneInfo(TIMEZONE)
    if target_date is None:
        target_date = datetime.datetime.now(tz).date()

    date_str = target_date.strftime("%Y-%m-%d")
    filepath = os.path.join(MIS_FOLDER, f"MIS_{date_str}.csv")

    conn = get_db()
    try:
        rows = conn.execute(
            "SELECT employee_id, cb_num_1, cb_num_2, generated_at, ip_address "
            "FROM logs WHERE DATE(generated_at) = ? ORDER BY id",
            (date_str,)
        ).fetchall()
    finally:
        conn.close()

    summary: dict[str, int] = {}
    for r in rows:
        summary[r["employee_id"]] = summary.get(r["employee_id"], 0) + 1

    with open(filepath, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["CB APPLICATION FORM — MIS REPORT"])
        w.writerow(["Date:", date_str])
        w.writerow(["Generated at:", datetime.datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")])
        w.writerow(["Total Forms Generated:", len(rows)])
        w.writerow([])
        w.writerow(["EMPLOYEE SUMMARY"])
        w.writerow(["Employee ID", "Forms Generated", "Application Numbers Used"])
        for eid, count in sorted(summary.items()):
            w.writerow([eid, count, count * 2])
        w.writerow([])
        w.writerow(["DETAIL LOG"])
        w.writerow(["#", "Employee ID", "CB Number 1", "CB Number 2", "Generated At", "IP Address"])
        for i, r in enumerate(rows, 1):
            w.writerow([i, r["employee_id"], r["cb_num_1"],
                        r["cb_num_2"], r["generated_at"], r["ip_address"]])

    print(f"[MIS] {filepath} — {len(rows)} records")
    return filepath


# ═══════════════════════════════════════════════════════════════
#  SCHEDULER — 8:00 PM daily
# ═══════════════════════════════════════════════════════════════

scheduler = BackgroundScheduler(timezone=TIMEZONE)
scheduler.add_job(generate_mis, "cron", hour=20, minute=0)
scheduler.start()


# ═══════════════════════════════════════════════════════════════
#  HTML FRONTEND
# ═══════════════════════════════════════════════════════════════

HTML_PAGE = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CB Application Form Generator</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --ink:   #0f1923; --paper: #faf8f4; --gold:  #c8973a;
    --gold2: #e8b45a; --muted: #8a8070; --line:  #e2ddd4;
    --green: #2a7a4f; --red:   #b94040;
  }
  body {
    font-family: 'DM Sans', sans-serif;
    background: var(--paper); color: var(--ink);
    min-height: 100vh; display: flex; flex-direction: column; align-items: center;
  }
  header {
    width: 100%; background: var(--ink);
    padding: 18px 40px; display: flex; align-items: center; gap: 14px;
  }
  .logo-mark {
    width: 42px; height: 42px; background: var(--gold); border-radius: 4px;
    display: flex; align-items: center; justify-content: center;
    font-family: 'DM Serif Display', serif; font-size: 20px;
    color: var(--ink); flex-shrink: 0;
  }
  header h1 { font-family: 'DM Serif Display', serif; font-size: 20px; color: #fff; font-weight: 400; }
  header span { color: var(--gold2); }
  main { width: 100%; max-width: 560px; padding: 48px 20px 80px; }
  .eyebrow { font-size: 11px; font-weight: 600; letter-spacing: 2.5px; text-transform: uppercase; color: var(--gold); margin-bottom: 10px; }
  .headline { font-family: 'DM Serif Display', serif; font-size: 36px; line-height: 1.15; margin-bottom: 8px; }
  .headline em { font-style: italic; color: var(--muted); }
  .sub { font-size: 14px; color: var(--muted); line-height: 1.6; margin-bottom: 40px; }
  .card { background: #fff; border: 1px solid var(--line); border-radius: 12px; padding: 36px; box-shadow: 0 4px 32px rgba(15,25,35,0.06); }
  label { display: block; font-size: 12px; font-weight: 600; letter-spacing: 1.2px; text-transform: uppercase; color: var(--muted); margin-bottom: 8px; }
  input[type="text"] {
    width: 100%; padding: 13px 16px;
    font-family: 'DM Sans', sans-serif; font-size: 15px; color: var(--ink);
    background: var(--paper); border: 1.5px solid var(--line); border-radius: 8px;
    outline: none; transition: border-color .2s, box-shadow .2s; letter-spacing: 1px;
  }
  input[type="text"]:focus { border-color: var(--gold); box-shadow: 0 0 0 3px rgba(200,151,58,.12); }
  input[type="text"]::placeholder { color: #c0b9ae; }
  .divider { height: 1px; background: var(--line); margin: 28px 0; }
  .info-box {
    background: #f7f4ee; border-left: 3px solid var(--gold); border-radius: 4px;
    padding: 14px 16px; font-size: 13px; color: var(--muted); line-height: 1.6; margin-bottom: 28px;
  }
  .info-box strong { color: var(--ink); }
  .btn {
    width: 100%; padding: 15px; background: var(--ink); color: #fff;
    font-family: 'DM Sans', sans-serif; font-size: 14px; font-weight: 600;
    letter-spacing: 1px; text-transform: uppercase; border: none; border-radius: 8px;
    cursor: pointer; transition: background .2s, transform .1s;
    display: flex; align-items: center; justify-content: center; gap: 10px;
  }
  .btn:hover { background: #1e2f3d; }
  .btn:active { transform: scale(.98); }
  .btn:disabled { background: #9aafbf; cursor: not-allowed; transform: none; }
  .result { display: none; margin-top: 28px; background: #f0f8f4; border: 1.5px solid #a8d8bc; border-radius: 10px; padding: 24px; }
  .result.error { background: #fdf2f2; border-color: #e8b0b0; }
  .result-label { font-size: 11px; font-weight: 600; letter-spacing: 1.5px; text-transform: uppercase; color: var(--green); margin-bottom: 12px; }
  .result.error .result-label { color: var(--red); }
  .cb-pair { display: flex; gap: 12px; margin-bottom: 18px; flex-wrap: wrap; }
  .cb-badge { background: var(--ink); color: var(--gold2); font-family: 'DM Serif Display', serif; font-size: 18px; padding: 8px 18px; border-radius: 6px; letter-spacing: 1.5px; }
  .dl-btn {
    display: inline-flex; align-items: center; gap: 8px; padding: 11px 22px;
    background: var(--green); color: #fff; font-family: 'DM Sans', sans-serif;
    font-size: 13px; font-weight: 600; border-radius: 7px;
    border: none; cursor: pointer; transition: opacity .2s;
  }
  .dl-btn:hover { opacity: .88; }
  .spinner { width: 18px; height: 18px; border: 2.5px solid rgba(255,255,255,.3); border-top-color: #fff; border-radius: 50%; animation: spin .7s linear infinite; display: none; }
  @keyframes spin { to { transform: rotate(360deg); } }
  footer { position: fixed; bottom: 0; left: 0; right: 0; background: var(--ink); padding: 10px 24px; display: flex; align-items: center; justify-content: center; gap: 20px; font-size: 12px; color: #8899aa; }
  footer a { color: var(--gold2); text-decoration: none; }
  footer a:hover { text-decoration: underline; }
</style>
</head>
<body>

<header>
  <div class="logo-mark">CB</div>
  <h1>Application Form Generator — <span>Central Series</span></h1>
</header>

<main>
  <div class="eyebrow">Controlled Issue System</div>
  <h2 class="headline">Generate Your<br><em>Application Form</em></h2>
  <p class="sub">Each generation issues two unique, sequential CB numbers stamped directly onto your form. All activity is logged centrally.</p>

  <div class="card">
    <label for="empId">Employee ID</label>
    <input type="text" id="empId" placeholder="e.g. EMP1042" maxlength="20" autocomplete="off">

    <div class="divider"></div>

    <div class="info-box">
      <strong>Each form gets two numbers:</strong><br>
      e.g. <strong>CB00000001</strong> &amp; <strong>CB00000002</strong> — sequential, unique, never repeated.
      Your Employee ID is logged against every issuance.
    </div>

    <button class="btn" id="genBtn" onclick="generateForm()">
      <div class="spinner" id="spin"></div>
      <span id="btnTxt">⚡ Generate &amp; Download Form</span>
    </button>

    <div class="result" id="result">
      <div class="result-label" id="resultLabel">Form Generated</div>
      <div class="cb-pair" id="cbPair"></div>
      <!-- FIX 2: download triggered via JS Blob URL — no server round-trip needed -->
      <button class="dl-btn" id="dlBtn" onclick="triggerDownload()">
        ↓ Download Stamped PDF
      </button>
    </div>
  </div>
</main>

<footer>
  <span>Numbers issued: Central Counter</span>
  <a href="/admin/mis?key={{ admin_key }}" target="_blank">Admin: Download Today's MIS</a>
</footer>

<script>
// Holds the base64 PDF data returned by /generate
let _pdfB64 = null;
let _cb1 = '', _cb2 = '';

async function generateForm() {
  const empId = document.getElementById('empId').value.trim();
  if (!empId) {
    alert('Please enter your Employee ID before generating.');
    document.getElementById('empId').focus();
    return;
  }

  const btn = document.getElementById('genBtn');
  const spin = document.getElementById('spin');
  const txt  = document.getElementById('btnTxt');
  const res  = document.getElementById('result');

  btn.disabled = true;
  spin.style.display = 'block';
  txt.textContent = 'Generating...';
  res.style.display = 'none';
  _pdfB64 = null;

  try {
    const resp = await fetch('/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ employee_id: empId })
    });

    // Always parse as JSON — our server always returns JSON
    const data = await resp.json();

    if (!resp.ok || data.error) {
      throw new Error(data.error || `Server error (${resp.status})`);
    }

    _pdfB64 = data.pdf_b64;    // base64-encoded PDF bytes
    _cb1    = data.cb1;
    _cb2    = data.cb2;

    document.getElementById('cbPair').innerHTML =
      `<div class="cb-badge">${data.cb1}</div><div class="cb-badge">${data.cb2}</div>`;

    document.getElementById('resultLabel').textContent = '✓ Form Generated Successfully';
    res.className = 'result';
    res.style.display = 'block';

    // Auto-trigger the download immediately
    triggerDownload();

  } catch (err) {
    document.getElementById('resultLabel').textContent = '✗ Error';
    document.getElementById('cbPair').innerHTML =
      `<span style="color:var(--red);font-size:14px">${err.message}</span>`;
    document.getElementById('dlBtn').style.display = 'none';
    res.className = 'result error';
    res.style.display = 'block';
  } finally {
    btn.disabled = false;
    spin.style.display = 'none';
    txt.textContent = '⚡ Generate & Download Form';
  }
}

function triggerDownload() {
  if (!_pdfB64) return;

  // Convert base64 → binary → Blob → object URL → click download
  // This approach works in all browsers and requires NO second server request
  const byteChars = atob(_pdfB64);
  const byteArr   = new Uint8Array(byteChars.length);
  for (let i = 0; i < byteChars.length; i++) {
    byteArr[i] = byteChars.charCodeAt(i);
  }
  const blob    = new Blob([byteArr], { type: 'application/pdf' });
  const url     = URL.createObjectURL(blob);
  const anchor  = document.createElement('a');
  anchor.href     = url;
  anchor.download = `ApplicationForm_${_cb1}_${_cb2}.pdf`;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  // Release memory after a short delay
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

document.getElementById('empId').addEventListener('keydown', e => {
  if (e.key === 'Enter') generateForm();
});
</script>
</body>
</html>
"""


# ═══════════════════════════════════════════════════════════════
#  ROUTES
# ═══════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template_string(HTML_PAGE, admin_key=ADMIN_KEY)


@app.route("/generate", methods=["POST"])
def api_generate():
    data        = request.get_json(silent=True) or {}
    employee_id = (data.get("employee_id") or "").strip()

    if not employee_id:
        return jsonify({"error": "Employee ID is required."}), 400
    if len(employee_id) > 20:
        return jsonify({"error": "Employee ID too long (max 20 chars)."}), 400

    try:
        cb1, cb2   = atomic_next_pair()
        pdf_bytes  = stamp_pdf(cb1, cb2)

        # FIX 2: encode PDF as base64 and return it directly in the JSON response.
        # The browser decodes it into a Blob — no second HTTP round-trip, no cache,
        # no token that can expire or be lost on server restart.
        pdf_b64 = base64.b64encode(pdf_bytes).decode("ascii")

        log_generation(employee_id, cb1, cb2, request.remote_addr)
        return jsonify({"cb1": cb1, "cb2": cb2, "pdf_b64": pdf_b64})

    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 500
    except Exception as e:
        app.logger.exception("Generate error")
        return jsonify({"error": f"Internal server error: {e}"}), 500


# ─── Admin routes ───────────────────────────────────────────────

def require_admin(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        key = request.args.get("key") or request.headers.get("X-Admin-Key", "")
        if key != ADMIN_KEY:
            abort(403)
        return f(*args, **kwargs)
    return wrapper


@app.route("/admin/mis")
@require_admin
def admin_mis():
    """On-demand: download today's MIS CSV."""
    filepath = generate_mis()
    return send_file(filepath, as_attachment=True,
                     download_name=os.path.basename(filepath),
                     mimetype="text/csv")


@app.route("/admin/mis/<date_str>")
@require_admin
def admin_mis_date(date_str):
    """Download MIS for a specific date: /admin/mis/2024-03-15?key=..."""
    try:
        target = datetime.date.fromisoformat(date_str)
    except ValueError:
        abort(400)
    filepath = generate_mis(target)
    return send_file(filepath, as_attachment=True,
                     download_name=os.path.basename(filepath),
                     mimetype="text/csv")


@app.route("/admin/stats")
@require_admin
def admin_stats():
    tz    = ZoneInfo(TIMEZONE)
    today = datetime.datetime.now(tz).date().isoformat()
    conn  = get_db()
    try:
        counter     = conn.execute("SELECT value FROM counter WHERE id=1").fetchone()["value"]
        total_today = conn.execute(
            "SELECT COUNT(*) AS c FROM logs WHERE DATE(generated_at)=?", (today,)
        ).fetchone()["c"]
        top_users   = conn.execute(
            "SELECT employee_id, COUNT(*) AS forms FROM logs "
            "WHERE DATE(generated_at)=? GROUP BY employee_id ORDER BY forms DESC LIMIT 10",
            (today,)
        ).fetchall()
    finally:
        conn.close()

    return jsonify({
        "last_number_issued":    counter,
        "last_cb_number":        f"CB{counter:08d}",
        "forms_generated_today": total_today,
        "numbers_issued_today":  total_today * 2,
        "top_users_today": [
            {"employee_id": r["employee_id"], "forms": r["forms"]} for r in top_users
        ]
    })


# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
