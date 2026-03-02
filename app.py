"""
CB Application Form Stamper — Central Web Server
=================================================
• Atomic sequential CB numbers (never duplicated, never skipped)
• 2 numbers issued per form generation
• Employee ID logging
• Daily MIS report at 8:00 PM (auto-generated + downloadable)
• PDF stamping and download
"""

import os
import io
import csv
import sqlite3
import threading
import datetime
from functools import wraps
from zoneinfo import ZoneInfo

from flask import (Flask, request, jsonify, send_file,
                   render_template_string, abort)
from apscheduler.schedulers.background import BackgroundScheduler
from reportlab.pdfgen import canvas as rl_canvas
from pypdf import PdfReader, PdfWriter

# ─────────────────────────── CONFIG ────────────────────────────
DB_PATH        = "stamper.db"
TEMPLATE_PDF   = "template.pdf"          # Place your PDF here
ADMIN_KEY      = os.environ.get("ADMIN_KEY", "admin1234")  # Change in production!
TIMEZONE       = "Asia/Kolkata"          # Change to your timezone
MIS_FOLDER     = "mis_reports"

# ── Number stamp positions (points from bottom-left of page) ──
# Adjust these after first test print
NUM1_X, NUM1_Y = 370, 718   # Position of first CB number
NUM2_X, NUM2_Y = 480, 718   # Position of second CB number (next to first)
FONT_NAME      = "Helvetica-Bold"
FONT_SIZE      = 10
TARGET_PAGE    = 0           # 0 = first page
# ───────────────────────────────────────────────────────────────

app = Flask(__name__)
os.makedirs(MIS_FOLDER, exist_ok=True)

# Thread-level DB lock (extra safety beyond SQLite's own locking)
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
    with get_db() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS counter (
                id      INTEGER PRIMARY KEY CHECK (id = 1),
                value   INTEGER NOT NULL DEFAULT 0
            );
            INSERT OR IGNORE INTO counter (id, value) VALUES (1, 0);

            CREATE TABLE IF NOT EXISTS logs (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id TEXT    NOT NULL,
                cb_num_1    TEXT    NOT NULL,
                cb_num_2    TEXT    NOT NULL,
                generated_at TEXT   NOT NULL,
                ip_address  TEXT
            );
        """)


def atomic_next_pair():
    """
    Atomically fetch and increment the counter by 2.
    Returns (cb_num_1, cb_num_2) as formatted strings.
    Guaranteed unique even under high concurrency.
    """
    with _db_lock:
        with get_db() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute("SELECT value FROM counter WHERE id=1").fetchone()
            current = row["value"]
            next_val = current + 2
            conn.execute("UPDATE counter SET value=? WHERE id=1", (next_val,))
            conn.commit()

    n1 = current + 1
    n2 = current + 2
    return (f"CB{n1:08d}", f"CB{n2:08d}")


def log_generation(employee_id, cb1, cb2, ip):
    tz = ZoneInfo(TIMEZONE)
    ts = datetime.datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
    with get_db() as conn:
        conn.execute(
            "INSERT INTO logs (employee_id, cb_num_1, cb_num_2, generated_at, ip_address) "
            "VALUES (?,?,?,?,?)",
            (employee_id.strip().upper(), cb1, cb2, ts, ip)
        )


# ═══════════════════════════════════════════════════════════════
#  PDF STAMPING
# ═══════════════════════════════════════════════════════════════

def stamp_pdf(cb1: str, cb2: str) -> bytes:
    if not os.path.exists(TEMPLATE_PDF):
        raise FileNotFoundError("template.pdf not found on server.")

    reader = PdfReader(TEMPLATE_PDF)
    page   = reader.pages[TARGET_PAGE]
    pw     = float(page.mediabox.width)
    ph     = float(page.mediabox.height)

    # Build overlay with both numbers
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
    out.seek(0)
    return out.read()


# ═══════════════════════════════════════════════════════════════
#  MIS REPORT GENERATOR
# ═══════════════════════════════════════════════════════════════

def generate_mis(target_date: datetime.date = None):
    tz = ZoneInfo(TIMEZONE)
    if target_date is None:
        target_date = datetime.datetime.now(tz).date()

    date_str = target_date.strftime("%Y-%m-%d")
    filename = f"MIS_{date_str}.csv"
    filepath = os.path.join(MIS_FOLDER, filename)

    with get_db() as conn:
        rows = conn.execute(
            "SELECT employee_id, cb_num_1, cb_num_2, generated_at, ip_address "
            "FROM logs WHERE DATE(generated_at) = ? ORDER BY id",
            (date_str,)
        ).fetchall()

        # Summary per employee
        summary = {}
        for r in rows:
            eid = r["employee_id"]
            summary[eid] = summary.get(eid, 0) + 1

    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        # ── Header block ──
        writer.writerow(["CB APPLICATION FORM — MIS REPORT"])
        writer.writerow(["Date:", date_str])
        writer.writerow(["Generated at:",
                         datetime.datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")])
        writer.writerow(["Total Forms Generated:", len(rows)])
        writer.writerow([])

        # ── Employee summary ──
        writer.writerow(["EMPLOYEE SUMMARY"])
        writer.writerow(["Employee ID", "Forms Generated", "Application Numbers Used"])
        for eid, count in sorted(summary.items()):
            writer.writerow([eid, count, count * 2])
        writer.writerow([])

        # ── Detail log ──
        writer.writerow(["DETAIL LOG"])
        writer.writerow(["#", "Employee ID", "CB Number 1",
                         "CB Number 2", "Generated At", "IP Address"])
        for i, r in enumerate(rows, 1):
            writer.writerow([i, r["employee_id"], r["cb_num_1"],
                             r["cb_num_2"], r["generated_at"], r["ip_address"]])

    print(f"[MIS] Report generated: {filepath} ({len(rows)} records)")
    return filepath


# ═══════════════════════════════════════════════════════════════
#  SCHEDULER — 8:00 PM daily
# ═══════════════════════════════════════════════════════════════

scheduler = BackgroundScheduler(timezone=TIMEZONE)
scheduler.add_job(generate_mis, "cron", hour=20, minute=0)
scheduler.start()


# ═══════════════════════════════════════════════════════════════
#  ROUTES
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
    --ink:    #0f1923;
    --paper:  #faf8f4;
    --gold:   #c8973a;
    --gold2:  #e8b45a;
    --muted:  #8a8070;
    --line:   #e2ddd4;
    --green:  #2a7a4f;
    --red:    #b94040;
  }

  body {
    font-family: 'DM Sans', sans-serif;
    background: var(--paper);
    color: var(--ink);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
  }

  /* ── Header ── */
  header {
    width: 100%;
    background: var(--ink);
    padding: 18px 40px;
    display: flex;
    align-items: center;
    gap: 14px;
  }
  .logo-mark {
    width: 42px; height: 42px;
    background: var(--gold);
    border-radius: 4px;
    display: flex; align-items: center; justify-content: center;
    font-family: 'DM Serif Display', serif;
    font-size: 20px; color: var(--ink); font-weight: bold;
    flex-shrink: 0;
  }
  header h1 {
    font-family: 'DM Serif Display', serif;
    font-size: 20px; color: #fff; font-weight: 400;
    letter-spacing: 0.3px;
  }
  header span { color: var(--gold2); }

  /* ── Main card ── */
  main {
    width: 100%;
    max-width: 560px;
    padding: 48px 20px 80px;
  }

  .eyebrow {
    font-size: 11px; font-weight: 600; letter-spacing: 2.5px;
    text-transform: uppercase; color: var(--gold);
    margin-bottom: 10px;
  }

  .headline {
    font-family: 'DM Serif Display', serif;
    font-size: 36px; line-height: 1.15; color: var(--ink);
    margin-bottom: 8px;
  }
  .headline em { font-style: italic; color: var(--muted); }

  .sub {
    font-size: 14px; color: var(--muted); line-height: 1.6;
    margin-bottom: 40px;
  }

  /* ── Form card ── */
  .card {
    background: #fff;
    border: 1px solid var(--line);
    border-radius: 12px;
    padding: 36px;
    box-shadow: 0 4px 32px rgba(15,25,35,0.06);
  }

  label {
    display: block;
    font-size: 12px; font-weight: 600; letter-spacing: 1.2px;
    text-transform: uppercase; color: var(--muted);
    margin-bottom: 8px;
  }

  input[type="text"] {
    width: 100%;
    padding: 13px 16px;
    font-family: 'DM Sans', sans-serif;
    font-size: 15px; color: var(--ink);
    background: var(--paper);
    border: 1.5px solid var(--line);
    border-radius: 8px;
    outline: none;
    transition: border-color .2s, box-shadow .2s;
    letter-spacing: 1px;
  }
  input[type="text"]:focus {
    border-color: var(--gold);
    box-shadow: 0 0 0 3px rgba(200,151,58,.12);
  }
  input[type="text"]::placeholder { color: #c0b9ae; }

  .divider {
    height: 1px; background: var(--line);
    margin: 28px 0;
  }

  .info-box {
    background: #f7f4ee;
    border-left: 3px solid var(--gold);
    border-radius: 4px;
    padding: 14px 16px;
    font-size: 13px; color: var(--muted);
    line-height: 1.6;
    margin-bottom: 28px;
  }
  .info-box strong { color: var(--ink); }

  /* ── Button ── */
  .btn {
    width: 100%;
    padding: 15px;
    background: var(--ink);
    color: #fff;
    font-family: 'DM Sans', sans-serif;
    font-size: 14px; font-weight: 600; letter-spacing: 1px;
    text-transform: uppercase;
    border: none; border-radius: 8px;
    cursor: pointer;
    transition: background .2s, transform .1s;
    display: flex; align-items: center; justify-content: center; gap: 10px;
  }
  .btn:hover { background: #1e2f3d; }
  .btn:active { transform: scale(.98); }
  .btn:disabled { background: #9aafbf; cursor: not-allowed; transform: none; }

  /* ── Result ── */
  .result {
    display: none;
    margin-top: 28px;
    background: #f0f8f4;
    border: 1.5px solid #a8d8bc;
    border-radius: 10px;
    padding: 24px;
  }
  .result.error {
    background: #fdf2f2;
    border-color: #e8b0b0;
  }
  .result-label {
    font-size: 11px; font-weight: 600; letter-spacing: 1.5px;
    text-transform: uppercase; color: var(--green);
    margin-bottom: 12px;
  }
  .result.error .result-label { color: var(--red); }
  .cb-pair {
    display: flex; gap: 12px; margin-bottom: 18px; flex-wrap: wrap;
  }
  .cb-badge {
    background: var(--ink);
    color: var(--gold2);
    font-family: 'DM Serif Display', serif;
    font-size: 18px;
    padding: 8px 18px;
    border-radius: 6px;
    letter-spacing: 1.5px;
  }
  .dl-btn {
    display: inline-flex; align-items: center; gap: 8px;
    padding: 11px 22px;
    background: var(--green);
    color: #fff;
    font-family: 'DM Sans', sans-serif;
    font-size: 13px; font-weight: 600;
    border-radius: 7px;
    text-decoration: none;
    transition: opacity .2s;
  }
  .dl-btn:hover { opacity: .88; }

  /* ── Spinner ── */
  .spinner {
    width: 18px; height: 18px;
    border: 2.5px solid rgba(255,255,255,.3);
    border-top-color: #fff;
    border-radius: 50%;
    animation: spin .7s linear infinite;
    display: none;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* ── Footer ── */
  footer {
    position: fixed; bottom: 0; left: 0; right: 0;
    background: var(--ink);
    padding: 10px 24px;
    display: flex; align-items: center; justify-content: center;
    gap: 20px; font-size: 12px; color: #8899aa;
  }
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

    <button class="btn" id="genBtn" onclick="generate()">
      <div class="spinner" id="spin"></div>
      <span id="btnTxt">⚡ Generate &amp; Download Form</span>
    </button>

    <div class="result" id="result">
      <div class="result-label" id="resultLabel">Form Generated</div>
      <div class="cb-pair" id="cbPair"></div>
      <a href="#" class="dl-btn" id="dlLink" download>
        ↓ Download Stamped PDF
      </a>
    </div>
  </div>
</main>

<footer>
  <span>Numbers issued: Central Counter</span>
  <a href="/admin/mis?key={{ admin_key }}" target="_blank">Admin: View Today's MIS</a>
</footer>

<script>
async function generate() {
  const empId = document.getElementById('empId').value.trim();
  if (!empId) {
    alert('Please enter your Employee ID before generating.');
    document.getElementById('empId').focus();
    return;
  }

  const btn  = document.getElementById('genBtn');
  const spin = document.getElementById('spin');
  const txt  = document.getElementById('btnTxt');
  const res  = document.getElementById('result');

  btn.disabled = true;
  spin.style.display = 'block';
  txt.textContent = 'Generating...';
  res.style.display = 'none';

  try {
    const resp = await fetch('/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ employee_id: empId })
    });

    const data = await resp.json();

    if (!resp.ok || data.error) {
      throw new Error(data.error || 'Server error');
    }

    // Show numbers
    document.getElementById('cbPair').innerHTML =
      `<div class="cb-badge">${data.cb1}</div><div class="cb-badge">${data.cb2}</div>`;

    // Set download link
    const dlLink = document.getElementById('dlLink');
    dlLink.href = `/download/${data.token}`;
    dlLink.download = `ApplicationForm_${data.cb1}_${data.cb2}.pdf`;

    document.getElementById('resultLabel').textContent = '✓ Form Generated Successfully';
    res.className = 'result';
    res.style.display = 'block';

    // Auto-download
    dlLink.click();

  } catch (err) {
    document.getElementById('resultLabel').textContent = '✗ Error';
    document.getElementById('cbPair').innerHTML =
      `<span style="color:var(--red);font-size:14px">${err.message}</span>`;
    res.className = 'result error';
    res.style.display = 'block';
  } finally {
    btn.disabled = false;
    spin.style.display = 'none';
    txt.textContent = '⚡ Generate & Download Form';
  }
}

document.getElementById('empId').addEventListener('keydown', e => {
  if (e.key === 'Enter') generate();
});
</script>
</body>
</html>
"""

# Temporary in-memory store for generated PDFs (5-minute TTL)
import time
_pdf_cache = {}  # token -> (pdf_bytes, expiry)
_cache_lock = threading.Lock()


def cache_pdf(token, pdf_bytes):
    expiry = time.time() + 300   # 5 minutes
    with _cache_lock:
        _pdf_cache[token] = (pdf_bytes, expiry)
        # Prune old entries
        now = time.time()
        expired = [k for k, (_, exp) in _pdf_cache.items() if exp < now]
        for k in expired:
            del _pdf_cache[k]


def get_cached_pdf(token):
    with _cache_lock:
        entry = _pdf_cache.get(token)
        if not entry:
            return None
        pdf_bytes, expiry = entry
        if time.time() > expiry:
            del _pdf_cache[token]
            return None
        return pdf_bytes


import secrets


@app.route("/")
def index():
    return render_template_string(HTML_PAGE, admin_key=ADMIN_KEY)


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json(silent=True) or {}
    employee_id = (data.get("employee_id") or "").strip()

    if not employee_id:
        return jsonify({"error": "Employee ID is required."}), 400
    if len(employee_id) > 20:
        return jsonify({"error": "Employee ID too long (max 20 chars)."}), 400

    try:
        cb1, cb2 = atomic_next_pair()
        pdf_bytes = stamp_pdf(cb1, cb2)
        token = secrets.token_urlsafe(24)
        cache_pdf(token, pdf_bytes)
        log_generation(employee_id, cb1, cb2, request.remote_addr)
        return jsonify({"cb1": cb1, "cb2": cb2, "token": token})
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 500
    except Exception as e:
        app.logger.error(f"Generate error: {e}")
        return jsonify({"error": "Internal server error. Please try again."}), 500


@app.route("/download/<token>")
def download(token):
    pdf_bytes = get_cached_pdf(token)
    if not pdf_bytes:
        abort(410, "Link expired. Please generate a new form.")
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name="ApplicationForm.pdf"
    )


# ─── Admin routes (protected by key) ───────────────────────────

def require_admin(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        key = request.args.get("key") or request.headers.get("X-Admin-Key")
        if key != ADMIN_KEY:
            abort(403, "Invalid admin key.")
        return f(*args, **kwargs)
    return wrapper


@app.route("/admin/mis")
@require_admin
def admin_mis():
    """Download today's MIS as CSV (generates on demand too)."""
    filepath = generate_mis()
    return send_file(filepath, as_attachment=True,
                     download_name=os.path.basename(filepath),
                     mimetype="text/csv")


@app.route("/admin/mis/<date_str>")
@require_admin
def admin_mis_date(date_str):
    """Download MIS for a specific date: /admin/mis/2024-03-15"""
    try:
        target = datetime.date.fromisoformat(date_str)
    except ValueError:
        abort(400, "Date format must be YYYY-MM-DD")
    filepath = generate_mis(target)
    return send_file(filepath, as_attachment=True,
                     download_name=os.path.basename(filepath),
                     mimetype="text/csv")


@app.route("/admin/stats")
@require_admin
def admin_stats():
    """JSON endpoint: current counter value + today's stats."""
    tz = ZoneInfo(TIMEZONE)
    today = datetime.datetime.now(tz).date().isoformat()
    with get_db() as conn:
        counter = conn.execute("SELECT value FROM counter WHERE id=1").fetchone()["value"]
        total_today = conn.execute(
            "SELECT COUNT(*) as c FROM logs WHERE DATE(generated_at)=?", (today,)
        ).fetchone()["c"]
        top_users = conn.execute(
            "SELECT employee_id, COUNT(*) as forms FROM logs "
            "WHERE DATE(generated_at)=? GROUP BY employee_id ORDER BY forms DESC LIMIT 10",
            (today,)
        ).fetchall()
    return jsonify({
        "last_number_issued": counter,
        "last_cb_number": f"CB{counter:08d}",
        "forms_generated_today": total_today,
        "numbers_issued_today": total_today * 2,
        "top_users_today": [{"employee_id": r["employee_id"], "forms": r["forms"]} for r in top_users]
    })


# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    init_db()
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
