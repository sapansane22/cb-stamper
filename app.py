"""
CB Application Form Stamper — v3
=================================
CHANGES IN THIS VERSION:
  a) Numbers stamped on EVERY page of the PDF at the same X/Y position
  b) Series: CB + 9 digits starting from CB500000001
             (counter seeds at 500_000_000, increments by 2 each generation)
  c) Three templates — Non Defence / Prahri / Param
     (template files: template_nondefence.pdf, template_prahri.pdf, template_param.pdf)
     Single GLOBAL counter — numbers are unique across ALL three templates,
     so the same CB number can never appear on any template twice.
     MIS report now includes a 'Template' column.
"""

import os
import io
import csv
import base64
import sqlite3
import threading
import datetime
from functools import wraps
from zoneinfo import ZoneInfo

from flask import Flask, request, jsonify, send_file, render_template_string, abort
from apscheduler.schedulers.background import BackgroundScheduler
from reportlab.pdfgen import canvas as rl_canvas
from pypdf import PdfReader, PdfWriter

# ─────────────────────────── CONFIG ────────────────────────────
DB_PATH    = "stamper.db"
ADMIN_KEY  = os.environ.get("ADMIN_KEY", "admin1234")
TIMEZONE   = "Asia/Kolkata"
MIS_FOLDER = "mis_reports"

# ── CHANGE (b): 9-digit series starting from 500000001 ──────────
CB_START  = 500_000_000        # counter seeds here; first pair → 500000001 / 500000002
CB_FORMAT = "CB{:09d}"        # CB + 9 digits → CB500000001

# ── CHANGE (c): Template definitions ────────────────────────────
# Key   = internal ID sent from browser
# label = display name shown to user
# file  = PDF filename (upload these alongside app.py on the server)
TEMPLATES = {
    "nondefence": {
        "label": "Non Defence",
        "file":  "template_nondefence.pdf",
    },
    "prahri": {
        "label": "Prahri",
        "file":  "template_prahri.pdf",
    },
    "param": {
        "label": "Param",
        "file":  "template_param.pdf",
    },
}

# ── Stamp positions (points from bottom-left of page) ───────────
# A4 = 595 × 842 pt  |  Letter = 612 × 792 pt  |  1 inch = 72 pt
# CHANGE (a): same X/Y is applied to EVERY page in the document.
NUM1_X, NUM1_Y = 60, 835
NUM2_X, NUM2_Y = 280, 835
FONT_NAME      = "Helvetica-Bold"
FONT_SIZE      = 10
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
    """Create tables and seed the counter at CB_START. Safe to call repeatedly."""
    conn = get_db()
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS counter (
                id    INTEGER PRIMARY KEY CHECK (id = 1),
                value INTEGER NOT NULL DEFAULT 0
            )
        """)
        # Seed at CB_START so first issued number is CB_START+1
        conn.execute(
            "INSERT OR IGNORE INTO counter (id, value) VALUES (1, ?)",
            (CB_START,)
        )
        conn.execute("""
            CREATE TABLE IF NOT EXISTS logs (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id  TEXT NOT NULL,
                template_key TEXT NOT NULL,
                cb_num_1     TEXT NOT NULL,
                cb_num_2     TEXT NOT NULL,
                generated_at TEXT NOT NULL,
                ip_address   TEXT
            )
        """)
        conn.commit()
    finally:
        conn.close()


# Called at import time so gunicorn initialises the DB on startup
init_db()


def atomic_next_pair():
    """
    Atomically claim the next two sequential CB numbers from the GLOBAL counter.
    Thread-locked + SQLite IMMEDIATE transaction — safe under full concurrency.
    Returns e.g. ("CB500000001", "CB500000002")
    """
    with _db_lock:
        conn = get_db()
        try:
            conn.execute("BEGIN IMMEDIATE")
            current = conn.execute(
                "SELECT value FROM counter WHERE id=1"
            ).fetchone()["value"]
            conn.execute(
                "UPDATE counter SET value=? WHERE id=1", (current + 2,)
            )
            conn.execute("COMMIT")
        except Exception:
            try:
                conn.execute("ROLLBACK")
            except Exception:
                pass
            conn.close()
            raise
        finally:
            conn.close()

    return (CB_FORMAT.format(current + 1), CB_FORMAT.format(current + 2))


def log_generation(employee_id, template_key, cb1, cb2, ip):
    tz = ZoneInfo(TIMEZONE)
    ts = datetime.datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    try:
        conn.execute(
            "INSERT INTO logs "
            "(employee_id, template_key, cb_num_1, cb_num_2, generated_at, ip_address) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (employee_id.strip().upper(), template_key, cb1, cb2, ts, ip)
        )
        conn.commit()
    finally:
        conn.close()


# ═══════════════════════════════════════════════════════════════
#  PDF STAMPING
# ═══════════════════════════════════════════════════════════════

def stamp_pdf(template_key: str, cb1: str, cb2: str) -> bytes:
    """
    CHANGE (a): builds one overlay per page and merges it onto EVERY page
                so both CB numbers appear at the same position on each page.
    CHANGE (b): number format is now CB + 9 digits (handled by atomic_next_pair).
    CHANGE (c): selects the correct template file based on template_key.
    """
    tpl = TEMPLATES.get(template_key)
    if tpl is None:
        raise ValueError(f"Unknown template key: {template_key}")

    pdf_path = tpl["file"]
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(
            f"Template file '{pdf_path}' was not found on the server. "
            f"Please upload it alongside app.py."
        )

    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    for page in reader.pages:
        pw = float(page.mediabox.width)
        ph = float(page.mediabox.height)

        # Build a fresh overlay for this page's dimensions
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
        writer.add_page(page)

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
            "SELECT employee_id, template_key, cb_num_1, cb_num_2, "
            "generated_at, ip_address "
            "FROM logs WHERE DATE(generated_at) = ? ORDER BY id",
            (date_str,)
        ).fetchall()
    finally:
        conn.close()

    # Per-employee summary (total forms, breakdown by template)
    emp_summary: dict[str, dict] = {}
    for r in rows:
        eid  = r["employee_id"]
        tkey = r["template_key"]
        tlab = TEMPLATES.get(tkey, {}).get("label", tkey)
        if eid not in emp_summary:
            emp_summary[eid] = {"total": 0, "by_template": {}}
        emp_summary[eid]["total"] += 1
        emp_summary[eid]["by_template"][tlab] = \
            emp_summary[eid]["by_template"].get(tlab, 0) + 1

    # Per-template summary
    tpl_summary: dict[str, int] = {}
    for r in rows:
        tlab = TEMPLATES.get(r["template_key"], {}).get("label", r["template_key"])
        tpl_summary[tlab] = tpl_summary.get(tlab, 0) + 1

    with open(filepath, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)

        # Header
        w.writerow(["CB APPLICATION FORM — MIS REPORT"])
        w.writerow(["Date:", date_str])
        w.writerow(["Generated at:", datetime.datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")])
        w.writerow(["Total Forms Generated:", len(rows)])
        w.writerow(["Total CB Numbers Issued:", len(rows) * 2])
        w.writerow([])

        # Template breakdown
        w.writerow(["TEMPLATE BREAKDOWN"])
        w.writerow(["Template", "Forms Generated", "CB Numbers Used"])
        for tlab, count in sorted(tpl_summary.items()):
            w.writerow([tlab, count, count * 2])
        w.writerow([])

        # Employee summary
        w.writerow(["EMPLOYEE SUMMARY"])
        w.writerow(["Employee ID", "Total Forms", "CB Numbers Used",
                    "Non Defence", "Prahri", "Param"])
        for eid, data in sorted(emp_summary.items()):
            bt = data["by_template"]
            w.writerow([
                eid,
                data["total"],
                data["total"] * 2,
                bt.get("Non Defence", 0),
                bt.get("Prahri", 0),
                bt.get("Param", 0),
            ])
        w.writerow([])

        # Detail log
        w.writerow(["DETAIL LOG"])
        w.writerow(["#", "Employee ID", "Template", "CB Number 1",
                    "CB Number 2", "Generated At", "IP Address"])
        for i, r in enumerate(rows, 1):
            tlab = TEMPLATES.get(r["template_key"], {}).get("label", r["template_key"])
            w.writerow([i, r["employee_id"], tlab, r["cb_num_1"],
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

HTML_PAGE = """
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
    font-family: 'DM Serif Display', serif; font-size: 20px; color: var(--ink); flex-shrink: 0;
  }
  header h1 { font-family: 'DM Serif Display', serif; font-size: 20px; color: #fff; font-weight: 400; }
  header span { color: var(--gold2); }

  main { width: 100%; max-width: 580px; padding: 48px 20px 90px; }
  .eyebrow { font-size: 11px; font-weight: 600; letter-spacing: 2.5px; text-transform: uppercase; color: var(--gold); margin-bottom: 10px; }
  .headline { font-family: 'DM Serif Display', serif; font-size: 34px; line-height: 1.15; margin-bottom: 8px; }
  .headline em { font-style: italic; color: var(--muted); }
  .sub { font-size: 14px; color: var(--muted); line-height: 1.6; margin-bottom: 36px; }

  .card { background: #fff; border: 1px solid var(--line); border-radius: 12px; padding: 36px; box-shadow: 0 4px 32px rgba(15,25,35,0.06); }

  /* ── Template selector ── */
  .field { margin-bottom: 24px; }
  label.field-label {
    display: block; font-size: 12px; font-weight: 600;
    letter-spacing: 1.2px; text-transform: uppercase; color: var(--muted); margin-bottom: 10px;
  }
  .tpl-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; }
  .tpl-card {
    border: 2px solid var(--line); border-radius: 9px; padding: 14px 10px;
    text-align: center; cursor: pointer;
    transition: border-color .18s, background .18s, transform .12s;
    user-select: none;
  }
  .tpl-card:hover { border-color: var(--gold); background: #fdf9f2; }
  .tpl-card.selected {
    border-color: var(--gold); background: #fdf2d8;
    box-shadow: 0 0 0 3px rgba(200,151,58,.15);
  }
  .tpl-card:active { transform: scale(.97); }
  .tpl-icon { font-size: 22px; margin-bottom: 6px; }
  .tpl-name { font-size: 13px; font-weight: 600; color: var(--ink); }
  .tpl-card.selected .tpl-name { color: #8a5a0a; }

  input[type="text"] {
    width: 100%; padding: 13px 16px;
    font-family: 'DM Sans', sans-serif; font-size: 15px; color: var(--ink);
    background: var(--paper); border: 1.5px solid var(--line); border-radius: 8px;
    outline: none; transition: border-color .2s, box-shadow .2s; letter-spacing: 1px;
  }
  input[type="text"]:focus { border-color: var(--gold); box-shadow: 0 0 0 3px rgba(200,151,58,.12); }
  input[type="text"]::placeholder { color: #c0b9ae; }

  .divider { height: 1px; background: var(--line); margin: 24px 0; }

  .info-box {
    background: #f7f4ee; border-left: 3px solid var(--gold); border-radius: 4px;
    padding: 13px 15px; font-size: 13px; color: var(--muted); line-height: 1.6; margin-bottom: 26px;
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

  .result { display: none; margin-top: 26px; background: #f0f8f4; border: 1.5px solid #a8d8bc; border-radius: 10px; padding: 22px; }
  .result.error { background: #fdf2f2; border-color: #e8b0b0; }
  .result-label { font-size: 11px; font-weight: 600; letter-spacing: 1.5px; text-transform: uppercase; color: var(--green); margin-bottom: 10px; }
  .result.error .result-label { color: var(--red); }
  .result-meta { font-size: 12px; color: var(--muted); margin-bottom: 12px; }
  .cb-pair { display: flex; gap: 10px; margin-bottom: 16px; flex-wrap: wrap; }
  .cb-badge {
    background: var(--ink); color: var(--gold2);
    font-family: 'DM Serif Display', serif; font-size: 17px;
    padding: 8px 16px; border-radius: 6px; letter-spacing: 1.5px;
  }
  .dl-btn {
    display: inline-flex; align-items: center; gap: 8px; padding: 11px 22px;
    background: var(--green); color: #fff; font-family: 'DM Sans', sans-serif;
    font-size: 13px; font-weight: 600; border-radius: 7px;
    border: none; cursor: pointer; transition: opacity .2s;
  }
  .dl-btn:hover { opacity: .88; }

  .spinner { width: 18px; height: 18px; border: 2.5px solid rgba(255,255,255,.3); border-top-color: #fff; border-radius: 50%; animation: spin .7s linear infinite; display: none; }
  @keyframes spin { to { transform: rotate(360deg); } }

  footer { position: fixed; bottom: 0; left: 0; right: 0; background: var(--ink); padding: 10px 24px; display: flex; align-items: center; justify-content: center; gap: 20px; font-size: 12px; color: #8899aa; flex-wrap: wrap; }
  footer a { color: var(--gold2); text-decoration: none; }
  footer a:hover { text-decoration: underline; }

  @media (max-width: 480px) {
    .tpl-grid { grid-template-columns: 1fr; }
    .card { padding: 24px; }
  }
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
  <p class="sub">Select a product, enter your Employee ID, and download your uniquely numbered form. All numbers are globally sequential and never repeat across products.</p>

  <div class="card">

    <!-- Step 1: Template selector -->
    <div class="field">
      <label class="field-label">Step 1 — Select Product</label>
      <div class="tpl-grid">
        <div class="tpl-card" id="tpl-nondefence" onclick="selectTemplate('nondefence')">
          <div class="tpl-icon">🛡️</div>
          <div class="tpl-name">Non Defence</div>
        </div>
        <div class="tpl-card" id="tpl-prahri" onclick="selectTemplate('prahri')">
          <div class="tpl-icon">⚔️</div>
          <div class="tpl-name">Prahri</div>
        </div>
        <div class="tpl-card" id="tpl-param" onclick="selectTemplate('param')">
          <div class="tpl-icon">🏆</div>
          <div class="tpl-name">Param</div>
        </div>
      </div>
    </div>

    <div class="divider"></div>

    <!-- Step 2: Employee ID -->
    <div class="field">
      <label class="field-label" for="empId">Step 2 — Employee ID</label>
      <input type="text" id="empId" placeholder="e.g. EMP1042" maxlength="20" autocomplete="off">
    </div>

    <div class="info-box">
      <strong>Two globally unique numbers per form:</strong><br>
      e.g. <strong>CB500000001</strong> &amp; <strong>CB500000002</strong> — the same number will
      never appear again on any product's form.
    </div>

    <button class="btn" id="genBtn" onclick="generateForm()">
      <div class="spinner" id="spin"></div>
      <span id="btnTxt">⚡ Generate &amp; Download Form</span>
    </button>

    <div class="result" id="result">
      <div class="result-label" id="resultLabel">Form Generated</div>
      <div class="result-meta" id="resultMeta"></div>
      <div class="cb-pair" id="cbPair"></div>
      <button class="dl-btn" id="dlBtn" onclick="triggerDownload()">
        ↓ Download Stamped PDF
      </button>
    </div>
  </div>
</main>

<footer>
  <span>Global counter — unique across all products</span>
  <a href="/admin/mis?key={{ admin_key }}" target="_blank">Admin: Today's MIS</a>
  <a href="/admin/stats?key={{ admin_key }}" target="_blank">Admin: Live Stats</a>
</footer>

<script>
let _selectedTemplate = null;
let _pdfB64 = null, _cb1 = '', _cb2 = '';

function selectTemplate(key) {
  _selectedTemplate = key;
  document.querySelectorAll('.tpl-card').forEach(el => el.classList.remove('selected'));
  document.getElementById('tpl-' + key).classList.add('selected');
}

async function generateForm() {
  if (!_selectedTemplate) {
    alert('Please select a product before generating.');
    return;
  }
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
  _pdfB64 = null;

  try {
    const resp = await fetch('/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ employee_id: empId, template: _selectedTemplate })
    });

    const data = await resp.json();
    if (!resp.ok || data.error) throw new Error(data.error || `Server error (${resp.status})`);

    _pdfB64 = data.pdf_b64;
    _cb1    = data.cb1;
    _cb2    = data.cb2;

    const tplNames = { nondefence: 'Non Defence', prahri: 'Prahri', param: 'Param' };
    document.getElementById('resultMeta').textContent =
      'Product: ' + (tplNames[_selectedTemplate] || _selectedTemplate);
    document.getElementById('cbPair').innerHTML =
      `<div class="cb-badge">${data.cb1}</div><div class="cb-badge">${data.cb2}</div>`;
    document.getElementById('resultLabel').textContent = '✓ Form Generated Successfully';
    document.getElementById('dlBtn').style.display = '';
    res.className = 'result';
    res.style.display = 'block';

    triggerDownload();

  } catch (err) {
    document.getElementById('resultLabel').textContent = '✗ Error';
    document.getElementById('resultMeta').textContent = '';
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
  const bytes = atob(_pdfB64);
  const arr   = new Uint8Array(bytes.length);
  for (let i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i);
  const blob   = new Blob([arr], { type: 'application/pdf' });
  const url    = URL.createObjectURL(blob);
  const a      = document.createElement('a');
  a.href       = url;
  a.download   = `ApplicationForm_${_cb1}_${_cb2}.pdf`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
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
    data         = request.get_json(silent=True) or {}
    employee_id  = (data.get("employee_id") or "").strip()
    template_key = (data.get("template") or "").strip().lower()

    if not employee_id:
        return jsonify({"error": "Employee ID is required."}), 400
    if len(employee_id) > 20:
        return jsonify({"error": "Employee ID too long (max 20 chars)."}), 400
    if template_key not in TEMPLATES:
        return jsonify({"error": "Please select a valid product template."}), 400

    try:
        cb1, cb2  = atomic_next_pair()
        pdf_bytes = stamp_pdf(template_key, cb1, cb2)
        pdf_b64   = base64.b64encode(pdf_bytes).decode("ascii")
        log_generation(employee_id, template_key, cb1, cb2, request.remote_addr)
        return jsonify({"cb1": cb1, "cb2": cb2, "pdf_b64": pdf_b64})

    except (FileNotFoundError, ValueError) as e:
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
    filepath = generate_mis()
    return send_file(filepath, as_attachment=True,
                     download_name=os.path.basename(filepath),
                     mimetype="text/csv")


@app.route("/admin/mis/<date_str>")
@require_admin
def admin_mis_date(date_str):
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
        by_template = conn.execute(
            "SELECT template_key, COUNT(*) AS forms FROM logs "
            "WHERE DATE(generated_at)=? GROUP BY template_key ORDER BY forms DESC",
            (today,)
        ).fetchall()
        top_users = conn.execute(
            "SELECT employee_id, COUNT(*) AS forms FROM logs "
            "WHERE DATE(generated_at)=? GROUP BY employee_id ORDER BY forms DESC LIMIT 10",
            (today,)
        ).fetchall()
    finally:
        conn.close()

    return jsonify({
        "last_cb_number_issued":  CB_FORMAT.format(counter),
        "forms_generated_today":  total_today,
        "numbers_issued_today":   total_today * 2,
        "by_template_today": [
            {
                "template": TEMPLATES.get(r["template_key"], {}).get("label", r["template_key"]),
                "forms": r["forms"]
            }
            for r in by_template
        ],
        "top_users_today": [
            {"employee_id": r["employee_id"], "forms": r["forms"]}
            for r in top_users
        ]
    })


# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
