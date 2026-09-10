"""
tMart Rebate DN Dashboard — Cloud Edition
==========================================
• Runs on Render.com (free tier) with LibreOffice for PDF conversion
• Each colleague gets their own private session
• Email sent via their own Gmail (App Password)
• Upload Excel (Sheet1 used), generate PDFs, download, send email
"""

import os, sys, re, io, json, uuid, shutil, tempfile, zipfile, threading, datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text       import MIMEText
from email.mime.application import MIMEApplication
from flask import Flask, jsonify, send_file, request, session, Response

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "tmart-dn-cloud-change-this-secret")

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
SESSIONS_DIR = os.path.join(tempfile.gettempdir(), "tmart_sessions")
os.makedirs(SESSIONS_DIR, exist_ok=True)

# CC and email always come from the uploaded Excel — no hardcoded values

# ── Session helpers ────────────────────────────────────────────────────────────

def get_sid():
    if "sid" not in session:
        session["sid"] = str(uuid.uuid4())[:12]
    return session["sid"]

def sdir(sid=None):
    d = os.path.join(SESSIONS_DIR, sid or get_sid())
    os.makedirs(d, exist_ok=True)
    return d

def sfile(name, sid=None):   return os.path.join(sdir(sid), name)
def outdir(sid=None):
    d = sfile("output", sid)
    os.makedirs(d, exist_ok=True)
    return d
def excel_path(sid=None):  return sfile("DN_DATA.xlsx", sid)
def template_path(sid=None):
    p = sfile("template.docx", sid)
    if os.path.exists(p): return p
    g = os.path.join(BASE_DIR, "DN_INVOICE _FORMAT.docx")
    return g if os.path.exists(g) else None

# ── Value helpers ──────────────────────────────────────────────────────────────

def str_val(v):
    if v is None: return ""
    if isinstance(v, (datetime.datetime, datetime.date)): return v.strftime("%d.%m.%Y")
    if isinstance(v, float): return str(int(v)) if v == int(v) else f"{v:.3f}".rstrip("0").rstrip(".")
    if isinstance(v, int): return str(v)
    return str(v).strip()

def col(row, name):
    if name in row: return row[name]
    ns = name.strip()
    if ns in row: return row[ns]
    nl = ns.lower()
    for k, v in row.items():
        if k.strip().lower() == nl: return v
    return ""

def pdf_name(row):
    si = str(row.get("SI No") or row.get("_row_idx", "")).strip()
    fn = col(row, "File Name") or col(row, "Supplier name") or "unknown"
    fn = re.sub(r'[\\/:*?"<>|]', "_", str(fn)).strip()
    return f"{si}_{fn}.pdf"

# ── Excel reading ──────────────────────────────────────────────────────────────

def read_excel(sid=None):
    ep = excel_path(sid)
    if not os.path.exists(ep):
        return {"error": "No Excel uploaded — click 📂 Upload Excel"}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ep, data_only=True)
        ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.active
        headers = [str(c.value).strip() if c.value else f"col{i}" for i, c in enumerate(ws[1])]
        rows = []
        for idx, r in enumerate(ws.iter_rows(min_row=2, values_only=True), 1):
            if all(v is None for v in r): continue
            row = {h: str_val(v) for h, v in zip(headers, r)}
            row["_row_idx"] = str(idx)
            rows.append(row)
        wb.close()
        return rows
    except Exception as e:
        return {"error": str(e)}

# ── Template filling ───────────────────────────────────────────────────────────

def fill_template(template, output, row):
    from docx import Document
    total_str = col(row, "Total rebate amount")
    try:    is_credit = float(total_str) < 0
    except: is_credit = False

    reps = {
        "<<BREPLACE>>": col(row, "Invoice Date"),
        "<<CREPLACE>>": col(row, "Finance ID"),
        "<<DREPLACE>>": col(row, "Supplier name"),
        "<<EREPLACE>>": col(row, "Exlusive -0%"),
        "<<FREPLACE>>": col(row, "Exlusive-10%"),
        "<<GREPLACE>>": col(row, "VAT Amount"),
        "<<HREPLACE>>": col(row, "Ex 10%+Vat"),
        "<<IREPLACE>>": col(row, "T_rebate_excl"),
        "<<JREPLACE>>": col(row, "Total rebate amount"),
        "<<KREPLACE>>": col(row, "Debit note number"),
        "<<LREPLACE>>": col(row, "Description"),
        "<<MREPLACE>>": col(row, "VAT No."),
        "<<NREPLACE>>": col(row, "Address"),
    }

    def process(para):
        full = "".join(r.text for r in para.runs)
        if not full.strip(): return
        changed = False
        for old, new in reps.items():
            if old in full:
                full = full.replace(old, str(new) if new else "")
                changed = True
        if is_credit:
            for dn, cn in (("DEBIT NOTE","CREDIT NOTE"),("Debit Note","Credit Note"),("Debit note","Credit Note")):
                if dn in full: full = full.replace(dn, cn); changed = True
        cleaned = re.sub(r"<<[A-Z]+REPLACE>>", "", full)
        if cleaned != full: full = cleaned; changed = True
        if changed and para.runs:
            para.runs[0].text = full
            for r in para.runs[1:]: r.text = ""

    doc = Document(template)
    for p in doc.paragraphs: process(p)
    for tbl in doc.tables:
        for tr in tbl.rows:
            for cell in tr.cells:
                for p in cell.paragraphs: process(p)
    doc.save(output)

# ── PDF conversion ─────────────────────────────────────────────────────────────

def convert_to_pdf(docx_path, pdf_path):
    """LibreOffice (cloud/Linux) → Word COM fallback (Windows)."""
    lo = shutil.which("libreoffice") or shutil.which("soffice")
    if lo:
        tmp = tempfile.mkdtemp()
        try:
            r = subprocess_run_safe([lo, "--headless", "--convert-to", "pdf",
                                     "--outdir", tmp, os.path.abspath(docx_path)])
            base = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
            src  = os.path.join(tmp, base)
            if os.path.exists(src):
                shutil.move(src, pdf_path)
                return True, None
            return False, r.get("stderr") or "LibreOffice: PDF not created"
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    # Windows fallback — Word COM via PowerShell
    td = docx_path.replace('"', '`"')
    op = pdf_path.replace('"',  '`"')
    ps = (f'$w=New-Object -ComObject Word.Application;$w.Visible=$false;$w.DisplayAlerts=0;'
          f'$d=$w.Documents.Open("{td}");$d.ExportAsFixedFormat("{op}",17);'
          f'$d.Close($false);$w.Quit()')
    r = subprocess_run_safe(["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps])
    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
        return True, None
    return False, r.get("stderr") or "Word COM: PDF not created"

import subprocess
def subprocess_run_safe(cmd, timeout=120):
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        return {"returncode": r.returncode, "stdout": r.stdout, "stderr": r.stderr}
    except Exception as e:
        return {"returncode": -1, "stdout": "", "stderr": str(e)}

# ── Single PDF generation ──────────────────────────────────────────────────────

def generate_one(row, sid=None):
    tmpl = template_path(sid)
    if not tmpl:
        return {"pdf_name": pdf_name(row), "ok": False,
                "error": "No template — upload DN_INVOICE _FORMAT.docx"}
    pn       = pdf_name(row)
    od       = outdir(sid)
    out_pdf  = os.path.join(od, pn)
    temp_doc = os.path.join(od, pn.replace(".pdf", "_filled.docx"))
    try:
        fill_template(tmpl, temp_doc, row)
        ok, err = convert_to_pdf(temp_doc, out_pdf)
        if ok and os.path.exists(out_pdf) and os.path.getsize(out_pdf) > 0:
            return {"pdf_name": pn, "ok": True,  "error": None}
        return  {"pdf_name": pn, "ok": False, "error": err or "PDF empty"}
    except Exception as e:
        return  {"pdf_name": pn, "ok": False, "error": str(e)}
    finally:
        try:
            if os.path.exists(temp_doc): os.remove(temp_doc)
        except: pass

# ── Background batch generation ────────────────────────────────────────────────

_jobs = {}   # sid → status dict

def _run_batch(sid, rows_list):
    status = _jobs[sid]
    for row in rows_list:
        if status.get("cancelled"): break
        r = generate_one(row, sid=sid)   # pass sid — no Flask session in thread
        status["done"]   += 1
        status["results"].append(r)
        if not r["ok"]: status["errors"].append(r)
        _write_job(sid, status)
    status["running"] = False
    _write_job(sid, status)

def _write_job(sid, status):
    path = os.path.join(SESSIONS_DIR, sid, "job.json")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(status, f)

def _read_job(sid):
    path = os.path.join(SESSIONS_DIR, sid, "job.json")
    try:
        with open(path) as f: return json.load(f)
    except:
        return {"total":0,"done":0,"running":False,"results":[],"errors":[]}

# ── Gmail SMTP ─────────────────────────────────────────────────────────────────

def send_gmail(to, cc, subject, body_text, attachment_path, gmail_user, gmail_pass):
    msg = MIMEMultipart()
    msg["From"]    = gmail_user
    msg["To"]      = to
    msg["CC"]      = cc
    msg["Subject"] = subject
    msg.attach(MIMEText(body_text, "plain"))
    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(attachment_path)}"'
        msg.attach(part)
    recipients = [a.strip() for a in (to + ";" + cc).split(";") if a.strip()]
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(gmail_user, gmail_pass)
        s.sendmail(gmail_user, recipients, msg.as_string())

# ── API routes ─────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return DASHBOARD_HTML

@app.route("/api/session-id")
def api_session_id():
    return jsonify({"sid": get_sid()})

@app.route("/api/rows")
def api_rows():
    rows = read_excel()
    if isinstance(rows, dict):
        return jsonify({"ok": False, "error": rows["error"]})
    od = outdir()
    sent = session.get("sent", {})
    result = []
    for i, row in enumerate(rows):
        pn = pdf_name(row)
        result.append({
            "idx":        i,
            "si_no":      row.get("_row_idx", str(i+1)),
            "supplier":   col(row, "Supplier name"),
            "dn_number":  col(row, "Debit note number"),
            "email":      col(row, "Email") or col(row, "email"),
            "cc":         col(row, "cc") or col(row, "CC") or col(row, "Cc") or "",
            "pdf_name":   pn,
            "pdf_exists": os.path.exists(os.path.join(od, pn)),
            "email_sent": sent.get(pn, False),
            "row":        row,
        })
    return jsonify({"ok": True, "rows": result})

@app.route("/api/generate-start", methods=["POST"])
def api_generate_start():
    body      = request.json or {}
    rows_list = [item["row"] for item in body.get("rows", [])]
    sid       = get_sid()
    status    = {"total": len(rows_list), "done": 0, "running": True,
                 "results": [], "errors": [], "cancelled": False}
    _jobs[sid] = status
    _write_job(sid, status)
    t = threading.Thread(target=_run_batch, args=(sid, rows_list), daemon=True)
    t.start()
    return jsonify({"ok": True, "total": len(rows_list)})

@app.route("/api/generate-status")
def api_generate_status():
    return jsonify(_read_job(get_sid()))

@app.route("/api/generate", methods=["POST"])   # single-row quick generate
def api_generate():
    body      = request.json or {}
    rows_list = [item["row"] for item in body.get("rows", [])]
    sid = get_sid()
    results   = [generate_one(r, sid=sid) for r in rows_list]
    return jsonify({"results": results})

@app.route("/api/send-email", methods=["POST"])
def api_send_email():
    body     = request.json or {}
    pn       = body.get("pdf_name", "")
    pdf_path = os.path.join(outdir(), pn)
    if not os.path.exists(pdf_path):
        return jsonify({"ok": False, "error": "PDF not found — generate it first"})
    gu = session.get("gmail_user", "")
    gp = session.get("gmail_pass", "")
    if not gu or not gp:
        return jsonify({"ok": False, "error": "Gmail not set — click ⚙ Gmail Settings"})
    try:
        to       = body.get("email", "")
        cc       = body.get("cc", "")
        supplier = body.get("supplier", "Supplier")
        dn_num   = body.get("dn_number", "")
        subject  = f"Rebate Debit Note {dn_num} - DH Store Bahrain (tMart)".strip()
        body_t   = (f"Dear {supplier},\n\n"
                    "Please find attached your Rebate Debit Note.\n\n"
                    "Regards,\ntMart Finance Team")
        send_gmail(to, cc, subject, body_t, pdf_path, gu, gp)
        sent = session.get("sent", {})
        sent[pn] = True
        session["sent"] = sent
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

@app.route("/api/gmail-settings", methods=["POST"])
def api_gmail_settings():
    body = request.json or {}
    session["gmail_user"] = body.get("user", "").strip()
    session["gmail_pass"] = body.get("password", "").strip()
    return jsonify({"ok": True, "user": session["gmail_user"]})

@app.route("/api/gmail-status")
def api_gmail_status():
    return jsonify({"configured": bool(session.get("gmail_user")),
                    "user": session.get("gmail_user", "")})

@app.route("/api/upload-excel", methods=["POST"])
def api_upload_excel():
    try:
        import openpyxl
        f = request.files.get("file")
        if not f: return jsonify({"ok": False, "error": "No file"})
        dest = excel_path()
        f.save(dest)
        wb = openpyxl.load_workbook(dest, data_only=True)
        sheets = wb.sheetnames; wb.close()
        if "Sheet1" not in sheets:
            return jsonify({"ok": False, "error": f"Sheet1 not found. Sheets available: {sheets}"})
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

@app.route("/api/upload-template", methods=["POST"])
def api_upload_template():
    f = request.files.get("file")
    if not f: return jsonify({"ok": False, "error": "No file"})
    f.save(sfile("template.docx"))
    return jsonify({"ok": True})

@app.route("/api/template-status")
def api_template_status():
    t = template_path()
    return jsonify({"exists": bool(t), "path": os.path.basename(t) if t else ""})

@app.route("/preview/<sid>/<path:filename>")
def preview(sid, filename):
    path = os.path.join(SESSIONS_DIR, sid, "output", filename)
    if os.path.exists(path): return send_file(path, mimetype="application/pdf")
    return jsonify({"error": "Not found"}), 404

@app.route("/download/<sid>/<path:filename>")
def download(sid, filename):
    path = os.path.join(SESSIONS_DIR, sid, "output", filename)
    if os.path.exists(path): return send_file(path, as_attachment=True, download_name=filename)
    return jsonify({"error": "Not found"}), 404

@app.route("/download-all/<sid>")
def download_all(sid):
    od   = os.path.join(SESSIONS_DIR, sid, "output")
    pdfs = [f for f in os.listdir(od) if f.endswith(".pdf")] if os.path.exists(od) else []
    if not pdfs: return jsonify({"error": "No PDFs yet"}), 404
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in pdfs: zf.write(os.path.join(od, f), f)
    buf.seek(0)
    return send_file(buf, mimetype="application/zip",
                     as_attachment=True, download_name="Rebate_DNs.zip")

@app.route("/api/delete/<path:filename>", methods=["DELETE"])
def delete_pdf(filename):
    path = os.path.join(outdir(), filename)
    try:
        if os.path.exists(path): os.remove(path)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

# ── Dashboard HTML ─────────────────────────────────────────────────────────────

DASHBOARD_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>tMart – Rebate DN Dashboard</title>
<style>
:root{--or:#FF6A00;--dark:#1C1C2E;--card:#27273D;--bdr:#3A3A55;--grn:#22C55E;--red:#EF4444;--blu:#3B82F6;--pur:#8B5CF6;--tl:#14B8A6;}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:"Segoe UI",sans-serif;background:var(--dark);color:#E2E2F0;min-height:100vh;}
nav{background:var(--card);border-bottom:3px solid var(--or);padding:14px 28px;display:flex;align-items:center;gap:14px;flex-wrap:wrap;}
.logo{font-size:22px;font-weight:800;color:var(--or);}.sub{font-size:13px;color:#9999CC;flex:1;}
.cards{display:flex;gap:16px;padding:22px 28px 0;flex-wrap:wrap;}
.card{flex:1;min-width:140px;background:var(--card);border:1px solid var(--bdr);border-radius:10px;padding:16px 20px;}
.card .num{font-size:28px;font-weight:700;color:var(--or);}.card .lbl{font-size:11px;color:#9999CC;margin-top:2px;text-transform:uppercase;letter-spacing:.5px;}
.toolbar{display:flex;gap:8px;padding:16px 28px;flex-wrap:wrap;align-items:center;}
input[type=file]{display:none;}
button,label.btn{cursor:pointer;border:none;border-radius:7px;padding:9px 16px;font-size:13px;font-weight:600;transition:opacity .15s;display:inline-flex;align-items:center;gap:5px;}
button:hover,label.btn:hover{opacity:.85;}
.or{background:var(--or);color:#fff;}.grn{background:var(--grn);color:#fff;}.blu{background:var(--blu);color:#fff;}
.pur{background:var(--pur);color:#fff;}.red{background:var(--red);color:#fff;}.tl{background:var(--tl);color:#fff;}
.out{background:transparent;border:1px solid var(--bdr);color:#CCC;}
.sm{padding:4px 10px;font-size:12px;border-radius:5px;}.sep{flex:1;}
.tbl-wrap{padding:0 28px 40px;overflow-x:auto;}
table{width:100%;border-collapse:collapse;font-size:13px;}
thead th{background:var(--card);color:var(--or);padding:10px 12px;text-align:left;border-bottom:2px solid var(--bdr);white-space:nowrap;}
tbody tr{border-bottom:1px solid var(--bdr);}tbody tr:hover{background:rgba(255,106,0,.05);}
td{padding:9px 12px;vertical-align:middle;}
.acts{white-space:nowrap;display:flex;gap:5px;flex-wrap:wrap;}
.badge{display:inline-block;padding:3px 9px;border-radius:12px;font-size:11px;font-weight:600;}
.b-ok{background:rgba(34,197,94,.15);color:var(--grn);}.b-no{background:rgba(239,68,68,.1);color:var(--red);}.b-wait{background:rgba(255,106,0,.12);color:var(--or);}
.spin{display:inline-block;width:13px;height:13px;border:2px solid rgba(255,255,255,.3);border-top-color:#fff;border-radius:50%;animation:spin .6s linear infinite;}
@keyframes spin{to{transform:rotate(360deg);}}
#toast-wrap{position:fixed;bottom:22px;right:22px;display:flex;flex-direction:column;gap:8px;z-index:999;}
.toast{background:var(--card);border-left:4px solid var(--grn);padding:11px 16px;border-radius:7px;font-size:13px;box-shadow:0 4px 16px rgba(0,0,0,.4);animation:fi .2s;min-width:260px;max-width:420px;word-break:break-word;}
.toast.err{border-color:var(--red);}@keyframes fi{from{opacity:0;transform:translateY(8px)}to{opacity:1}}
#pw{display:none;min-width:220px;}.pbar{height:6px;background:var(--bdr);border-radius:3px;overflow:hidden;margin-top:6px;}
.pbar .fill{height:100%;background:var(--or);transition:width .4s;}
#modal{display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:1000;align-items:center;justify-content:center;}
#modal.show{display:flex;}
#modal-inner{background:var(--card);border-radius:12px;width:90%;max-width:900px;height:82vh;display:flex;flex-direction:column;overflow:hidden;}
#modal-hdr{padding:14px 18px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid var(--bdr);}
#modal-title{font-weight:700;color:var(--or);}#modal-close{cursor:pointer;font-size:20px;color:#888;background:none;border:none;}
#modal-frame{flex:1;border:none;}
/* Gmail modal */
#gmodal{display:none;position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:1001;align-items:center;justify-content:center;}
#gmodal.show{display:flex;}
#gmodal-inner{background:var(--card);border-radius:14px;width:440px;padding:28px;box-shadow:0 8px 40px rgba(0,0,0,.5);}
#gmodal h3{color:var(--or);margin-bottom:16px;font-size:17px;}
.field{margin-bottom:14px;}
.field label{display:block;font-size:12px;color:#9999CC;margin-bottom:5px;text-transform:uppercase;letter-spacing:.4px;}
.field input{width:100%;background:#1a1a2e;border:1px solid var(--bdr);border-radius:7px;padding:9px 12px;color:#E2E2F0;font-size:14px;outline:none;}
.field input:focus{border-color:var(--or);}
.hint{font-size:11px;color:#666;margin-top:4px;line-height:1.5;}
.hint a{color:var(--tl);}
#gmail-status{font-size:12px;padding:4px 10px;border-radius:5px;background:rgba(239,68,68,.1);color:var(--red);}
#gmail-status.ok{background:rgba(34,197,94,.12);color:var(--grn);}
</style>
</head>
<body>
<nav>
  <span class="logo">tMart</span>
  <span class="sub">Rebate Debit Note Dashboard</span>
  <span id="gmail-status">⚙ Gmail not set</span>
  <button class="tl" onclick="openGmail()">⚙ Gmail Settings</button>
</nav>
<div class="cards">
  <div class="card"><div class="num" id="c-total">—</div><div class="lbl">Total Rows</div></div>
  <div class="card"><div class="num" id="c-pdf">—</div><div class="lbl">PDFs Generated</div></div>
  <div class="card"><div class="num" id="c-email">—</div><div class="lbl">Emails Sent</div></div>
</div>
<div class="toolbar">
  <label class="btn or">📂 Upload Excel<input type="file" accept=".xlsx" onchange="uploadExcel(this)"/></label>
  <label class="btn out">📄 Upload Template<input type="file" accept=".docx" onchange="uploadTemplate(this)"/></label>
  <button class="or" onclick="generateAll()">⚡ Generate All PDFs</button>
  <button class="grn" onclick="sendAll()">📧 Send All Emails</button>
  <button class="blu" id="dl-all-btn" onclick="dlAll()">⬇ Download All (ZIP)</button>
  <div class="sep"></div>
  <div id="pw">
    <div style="font-size:12px;color:#9999CC;" id="pl">Working…</div>
    <div class="pbar"><div class="fill" id="pb" style="width:0%"></div></div>
  </div>
  <button class="out" onclick="loadRows()">↺ Refresh</button>
</div>
<div class="tbl-wrap">
<table>
  <thead>
    <tr>
      <th><input type="checkbox" id="chk-all" onchange="toggleAll(this)"></th>
      <th>#</th><th>Supplier</th><th>DN Number</th><th>Email</th>
      <th>PDF</th><th>Email Sent</th><th>Actions</th>
    </tr>
  </thead>
  <tbody id="tbody">
    <tr><td colspan="8" style="text-align:center;padding:40px;color:#666;">Upload DN_DATA.xlsx to begin</td></tr>
  </tbody>
</table>
</div>

<!-- PDF Preview Modal -->
<div id="modal"><div id="modal-inner">
  <div id="modal-hdr">
    <span id="modal-title">PDF Preview</span>
    <button id="modal-close" onclick="closeModal()">✕</button>
  </div>
  <iframe id="modal-frame" src=""></iframe>
</div></div>

<!-- Gmail Settings Modal -->
<div id="gmodal"><div id="gmodal-inner">
  <h3>⚙ Gmail Settings</h3>
  <div class="field">
    <label>Gmail Address</label>
    <input type="email" id="g-user" placeholder="yourname@gmail.com"/>
  </div>
  <div class="field">
    <label>Gmail App Password</label>
    <input type="password" id="g-pass" placeholder="xxxx xxxx xxxx xxxx"/>
    <div class="hint">
      Use an <strong>App Password</strong>, not your regular password.<br>
      Get it at: <a href="https://myaccount.google.com/apppasswords" target="_blank">myaccount.google.com/apppasswords</a><br>
      (Requires 2-Step Verification enabled on your Google account)
    </div>
  </div>
  <div style="display:flex;gap:10px;margin-top:18px;">
    <button class="tl" onclick="saveGmail()" style="flex:1;">💾 Save &amp; Connect</button>
    <button class="out" onclick="closeGmail()" style="flex:1;">Cancel</button>
  </div>
</div></div>

<div id="toast-wrap"></div>
<script>
let rows=[], SID='';

// Init
(async()=>{
  const r=await fetch('/api/session-id').then(r=>r.json());
  SID=r.sid;
  loadRows();
  checkGmail();
})();

async function loadRows(){
  const r=await fetch('/api/rows').then(r=>r.json());
  if(!r.ok){toast('Error: '+r.error,true);return;}
  rows=r.rows; render(); cards();
}
function cards(){
  document.getElementById('c-total').textContent=rows.length;
  document.getElementById('c-pdf').textContent=rows.filter(r=>r.pdf_exists).length;
  document.getElementById('c-email').textContent=rows.filter(r=>r.email_sent).length;
}
function render(){
  const b=document.getElementById('tbody');
  if(!rows.length){b.innerHTML='<tr><td colspan="8" style="text-align:center;padding:40px;color:#666;">No data — upload DN_DATA.xlsx</td></tr>';return;}
  b.innerHTML=rows.map((r,i)=>`
  <tr id="tr-${i}">
    <td><input type="checkbox" class="rc" data-idx="${i}"></td>
    <td style="color:#888">${r.si_no}</td>
    <td><strong>${r.supplier}</strong></td>
    <td style="color:#CCC;font-size:12px">${r.dn_number||''}</td>
    <td style="font-size:11px;color:#9999CC;max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${r.email}">${r.email}</td>
    <td>${r.pdf_exists?'<span class="badge b-ok">✓ Ready</span>':'<span class="badge b-no">Pending</span>'}</td>
    <td>${r.email_sent?'<span class="badge b-ok">✓ Sent</span>':'<span class="badge b-wait">Not sent</span>'}</td>
    <td class="acts">
      <button class="or sm" id="gb-${i}" onclick="genOne(${i})">⚡</button>
      ${r.pdf_exists?`<button class="pur sm" onclick="previewOne('${r.pdf_name}','${r.supplier}')">👁</button>`:''}
      ${r.pdf_exists?`<button class="blu sm" onclick="dlOne('${r.pdf_name}')">⬇</button>`:''}
      ${r.pdf_exists?`<button class="grn sm" id="mb-${i}" onclick="sendOne(${i})">📧</button>`:''}
      ${r.pdf_exists?`<button class="red sm" onclick="delOne(${i},'${r.pdf_name}')">🗑</button>`:''}
    </td>
  </tr>`).join('');
}

function sel(){
  const idx=[...document.querySelectorAll('.rc:checked')].map(c=>+c.dataset.idx);
  return idx.length?idx.map(i=>rows[i]):rows;
}
function toggleAll(cb){document.querySelectorAll('.rc').forEach(c=>c.checked=cb.checked);}

/* Single generate */
async function genOne(i){
  const btn=document.getElementById(`gb-${i}`);
  btn.innerHTML='<span class="spin"></span>';btn.disabled=true;
  try{
    const res=await post('/api/generate',{rows:[rows[i]]});
    const r0=res.results[0];
    if(r0.ok){rows[i].pdf_exists=true;toast('✓ PDF: '+r0.pdf_name);}
    else toast('Error: '+r0.error,true);
  }catch(e){toast('Error: '+e.message,true);}
  btn.innerHTML='⚡';btn.disabled=false;render();cards();
}

/* Batch generate with polling */
async function generateAll(){
  const s=sel();
  if(!s.length){toast('No rows',true);return;}
  prog(0,s.length,'Starting generation…');
  try{
    const res=await post('/api/generate-start',{rows:s});
    if(!res.ok){hideProg();toast('Error starting',true);return;}
    const total=res.total;
    const poll=setInterval(async()=>{
      const st=await fetch('/api/generate-status').then(r=>r.json());
      prog(st.done,total,`Generating… ${st.done}/${total}`);
      if(!st.running){
        clearInterval(poll);
        st.results.forEach(r=>{
          if(r.ok){const m=rows.find(x=>x.pdf_name===r.pdf_name);if(m)m.pdf_exists=true;}
        });
        const ok=st.results.filter(r=>r.ok).length;
        const err=st.results.filter(r=>!r.ok).length;
        hideProg();
        toast(err===0?`✓ ${ok} PDFs generated`:`${ok} ok, ${err} failed`,err>0);
        if(err>0) st.errors.slice(0,3).forEach(r=>toast(r.pdf_name+': '+r.error,true));
        render();cards();
      }
    },2500);
  }catch(e){hideProg();toast('Error: '+e.message,true);}
}

/* Preview */
function previewOne(pn,sup){
  document.getElementById('modal-title').textContent=sup+' — '+pn;
  document.getElementById('modal-frame').src=`/preview/${SID}/${encodeURIComponent(pn)}`;
  document.getElementById('modal').classList.add('show');
}
function closeModal(){document.getElementById('modal').classList.remove('show');document.getElementById('modal-frame').src='';}
document.getElementById('modal').addEventListener('click',function(e){if(e.target===this)closeModal();});

/* Download */
function dlOne(pn){window.open(`/download/${SID}/${encodeURIComponent(pn)}`,'_blank');}
function dlAll(){window.open(`/download-all/${SID}`,'_blank');}

/* Send email */
async function sendOne(i){
  const row=rows[i];
  if(!row.pdf_exists){toast('Generate first',true);return;}
  const btn=document.getElementById(`mb-${i}`);
  btn.innerHTML='<span class="spin"></span>';btn.disabled=true;
  const res=await post('/api/send-email',{pdf_name:row.pdf_name,email:row.email,cc:row.cc,supplier:row.supplier,dn_number:row.dn_number});
  if(res.ok){rows[i].email_sent=true;toast('✓ Email sent to '+row.email);}
  else toast('Email error: '+res.error,true);
  btn.innerHTML='📧';btn.disabled=false;render();cards();
}
async function sendAll(){
  const s=sel().filter(r=>r.pdf_exists);
  if(!s.length){toast('No PDFs ready',true);return;}
  prog(0,s.length,'Sending emails…');let done=0,err=0;
  for(const row of s){
    const res=await post('/api/send-email',{pdf_name:row.pdf_name,email:row.email,cc:row.cc,supplier:row.supplier,dn_number:row.dn_number});
    done++;if(res.ok){const m=rows.find(r=>r.pdf_name===row.pdf_name);if(m)m.email_sent=true;}else err++;
    prog(done,s.length,`Sending… ${done}/${s.length}`);
  }
  hideProg();toast(err===0?`✓ ${done} emails sent`:`${done-err} sent, ${err} failed`,err>0);render();cards();
}

/* Delete */
async function delOne(i,pn){
  if(!confirm(`Delete:\n${pn}?`))return;
  const res=await fetch('/api/delete/'+encodeURIComponent(pn),{method:'DELETE'}).then(r=>r.json());
  if(res.ok){rows[i].pdf_exists=false;rows[i].email_sent=false;toast('Deleted');}
  else toast('Delete failed: '+res.error,true);
  render();cards();
}

/* Upload Excel */
async function uploadExcel(inp){
  if(!inp.files.length)return;
  const fd=new FormData();fd.append('file',inp.files[0]);
  const res=await fetch('/api/upload-excel',{method:'POST',body:fd}).then(r=>r.json());
  if(res.ok){toast('✓ Excel uploaded');setTimeout(loadRows,400);}
  else toast('Upload failed: '+res.error,true);
  inp.value='';
}

/* Upload template */
async function uploadTemplate(inp){
  if(!inp.files.length)return;
  const fd=new FormData();fd.append('file',inp.files[0]);
  const res=await fetch('/api/upload-template',{method:'POST',body:fd}).then(r=>r.json());
  toast(res.ok?'✓ Template uploaded':'Template upload failed: '+res.error,!res.ok);
  inp.value='';
}

/* Gmail settings */
async function checkGmail(){
  const r=await fetch('/api/gmail-status').then(r=>r.json());
  const el=document.getElementById('gmail-status');
  if(r.configured){el.textContent='✓ Gmail: '+r.user;el.className='ok';}
  else{el.textContent='⚠ Gmail not set';el.className='';}
}
function openGmail(){
  fetch('/api/gmail-status').then(r=>r.json()).then(r=>{
    if(r.configured) document.getElementById('g-user').value=r.user;
  });
  document.getElementById('gmodal').classList.add('show');
}
function closeGmail(){document.getElementById('gmodal').classList.remove('show');}
async function saveGmail(){
  const user=document.getElementById('g-user').value.trim();
  const pass=document.getElementById('g-pass').value.trim();
  if(!user||!pass){toast('Enter Gmail and App Password',true);return;}
  const res=await post('/api/gmail-settings',{user,password:pass});
  if(res.ok){toast('✓ Gmail saved: '+res.user);checkGmail();closeGmail();}
  else toast('Error: '+res.error,true);
}

/* Helpers */
async function post(url,data){return fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)}).then(r=>r.json());}
function prog(done,total,lbl){
  document.getElementById('pw').style.display='block';
  document.getElementById('pl').textContent=lbl;
  document.getElementById('pb').style.width=(total>0?Math.round(done/total*100):0)+'%';
}
function hideProg(){document.getElementById('pw').style.display='none';}
function toast(msg,err=false,dur=5000){
  const w=document.getElementById('toast-wrap');
  const d=document.createElement('div');d.className='toast'+(err?' err':'');d.textContent=msg;
  w.appendChild(d);setTimeout(()=>d.remove(),dur);
}
</script>
</body>
</html>"""

# ── Startup ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import threading, webbrowser
    threading.Timer(1.2, lambda: webbrowser.open("http://localhost:5000")).start()
    print("=" * 55)
    print("  tMart DN Cloud App  —  http://localhost:5000")
    print("=" * 55)
    app.run(host="0.0.0.0", port=5000, debug=False)
