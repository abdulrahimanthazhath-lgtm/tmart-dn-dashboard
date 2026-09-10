"""
tMart Rebate DN Dashboard — Cloud Edition
==========================================
• Runs on Render.com (free tier) with LibreOffice for PDF conversion
• Each colleague gets their own private session
• Email auto-sent via Gmail API (OAuth) or Gmail SMTP (App Password)
"""

import os, re, io, json, uuid, shutil, tempfile, zipfile, threading, datetime
import smtplib, urllib.request as urlreq, base64 as b64lib
from email.mime.multipart import MIMEMultipart
from email.mime.text       import MIMEText
from email.mime.application import MIMEApplication
from flask import Flask, jsonify, send_file, request, session

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "tmart-dn-cloud-change-this-secret")

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
SESSIONS_DIR = os.path.join(tempfile.gettempdir(), "tmart_sessions")
os.makedirs(SESSIONS_DIR, exist_ok=True)

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
    import subprocess
    lo = shutil.which("libreoffice") or shutil.which("soffice")
    if lo:
        tmp = tempfile.mkdtemp()
        try:
            r = subprocess.run([lo, "--headless", "--convert-to", "pdf",
                                "--outdir", tmp, os.path.abspath(docx_path)],
                               capture_output=True, text=True, timeout=120)
            base = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
            src  = os.path.join(tmp, base)
            if os.path.exists(src):
                shutil.move(src, pdf_path)
                return True, None
            return False, r.stderr or "LibreOffice: PDF not created"
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    td = docx_path.replace('"', '`"')
    op = pdf_path.replace('"',  '`"')
    ps = (f'$w=New-Object -ComObject Word.Application;$w.Visible=$false;$w.DisplayAlerts=0;'
          f'$d=$w.Documents.Open("{td}");$d.ExportAsFixedFormat("{op}",17);'
          f'$d.Close($false);$w.Quit()')
    r = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps],
                       capture_output=True, text=True, timeout=120)
    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
        return True, None
    return False, r.stderr or "Word COM: PDF not created"

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

_jobs = {}

def _run_batch(sid, rows_list):
    status = _jobs[sid]
    for row in rows_list:
        if status.get("cancelled"): break
        r = generate_one(row, sid=sid)
        status["done"]   += 1
        status["results"].append(r)
        if not r["ok"]: status["errors"].append(r)
        _write_job(sid, status)
    status["running"] = False
    _write_job(sid, status)

def _write_job(sid, status):
    path = os.path.join(SESSIONS_DIR, sid, "job.json")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f: json.dump(status, f)

def _read_job(sid):
    path = os.path.join(SESSIONS_DIR, sid, "job.json")
    try:
        with open(path) as f: return json.load(f)
    except:
        return {"total":0,"done":0,"running":False,"results":[],"errors":[]}

# ── Gmail SMTP ─────────────────────────────────────────────────────────────────

def send_gmail_smtp(to, cc, subject, body_text, attachment_path, gmail_user, gmail_pass):
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

# ── Gmail API (OAuth) ──────────────────────────────────────────────────────────

def send_gmail_oauth(access_token, to, cc, subject, body_text, pdf_path):
    """Send email via Gmail API using an OAuth2 access token."""
    msg = MIMEMultipart()
    msg["To"]      = to
    if cc: msg["CC"] = cc
    msg["Subject"] = subject
    msg.attach(MIMEText(body_text, "plain"))
    if pdf_path and os.path.exists(pdf_path):
        with open(pdf_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(pdf_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(pdf_path)}"'
        msg.attach(part)
    raw = b64lib.urlsafe_b64encode(msg.as_bytes()).decode().rstrip("=")
    payload = json.dumps({"raw": raw}).encode()
    req = urlreq.Request(
        "https://gmail.googleapis.com/gmail/v1/users/me/messages/send",
        data=payload,
        headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    )
    try:
        with urlreq.urlopen(req, timeout=30) as resp:
            return True, json.loads(resp.read()).get("id", "sent")
    except Exception as e:
        err_body = ""
        if hasattr(e, 'read'):
            try: err_body = json.loads(e.read()).get("error", {}).get("message", "")
            except: pass
        return False, err_body or str(e)

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
    od   = outdir()
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

@app.route("/api/generate", methods=["POST"])
def api_generate():
    body      = request.json or {}
    rows_list = [item["row"] for item in body.get("rows", [])]
    sid = get_sid()
    results   = [generate_one(r, sid=sid) for r in rows_list]
    return jsonify({"results": results})

@app.route("/api/send-email", methods=["POST"])
def api_send_email():
    """SMTP App Password mode."""
    body = request.json or {}
    pn   = body.get("pdf_name", "")
    pdf_path = os.path.join(outdir(), pn)
    if not os.path.exists(pdf_path):
        return jsonify({"ok": False, "error": "PDF not found — generate it first"})
    gu = session.get("gmail_user", "")
    gp = session.get("gmail_pass", "")
    if not gu or not gp:
        return jsonify({"ok": False, "error": "Gmail not configured — click ⚙ Gmail Settings"})
    try:
        to       = body.get("email", "")
        cc       = body.get("cc", "")
        supplier = body.get("supplier", "Supplier")
        dn_num   = body.get("dn_number", "")
        subject  = f"Rebate Debit Note {dn_num} - DH Store Bahrain (tMart)".strip()
        body_t   = (f"Dear {supplier},\n\n"
                    "Please find attached your Rebate Debit Note.\n\n"
                    "Regards,\ntMart Finance Team")
        send_gmail_smtp(to, cc, subject, body_t, pdf_path, gu, gp)
        sent = session.get("sent", {}); sent[pn] = True; session["sent"] = sent
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

@app.route("/api/send-oauth-email", methods=["POST"])
def api_send_oauth_email():
    """Gmail API OAuth mode — access_token comes from the browser (Google Identity Services)."""
    body         = request.json or {}
    access_token = body.get("access_token", "")
    pn           = body.get("pdf_name", "")
    if not access_token:
        return jsonify({"ok": False, "error": "No OAuth token — click Sign in with Google"})
    pdf_path = os.path.join(outdir(), pn)
    if not os.path.exists(pdf_path):
        return jsonify({"ok": False, "error": "PDF not found — generate it first"})
    to       = body.get("email", "")
    cc       = body.get("cc", "")
    supplier = body.get("supplier", "Supplier")
    dn_num   = body.get("dn_number", "")
    subject  = f"Rebate Debit Note {dn_num} - DH Store Bahrain (tMart)".strip()
    body_t   = (f"Dear {supplier},\n\n"
                "Please find attached your Rebate Debit Note.\n\n"
                "Regards,\ntMart Finance Team")
    ok, result = send_gmail_oauth(access_token, to, cc, subject, body_t, pdf_path)
    if ok:
        sent = session.get("sent", {}); sent[pn] = True; session["sent"] = sent
        return jsonify({"ok": True, "message_id": result})
    return jsonify({"ok": False, "error": result})

@app.route("/api/gmail-settings", methods=["POST"])
def api_gmail_settings():
    body = request.json or {}
    mode = body.get("mode", "smtp")
    session["gmail_mode"] = mode
    session["gmail_user"] = body.get("user", "").strip()
    session["gmail_pass"] = body.get("password", "").strip() if mode == "smtp" else ""
    return jsonify({"ok": True, "user": session["gmail_user"], "mode": mode})

@app.route("/api/gmail-status")
def api_gmail_status():
    mode = session.get("gmail_mode", "")
    user = session.get("gmail_user", "")
    configured = bool(user) if mode in ("browser", "oauth") else bool(user and session.get("gmail_pass"))
    return jsonify({"configured": configured, "user": user, "mode": mode})

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
            return jsonify({"ok": False, "error": f"Sheet1 not found. Available: {sheets}"})
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

DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>tMart – Rebate DN Dashboard</title>
<script src="https://accounts.google.com/gsi/client" async defer></script>
<style>
:root{--or:#E05500;--dark:#F0F2F5;--card:#FFFFFF;--bdr:#D8DBE8;--grn:#16A34A;--red:#DC2626;--blu:#2563EB;--pur:#7C3AED;--tl:#0D9488;}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:"Segoe UI",sans-serif;background:var(--dark);color:#1A1A2E;min-height:100vh;}
nav{background:var(--card);border-bottom:3px solid var(--or);padding:14px 28px;display:flex;align-items:center;gap:14px;flex-wrap:wrap;}
.sub{font-size:13px;color:#666;flex:1;}
.cards{display:flex;gap:16px;padding:22px 28px 0;flex-wrap:wrap;}
.card{flex:1;min-width:140px;background:var(--card);border:1px solid var(--bdr);border-radius:10px;padding:16px 20px;}
.card .num{font-size:28px;font-weight:700;color:var(--or);}.card .lbl{font-size:11px;color:#666;margin-top:2px;text-transform:uppercase;letter-spacing:.5px;}
.toolbar{display:flex;gap:8px;padding:16px 28px;flex-wrap:wrap;align-items:center;}
input[type=file]{display:none;}
button,label.btn{cursor:pointer;border:none;border-radius:7px;padding:9px 16px;font-size:13px;font-weight:600;transition:opacity .15s;display:inline-flex;align-items:center;gap:5px;}
button:hover,label.btn:hover{opacity:.85;}
.or{background:var(--or);color:#fff;}.grn{background:var(--grn);color:#fff;}.blu{background:var(--blu);color:#fff;}
.pur{background:var(--pur);color:#fff;}.red{background:var(--red);color:#fff;}.tl{background:var(--tl);color:#fff;}
.out{background:transparent;border:1px solid var(--bdr);color:#555;}
.sm{padding:5px 10px;font-size:12px;border-radius:5px;white-space:nowrap;}.sep{flex:1;}
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
.toast{background:var(--card);border-left:4px solid var(--grn);padding:11px 16px;border-radius:7px;font-size:13px;box-shadow:0 4px 16px rgba(0,0,0,.15);animation:fi .2s;min-width:260px;max-width:420px;word-break:break-word;color:#1A1A2E;}
.toast.err{border-color:var(--red);}@keyframes fi{from{opacity:0;transform:translateY(8px)}to{opacity:1}}
#pw{display:none;min-width:220px;}.pbar{height:6px;background:var(--bdr);border-radius:3px;overflow:hidden;margin-top:6px;}
.pbar .fill{height:100%;background:var(--or);transition:width .4s;}
/* Modals */
#modal{display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:1000;align-items:center;justify-content:center;}
#modal.show{display:flex;}
#modal-inner{background:var(--card);border-radius:12px;width:90%;max-width:900px;height:82vh;display:flex;flex-direction:column;overflow:hidden;}
#modal-hdr{padding:14px 18px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid var(--bdr);}
#modal-title{font-weight:700;color:var(--or);}#modal-close{cursor:pointer;font-size:20px;color:#888;background:none;border:none;}
#modal-frame{flex:1;border:none;}
#gmodal{display:none;position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:1001;align-items:center;justify-content:center;}
#gmodal.show{display:flex;}
#gmodal-inner{background:var(--card);border-radius:14px;width:480px;max-height:90vh;overflow-y:auto;padding:28px;box-shadow:0 8px 40px rgba(0,0,0,.4);}
#gmodal h3{color:var(--or);margin-bottom:16px;font-size:17px;}
.tabs{display:flex;gap:6px;margin-bottom:18px;}
.tab-btn{flex:1;padding:10px 6px;border-radius:8px;border:2px solid var(--bdr);background:transparent;color:#888;cursor:pointer;font-size:12px;font-weight:600;text-align:center;line-height:1.4;transition:all .15s;}
.tab-btn.active{border-color:var(--tl);background:var(--tl);color:#fff;}
.panel{display:none;}.panel.show{display:block;}
.field{margin-bottom:14px;}
.field label{display:block;font-size:12px;color:#666;margin-bottom:5px;text-transform:uppercase;letter-spacing:.4px;}
.field input{width:100%;background:#F7F8FA;border:1px solid var(--bdr);border-radius:7px;padding:9px 12px;color:#1A1A2E;font-size:14px;outline:none;}
.field input:focus{border-color:var(--or);}
.hint{font-size:11px;color:#666;margin-top:5px;line-height:1.6;}
.hint a{color:var(--tl);text-decoration:none;}
.hint a:hover{text-decoration:underline;}
#gstatus-bar{font-size:12px;padding:6px 12px;border-radius:6px;background:rgba(239,68,68,.1);color:var(--red);margin-bottom:14px;display:none;}
#gstatus-bar.show{display:block;}
#gstatus-bar.ok{background:rgba(34,197,94,.12);color:var(--grn);}
#oauth-btn-wrap{margin:12px 0;}
#oauth-connected{background:rgba(34,197,94,.1);border:1px solid var(--grn);border-radius:8px;padding:10px 14px;font-size:13px;color:var(--grn);display:none;margin:12px 0;}
</style>
</head>
<body>
<div style="background:#fff;border-bottom:3px solid #FF6200;padding:8px 0;text-align:center;">
  <img src="data:image/png;base64,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Ba6IiCUKXBERSxS4IiKWKHBFRCxR4IqIWKLAFRGxRIErImKJAldAxBIFroiIJQpcERFLFLgiIpYocEVELFHgiohYosAVEbFEgSsiYokCV0TEEgWuiIglClwREUsUuCIilihwRUQsUeCKiFiiwBURsUSBKyJiiQJXRMQSBa6IiCUKXBERSxS4IiKWKHBFRCxR4IqIWKLAFRGxRIErImKJAldAxBIFroiIJQpcERFLFLgiIpYocEVELFHgiohYosAVEbFEgSsiYokCV0TEEgWuiIglClwREUsUuCIilihwRUQsUeCKiFiiwBURsUSBKyJiiQJXRMQSBa6IiCUKXBERSxS4IiKWKHBFRCxR4IqIWKLAFRGxRIErImKJAldAxBIFroiIJQpcERFLFLgiIpYocEVELFHgiohYosAVEbFEgSsiYokCV0TEEgWuiIglClwREUsUuCIilihwRUQsUeCKiFiiwBURsUSBKyJiiQJXRMQSBa6IiCUKXBERSxS4IiKWKHBFRCxR4IqIWKLAFRGxRIErImKJAldAxBIFroiIJQpcERFLFLgiIpYocEVELFHgiohYosAVEbFEgSsiYokCV0TEEgWuiIglClwREUsUuCIilihwRUQsUeCKiFiiwBURsUSBKyJiiQJXRMQSBa6IiCUKXBERSxS4IiKWKHBFRCxR4IqIWKLAFRGxRIErImKJAldAxBIFroiIJQpcERFLFLgiIpYocEVELFHgiohYosAVEbFEgSsiYokCV0TEEgWuiIglClwREUsUuCIilihwRUQsUeCKiFiiwBURsUSBKyJiiQJXRMQSBa6IiCUKXBERSxS4IiKWKHBFRCxR4IqIWKLAFRGxRIErImKJAldAxBIFroiIJQpcERFLFLgiIpYocEVELFHgiohYosAVEbFEgSsiYokCV0TEEgWuiIglClwREUsUuCIilihwRUQsUeCKiFiiwBURsUSBKyJiiQJXRMQSBa6IiCUKXBERSxS4IiKWKHBFRCxR4IqIWKLAFRGxRIErImKJAldAxBIFroiIJQpcERFLFLgiIpYocEVELFHgiohYosAVEbFEgSsiYokCV0TEEgWuiIglClwREUsUuCIilihwRUQsUeCKiFiiwBURsUSBKyJiiQJXRMQSBa6IiCUKXBERSxS4IiKWKHBFRCxR4IqIWKLAFRGxRIErImKJAldAxBIFroiIJQpcERFLFLgiIpYocEVELFHgiohYosAVEbFEgSsiYokCV0TEEgWuiIglClwREUsUuCIilihwRUQsUeCKiFiiwBURsUSBKyJiiQJXRMQSBa6IiCUKXBERSxS4IiKWKHBFRCxR4IqIWKLAFRGxRIErImKJAldAxBIFroiIJQpcERFLFLgiIpYocEVELFHgiohYosAVEbFEgSsiYokCV0TEEgWuiIglClwREUsUuCIilihwRUQsUeCKiFiiwBURsUSBKyJiiQJXRMQSBa6IiCUKXBERSxS4IiKWKHBFRCxR4IqIWKLAFRGxRIErImKJAldAxBIFroiIJQpcERFLFLgiIpYocEVELFHgiohYosAVEbFEgSsiYokCV0TEEgWuiIglClwREUsUuCIilihwRUQsUeCKiFiiwBURsUSBKyJiiQJXRMQSBa6IiCUKXBERSxS4IiKWKHBFRCxR4IqIWKLAFRGxRIErImKJAldA" alt="talabat mart" style="height:52px;object-fit:contain;"/>
</div>
<nav>
  <span class="sub" style="font-weight:600;">Rebate Debit Note Dashboard</span>
  <span id="gstatus-nav" style="font-size:12px;color:#888;">⚙ Gmail not set</span>
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
  <button class="blu" onclick="dlAll()">⬇ Download All (ZIP)</button>
  <div class="sep"></div>
  <div id="pw">
    <div style="font-size:12px;color:#666;" id="pl">Working…</div>
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
  <div class="tabs">
    <button class="tab-btn active" id="tab-oauth" onclick="setTab('oauth')">🔗 Auto Send<br><span style="font-weight:400;font-size:10px;">Sign in with Google</span></button>
    <button class="tab-btn" id="tab-smtp" onclick="setTab('smtp')">🔑 Auto Send<br><span style="font-weight:400;font-size:10px;">App Password</span></button>
    <button class="tab-btn" id="tab-browser" onclick="setTab('browser')">🌐 Manual<br><span style="font-weight:400;font-size:10px;">Chrome Gmail</span></button>
  </div>

  <!-- OAuth panel -->
  <div id="panel-oauth" class="panel show">
    <div class="field">
      <label>Google OAuth Client ID</label>
      <input type="text" id="g-client-id" placeholder="xxxxxxxx.apps.googleusercontent.com"/>
      <div class="hint">
        Get your Client ID from <a href="https://console.cloud.google.com" target="_blank">Google Cloud Console</a>:<br>
        1. Create project → APIs &amp; Services → Enable <strong>Gmail API</strong><br>
        2. Credentials → Create → OAuth Client ID → Web Application<br>
        3. Add your Render URL under <em>Authorized JavaScript origins</em><br>
        4. Copy the Client ID and paste it above.<br><br>
        ✅ <strong>One-time setup.</strong> After that, all colleagues just click Sign In — no App Password ever.
      </div>
    </div>
    <div id="oauth-connected" id="oauth-connected-box">
      ✅ <strong id="oauth-email-disp"></strong> connected — emails send automatically with PDF attached!
    </div>
    <div id="oauth-btn-wrap">
      <button class="tl" onclick="connectGoogle()" style="width:100%;padding:11px;">🔗 Sign in with Google</button>
    </div>
    <div style="display:flex;gap:10px;margin-top:14px;">
      <button class="tl" onclick="saveGmail()" style="flex:1;">💾 Save Settings</button>
      <button class="out" onclick="closeGmail()" style="flex:1;">Cancel</button>
    </div>
  </div>

  <!-- SMTP panel -->
  <div id="panel-smtp" class="panel">
    <div class="field">
      <label>Gmail Address</label>
      <input type="email" id="g-user-smtp" placeholder="yourname@gmail.com"/>
    </div>
    <div class="field">
      <label>Gmail App Password</label>
      <input type="password" id="g-pass" placeholder="xxxx xxxx xxxx xxxx"/>
      <div class="hint">
        Use an <strong>App Password</strong>, not your regular password.<br>
        Get it at: <a href="https://myaccount.google.com/apppasswords" target="_blank">myaccount.google.com/apppasswords</a><br>
        (Requires 2-Step Verification on your Google account)
      </div>
    </div>
    <div style="display:flex;gap:10px;margin-top:14px;">
      <button class="tl" onclick="saveGmail()" style="flex:1;">💾 Save &amp; Connect</button>
      <button class="out" onclick="closeGmail()" style="flex:1;">Cancel</button>
    </div>
  </div>

  <!-- Browser panel -->
  <div id="panel-browser" class="panel">
    <div class="field">
      <label>Your Gmail Address</label>
      <input type="email" id="g-user-browser" placeholder="yourname@gmail.com"/>
    </div>
    <div class="hint" style="margin-top:4px;">
      ℹ️ Manual mode — clicking Send opens Gmail in a new tab and downloads the PDF.<br>
      You attach the PDF yourself and click Send in Gmail.<br><br>
      For fully automatic sending, use <strong>Sign in with Google</strong> or <strong>App Password</strong> instead.
    </div>
    <div style="display:flex;gap:10px;margin-top:18px;">
      <button class="tl" onclick="saveGmail()" style="flex:1;">💾 Save</button>
      <button class="out" onclick="closeGmail()" style="flex:1;">Cancel</button>
    </div>
  </div>
</div></div>

<div id="toast-wrap"></div>
<script>
let rows=[], SID='';
let _gmailMode='oauth';
let _oauthToken=null, _tokenExpiry=0, _oauthEmail='';
let _clientId='';

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
    <td style="color:#666">${r.si_no}</td>
    <td><strong>${r.supplier}</strong></td>
    <td style="color:#888;font-size:12px">${r.dn_number||''}</td>
    <td style="font-size:11px;color:#555;max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${r.email}">${r.email}</td>
    <td>${r.pdf_exists?'<span class="badge b-ok">✓ Ready</span>':'<span class="badge b-no">Pending</span>'}</td>
    <td>${r.email_sent?'<span class="badge b-ok">✓ Sent</span>':'<span class="badge b-wait">Not sent</span>'}</td>
    <td class="acts">
      <button class="or sm" id="gb-${i}" onclick="genOne(${i})">⚡ Generate</button>
      ${r.pdf_exists?`<button class="pur sm" onclick="previewOne('${r.pdf_name}','${r.supplier}')">👁 Preview</button>`:''}
      ${r.pdf_exists?`<button class="blu sm" onclick="dlOne('${r.pdf_name}')">⬇ Download</button>`:''}
      ${r.pdf_exists?`<button class="grn sm" id="mb-${i}" onclick="sendOne(${i})">📧 Send</button>`:''}
      ${r.pdf_exists?`<button class="red sm" onclick="delOne(${i},'${r.pdf_name}')">🗑 Delete</button>`:''}
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
    if(r0.ok){rows[i].pdf_exists=true;toast('✓ PDF ready: '+r0.pdf_name);}
    else toast('Error: '+r0.error,true);
  }catch(e){toast('Error: '+e.message,true);}
  btn.innerHTML='⚡ Generate';btn.disabled=false;render();cards();
}

/* Batch generate */
async function generateAll(){
  const s=sel();
  if(!s.length){toast('No rows',true);return;}
  prog(0,s.length,'Starting…');
  try{
    const res=await post('/api/generate-start',{rows:s});
    if(!res.ok){hideProg();toast('Error starting',true);return;}
    const total=res.total;
    const poll=setInterval(async()=>{
      const st=await fetch('/api/generate-status').then(r=>r.json());
      prog(st.done,total,`Generating… ${st.done}/${total}`);
      if(!st.running){
        clearInterval(poll);
        st.results.forEach(r=>{if(r.ok){const m=rows.find(x=>x.pdf_name===r.pdf_name);if(m)m.pdf_exists=true;}});
        const ok=st.results.filter(r=>r.ok).length, err=st.results.filter(r=>!r.ok).length;
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

/* Delete */
async function delOne(i,pn){
  if(!confirm('Delete '+pn+'?'))return;
  await fetch(`/api/delete/${encodeURIComponent(pn)}`,{method:'DELETE'});
  rows[i].pdf_exists=false;rows[i].email_sent=false;
  render();cards();toast('🗑 Deleted');
}

/* ── Gmail OAuth ── */
function connectGoogle(){
  if(!_clientId){toast('Enter your Google Client ID first',true);return;}
  if(typeof google==='undefined'||!google.accounts){toast('Google library still loading — try again in a moment',true);return;}
  const client=google.accounts.oauth2.initTokenClient({
    client_id:_clientId,
    scope:'https://www.googleapis.com/auth/gmail.send https://www.googleapis.com/auth/userinfo.email',
    callback:(resp)=>{
      if(resp.error){toast('Google sign-in failed: '+resp.error,true);return;}
      _oauthToken=resp.access_token;
      _tokenExpiry=Date.now()+(resp.expires_in||3599)*1000;
      // Fetch user email
      fetch('https://www.googleapis.com/oauth2/v1/userinfo?alt=json',{headers:{Authorization:'Bearer '+_oauthToken}})
        .then(r=>r.json()).then(info=>{
          _oauthEmail=info.email||'Google user';
          showOAuthConnected();
          toast('✅ Connected as '+_oauthEmail+' — ready to auto-send!');
        }).catch(()=>{_oauthEmail='Google user';showOAuthConnected();});
    }
  });
  client.requestAccessToken();
}

function showOAuthConnected(){
  document.getElementById('oauth-email-disp').textContent=_oauthEmail;
  document.getElementById('oauth-connected').style.display='block';
  document.getElementById('oauth-btn-wrap').style.display='none';
}

async function ensureToken(){
  if(_oauthToken && Date.now()<_tokenExpiry-30000) return _oauthToken;
  if(!_clientId){return null;}
  return new Promise((resolve,reject)=>{
    if(typeof google==='undefined'){reject('Google library not loaded');return;}
    const client=google.accounts.oauth2.initTokenClient({
      client_id:_clientId,
      scope:'https://www.googleapis.com/auth/gmail.send',
      prompt:'',
      callback:(resp)=>{
        if(resp.error){reject(resp.error);return;}
        _oauthToken=resp.access_token;
        _tokenExpiry=Date.now()+(resp.expires_in||3599)*1000;
        resolve(_oauthToken);
      }
    });
    client.requestAccessToken();
  });
}

/* ── Send email ── */
function gmailComposeUrl(to,cc,subject,body){
  const p=new URLSearchParams({view:'cm',to,cc,su:subject,body});
  return 'https://mail.google.com/mail/?'+p.toString();
}

async function sendOne(i){
  const row=rows[i];
  if(!row.pdf_exists){toast('Generate PDF first',true);return;}
  const btn=document.getElementById(`mb-${i}`);
  btn.innerHTML='<span class="spin"></span>';btn.disabled=true;
  const subject=`Rebate Debit Note ${row.dn_number||''} - DH Store Bahrain (tMart)`.trim();
  const body=`Dear ${row.supplier},\\n\\nPlease find attached your Rebate Debit Note.\\n\\nRegards,\\ntMart Finance Team`;
  try{
    if(_gmailMode==='oauth'){
      const token=await ensureToken();
      if(!token){toast('Sign in with Google first — open Gmail Settings',true);btn.innerHTML='📧 Send';btn.disabled=false;return;}
      const res=await post('/api/send-oauth-email',{access_token:token,pdf_name:row.pdf_name,email:row.email,cc:row.cc,supplier:row.supplier,dn_number:row.dn_number});
      if(res.ok){rows[i].email_sent=true;toast('✅ Email auto-sent to '+row.email);}
      else toast('Email error: '+res.error,true);
    } else if(_gmailMode==='smtp'){
      const res=await post('/api/send-email',{pdf_name:row.pdf_name,email:row.email,cc:row.cc,supplier:row.supplier,dn_number:row.dn_number});
      if(res.ok){rows[i].email_sent=true;toast('✅ Email sent to '+row.email);}
      else toast('Email error: '+res.error,true);
    } else {
      window.open(`/download/${SID}/${encodeURIComponent(row.pdf_name)}`,'_blank');
      setTimeout(()=>window.open(gmailComposeUrl(row.email,row.cc,subject,body),'_blank'),800);
      rows[i].email_sent=true;
      toast('📥 PDF downloaded — attach it in the Gmail tab that opened');
    }
  }catch(e){toast('Error: '+e.message,true);}
  btn.innerHTML='📧 Send';btn.disabled=false;render();cards();
}

async function sendAll(){
  const s=sel().filter(r=>r.pdf_exists);
  if(!s.length){toast('No PDFs ready — generate first',true);return;}
  let token=null;
  if(_gmailMode==='oauth'){
    try{token=await ensureToken();}catch(e){}
    if(!token){toast('Sign in with Google first — open Gmail Settings',true);return;}
  }
  prog(0,s.length,'Sending emails…');
  let sent=0,failed=0;
  for(const row of s){
    try{
      if(_gmailMode==='oauth'){
        const res=await post('/api/send-oauth-email',{access_token:token,pdf_name:row.pdf_name,email:row.email,cc:row.cc,supplier:row.supplier,dn_number:row.dn_number});
        if(res.ok){row.email_sent=true;sent++;}
        else{failed++;toast('Failed '+row.supplier+': '+res.error,true);}
      } else if(_gmailMode==='smtp'){
        const res=await post('/api/send-email',{pdf_name:row.pdf_name,email:row.email,cc:row.cc,supplier:row.supplier,dn_number:row.dn_number});
        if(res.ok){row.email_sent=true;sent++;}
        else{failed++;toast('Failed '+row.supplier+': '+res.error,true);}
      } else {
        const subject=`Rebate Debit Note ${row.dn_number||''} - DH Store Bahrain (tMart)`.trim();
        const body=`Dear ${row.supplier},\\n\\nPlease find attached your Rebate Debit Note.\\n\\nRegards,\\ntMart Finance Team`;
        window.open(gmailComposeUrl(row.email,row.cc,subject,body),'_blank');
        row.email_sent=true;sent++;
        await new Promise(r=>setTimeout(r,600));
      }
    }catch(e){failed++;}
    prog(sent+failed,s.length,`Sending… ${sent+failed}/${s.length}`);
    await new Promise(r=>setTimeout(r,200));
  }
  hideProg();
  if(_gmailMode==='browser') toast(`📥 Download ZIP then attach PDFs to the ${sent} Gmail tabs`);
  else toast(failed===0?`✅ ${sent} emails sent!`:`${sent} sent, ${failed} failed`,failed>0);
  render();cards();
}

/* ── Gmail settings modal ── */
function openGmail(){document.getElementById('gmodal').classList.add('show');}
function closeGmail(){document.getElementById('gmodal').classList.remove('show');}
document.getElementById('gmodal').addEventListener('click',function(e){if(e.target===this)closeGmail();});

function setTab(t){
  ['oauth','smtp','browser'].forEach(x=>{
    document.getElementById('tab-'+x).classList.toggle('active',x===t);
    document.getElementById('panel-'+x).classList.toggle('show',x===t);
  });
  _gmailMode=t;
}

async function saveGmail(){
  const mode=_gmailMode;
  let user='',pass='';
  if(mode==='smtp'){
    user=document.getElementById('g-user-smtp').value.trim();
    pass=document.getElementById('g-pass').value.trim();
    if(!user||!pass){toast('Enter Gmail address and App Password',true);return;}
  } else if(mode==='browser'){
    user=document.getElementById('g-user-browser').value.trim();
    if(!user){toast('Enter your Gmail address',true);return;}
  } else {
    // oauth: save client_id
    _clientId=document.getElementById('g-client-id').value.trim();
    if(!_clientId){toast('Enter your Google Client ID',true);return;}
    user=_oauthEmail||'oauth';
  }
  await post('/api/gmail-settings',{mode,user,password:pass});
  closeGmail();
  checkGmail();
  toast('✅ Gmail settings saved');
}

async function checkGmail(){
  const r=await fetch('/api/gmail-status').then(r=>r.json());
  const nav=document.getElementById('gstatus-nav');
  if(r.configured){
    const modeLabel=r.mode==='oauth'?'OAuth':r.mode==='smtp'?'SMTP':'Browser';
    nav.textContent=`✅ ${r.user} (${modeLabel})`;
    nav.style.color='#16A34A';
    _gmailMode=r.mode||'oauth';
    setTab(_gmailMode);
  } else {
    nav.textContent='⚙ Gmail not set';
    nav.style.color='#888';
  }
}

/* ── Upload ── */
async function uploadExcel(el){
  const f=el.files[0];if(!f)return;
  const fd=new FormData();fd.append('file',f);
  const r=await fetch('/api/upload-excel',{method:'POST',body:fd}).then(r=>r.json());
  if(r.ok){toast('✅ Excel uploaded');loadRows();}
  else toast('Error: '+r.error,true);
  el.value='';
}
async function uploadTemplate(el){
  const f=el.files[0];if(!f)return;
  const fd=new FormData();fd.append('file',f);
  const r=await fetch('/api/upload-template',{method:'POST',body:fd}).then(r=>r.json());
  toast(r.ok?'✅ Template uploaded':'Error: '+r.error,!r.ok);
  el.value='';
}

/* ── Utilities ── */
async function post(url,data){
  const r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
  return r.json();
}
function toast(msg,err=false){
  const w=document.getElementById('toast-wrap');
  const d=document.createElement('div');
  d.className='toast'+(err?' err':'');d.textContent=msg;
  w.appendChild(d);setTimeout(()=>d.remove(),5000);
}
function prog(done,total,label){
  document.getElementById('pw').style.display='block';
  document.getElementById('pl').textContent=label;
  document.getElementById('pb').style.width=(total?Math.round(done/total*100):0)+'%';
}
function hideProg(){document.getElementById('pw').style.display='none';}
</script>
</body>
</html>"""

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
