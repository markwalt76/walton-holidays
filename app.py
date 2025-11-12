import os
import json
import smtplib
import ssl
from email.message import EmailMessage
from datetime import datetime, date, timedelta

from flask import Flask, render_template, request, jsonify, Response

# Google Sheets
import gspread
from google.oauth2.service_account import Credentials

# -------------------------------
# Flask & templates
# -------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
app = Flask(__name__, template_folder=TEMPLATE_DIR)

# -------------------------------
# Approvers (override via ENV)
# -------------------------------
APPROVERS = {
    "mark": os.environ.get("EMAIL_MARK", "mw@walton.fr"),
    "nhan": os.environ.get("EMAIL_NHAN", "nhan@walton.fr"),
    "anh":  os.environ.get("EMAIL_ANH",  "anh@walton.fr"),
}

BRAND = "Walton Time Off"

# -------------------------------
# Helpers
# -------------------------------
def _env(key, default=None):
    return os.environ.get(key, default)

def business_days(d1: date, d2: date) -> int:
    """
    Inclusive business days (Mon–Fri).
    If start == end, force 1 day.
    If end < start, return -1 (invalid).
    """
    if d2 < d1:
        return -1
    if d1 == d2:
        return 1
    days = 0
    cur = d1
    while cur <= d2:
        if cur.weekday() < 5:  # Mon..Fri
            days += 1
        cur += timedelta(days=1)
    return max(days, 1)

def send_mail(to_addrs, subject, html, cc_addrs=None) -> bool:
    """SMTP send (Gmail app password recommended)."""
    host = _env("SMTP_HOST", "smtp.gmail.com")
    port = int(_env("SMTP_PORT", "587"))
    user = _env("SMTP_USER")
    pwd  = _env("SMTP_PASSWORD")
    mail_from = _env("MAIL_FROM", user)

    if not user or not pwd:
        app.logger.error("❌ SMTP credentials missing (SMTP_USER/SMTP_PASSWORD)")
        return False

    if isinstance(to_addrs, str):
        to_addrs = [to_addrs]
    if cc_addrs and isinstance(cc_addrs, str):
        cc_addrs = [cc_addrs]

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = mail_from or user
    msg["To"] = ", ".join(to_addrs)
    if cc_addrs:
        msg["Cc"] = ", ".join(cc_addrs)
    msg.set_content("HTML")
    msg.add_alternative(html, subtype="html")

    try:
        with smtplib.SMTP(host, port, timeout=30) as s:
            s.ehlo()
            s.starttls(context=ssl.create_default_context())
            s.login(user, pwd)
            s.send_message(msg)
        app.logger.info(f"✅ Mail sent to {to_addrs} (cc={cc_addrs})")
        return True
    except Exception as e:
        app.logger.exception(f"❌ SMTP send failed: {e}")
        return False

def get_sheets_client():
    """Create gspread client from GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_APPLICATION_CREDENTIALS."""
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    sa_json_text = _env("GOOGLE_SERVICE_ACCOUNT_JSON")
    sa_file_path = _env("GOOGLE_APPLICATION_CREDENTIALS")

    if sa_json_text:
        info = json.loads(sa_json_text)
        creds = Credentials.from_service_account_info(info, scopes=scopes)
    elif sa_file_path and os.path.isfile(sa_file_path):
        creds = Credentials.from_service_account_file(sa_file_path, scopes=scopes)
    else:
        raise RuntimeError("Service account credentials not provided.")

    return gspread.authorize(creds)

SHEET_HEADERS = ["Timestamp", "Name", "Email", "Start", "End", "Reason", "Days", "Approver", "Status"]

def write_decision_to_sheet(email:str, name:str, start_date:str, end_date:str,
                            reason:str, days:int, approver_email:str, status:str):
    """Always APPEND a new row (never overwrite)."""
    sheet_id = _env("SPREADSHEET_ID")
    if not sheet_id:
        app.logger.error("❌ SPREADSHEET_ID missing")
        return False

    gc = get_sheets_client()
    sh = gc.open_by_key(sheet_id)
    ws = sh.sheet1

    # Initialize headers if sheet is empty
    values = ws.get_all_values()
    if not values:
        ws.append_row(SHEET_HEADERS)

    row_data = [
        datetime.utcnow().isoformat() + "Z",
        name,
        email,
        start_date,
        end_date,
        reason or "—",
        days,
        approver_email,
        "APPROVED" if status == "approved" else "REJECTED",
    ]

    ws.append_row(row_data)
    app.logger.info(f"✅ Sheet appended new row for {email}")
    return True

# -------------------------------
# Routes
# -------------------------------
@app.route("/", methods=["GET"])
def form():
    html = render_template("form.html", logo_url=_env("LOGO_URL", ""))
    return Response(html, status=200, mimetype="text/html; charset=utf-8")

@app.route("/submit", methods=["POST"])
def submit():
    name       = request.form.get("name","").strip()
    email      = request.form.get("email","").strip()
    approver_k = request.form.get("approver")
    start_date = request.form.get("start_date")
    end_date   = request.form.get("end_date")
    reason     = request.form.get("reason","").strip()

    approver_email = APPROVERS.get(approver_k)
    if not approver_email:
        return Response("Unknown approver.", 400)

    # Dates + validation
    try:
        d1 = datetime.strptime(start_date, "%Y-%m-%d").date()
        d2 = datetime.strptime(end_date, "%Y-%m-%d").date()
    except Exception:
        return Response("Invalid date format. Please use YYYY-MM-DD.", 400)

    days = business_days(d1, d2)
    if days == -1:
        return Response("End date cannot be before start date.", 400)

    base_url = _env("BASE_URL", request.host_url.rstrip("/"))
    approve_link = f"{base_url}/decision?status=approved&email={email}&name={name}&sd={start_date}&ed={end_date}&reason={reason}"
    reject_link  = f"{base_url}/decision?status=rejected&email={email}&name={name}&sd={start_date}&ed={end_date}&reason={reason}"

    html = (
        f"<h2>New time off request</h2>"
        f"<p><b>Employee:</b> {name} &lt;{email}&gt;</p>"
        f"<p><b>Period:</b> {start_date} → {end_date} ({days} business day(s))</p>"
        f"<p><b>Reason:</b><br>{(reason or '—').replace('\n','<br>')}</p>"
        f"<p>"
        f"<a href='{approve_link}' style='display:inline-block;padding:10px 14px;background:#16a34a;color:#fff;border-radius:8px;text-decoration:none;'>✅ Approve</a>"
        f"&nbsp;&nbsp;"
        f"<a href='{reject_link}'  style='display:inline-block;padding:10px 14px;background:#dc2626;color:#fff;border-radius:8px;text-decoration:none;'>❌ Reject</a>"
        f"</p>"
    )

    cc = [_env("ALWAYS_CC", "mw@walton.fr")]
    ok = send_mail([approver_email], f"[{BRAND}] Time off request — {name}", html, cc_addrs=cc)

    if ok:
        ack = (
            f"<p>Hello {name},</p>"
            f"<p>Your request ({start_date} → {end_date}, {days} business day(s)) has been sent to {approver_email}.</p>"
            "<p>You will receive an email once a decision is made.</p>"
        )
        send_mail([email], f"[{BRAND}] Your request was sent", ack, cc_addrs=cc)
        html2 = render_template(
            "submitted.html",
            name=name,
            approver=approver_email,
            start=start_date,
            end=end_date,
            days=days,
            logo_url=_env("LOGO_URL", "")
        )
        return Response(html2, status=200, mimetype="text/html; charset=utf-8")
    else:
        return Response("Email send error ❌ — check logs.", 500)

@app.route("/decision", methods=["GET"])
def decision():
    status = request.args.get("status")  # "approved" | "rejected"
    email  = request.args.get("email","").strip()
    name   = request.args.get("name","").strip()
    sd     = request.args.get("sd")
    ed     = request.args.get("ed")
    reason = request.args.get("reason","")

    # Recompute days
    try:
        d1 = datetime.strptime(sd, "%Y-%m-%d").date()
        d2 = datetime.strptime(ed, "%Y-%m-%d").date()
        days = business_days(d1, d2)
    except Exception:
        days = 1

    # Write to Google Sheet only on approval
    if status == "approved":
        approver_email = _env("ALWAYS_CC", "mw@walton.fr")
        try:
            write_decision_to_sheet(
                email=email,
                name=name,
                start_date=sd,
                end_date=ed,
                reason=reason,
                days=days,
                approver_email=approver_email,
                status=status
            )
        except Exception as e:
            app.logger.exception(f"❌ Sheet write failed: {e}")

    # Decision email
    note = f"approved ✅ ({days} business day(s))" if status == "approved" else "rejected ❌"
    send_mail(
        [email],
        f"[{BRAND}] Decision — {status}",
        f"<p>Hello {name},</p><p>Your time off request {sd} → {ed} is {note}.</p>",
        cc_addrs=[_env("ALWAYS_CC", "mw@walton.fr")]
    )

    html = render_template(
        "decision_result.html",
        name=name, email=email, status=status, start=sd, end=ed, days=days,
        logo_url=_env("LOGO_URL", "")
    )
    return Response(html, status=200, mimetype="text/html; charset=utf-8")

# -------------------------------
# Admin (Basic Auth) : /admin
# -------------------------------
from functools import wraps
from base64 import b64decode

def check_auth(auth_header: str) -> bool:
    user = _env("ADMIN_USER")
    pwd  = _env("ADMIN_PASS")
    if not auth_header or not auth_header.startswith("Basic "):
        return False
    try:
        raw = b64decode(auth_header.split(" ",1)[1]).decode("utf-8")
        given_user, given_pwd = raw.split(":",1)
        return (user and pwd and given_user == user and given_pwd == pwd)
    except Exception:
        return False

def require_auth(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not check_auth(request.headers.get("Authorization")):
            return Response("Authentication required", 401, {"WWW-Authenticate":"Basic realm=\"Walton Admin\""})
        return f(*args, **kwargs)
    return wrapper

@app.get("/admin")
@require_auth
def admin():
    """Display requests from Google Sheets (sorted by timestamp desc)."""
    try:
        sheet_id = _env("SPREADSHEET_ID")
        gc = get_sheets_client()
        sh = gc.open_by_key(sheet_id)
        ws = sh.sheet1
        rows = ws.get_all_values()
        if not rows:
            rows = [SHEET_HEADERS]
        headers = rows[0]
        data = rows[1:]

        ts_idx = headers.index("Timestamp") if "Timestamp" in headers else None

        if ts_idx is not None:
            def parse_ts(v):
                try:
                    return datetime.fromisoformat(v.replace("Z",""))
                except Exception:
                    return datetime.min
            data.sort(key=lambda r: parse_ts(r[ts_idx]) if len(r) > ts_idx else datetime.min, reverse=True)

        table_rows = []
        for r in data:
            r = (r + [""]*9)[:9]
            table_rows.append(
                f"<tr><td>{r[0]}</td><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]}</td>"
                f"<td>{r[4]}</td><td>{r[5]}</td><td style='text-align:right'>{r[6]}</td>"
                f"<td>{r[7]}</td><td>{r[8]}</td></tr>"
            )

        html = f"""
        <!doctype html><html lang='en'><head>
          <meta charset='utf-8'>
          <meta name='viewport' content='width=device-width, initial-scale=1'>
          <title>Admin - {BRAND}</title>
          <style>
            body{{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;margin:0;background:#f7f7fb}}
            .container{{max-width:1100px;margin:24px auto;background:#fff;border-radius:12px;padding:24px;box-shadow:0 6px 24px rgba(0,0,0,.06)}}
            h1{{margin-top:0}}
            table{{width:100%;border-collapse:collapse}}
            th,td{{padding:10px;border-bottom:1px solid #eee;font-size:14px}}
            th{{text-align:left;background:#fafafa}}
          </style>
        </head><body>
          <div class='container'>
            <h1>Admin — Time Off Requests</h1>
            <p>Sorted from newest to oldest. Total: {len(data)} rows.</p>
            <table>
              <thead>
                <tr>
                  <th>Timestamp</th><th>Name</th><th>Email</th><th>Start</th><th>End</th><th>Reason</th><th>Days</th><th>Approver</th><th>Status</th>
                </tr>
              </thead>
              <tbody>
                {''.join(table_rows)}
              </tbody>
            </table>
          </div>
        </body></html>
        """
        return Response(html, status=200, mimetype="text/html; charset=utf-8")
    except Exception as e:
        app.logger.exception(e)
        return jsonify(ok=False, error=str(e)), 500)

# -------------------------------
# Diagnostics
# -------------------------------
@app.get("/_smtp_test")
def smtp_test():
    """SMTP test."""
    to = _env("SMTP_USER")
    ok = send_mail([to], f"[{BRAND}] SMTP Test", "<p>SMTP test OK ✅</p>")
    if ok:
        return jsonify(ok=True, to=to)
    return jsonify(ok=False, error="SMTP send failed (see logs)"), 500

@app.get("/_sheet_test")
def sheet_test():
    """Google Sheet access test + init header if empty."""
    try:
        sheet_id = _env("SPREADSHEET_ID")
        if not sheet_id:
            return jsonify(ok=False, error="SPREADSHEET_ID missing"), 400
        gc = get_sheets_client()
        sh = gc.open_by_key(sheet_id)
        ws = sh.sheet1
        vals = ws.get_all_values()
        if not vals:
            ws.append_row(SHEET_HEADERS)
        return jsonify(ok=True, sheet=sh.title, rows=len(ws.get_all_values()))
    except Exception as e:
        app.logger.exception(e)
        return jsonify(ok=False, error=str(e)), 500

@app.route("/healthz")
def health():
    return "ok", 200

@app.route("/ping")
def ping():
    return "pong", 200

# -------------------------------
# Run (local)
# -------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(_env("PORT", "10000")))
