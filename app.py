import os
import json
import smtplib
from email.message import EmailMessage
from datetime import datetime, date, timedelta

from flask import (
    Flask, render_template, request, jsonify,
    Response, url_for
)

import gspread
from google.oauth2.service_account import Credentials

# -----------------------------------------------------------------------------
# Flask app (static_folder='static' to serve /static/*)
# -----------------------------------------------------------------------------
app = Flask(__name__, static_folder="static")

# -----------------------------------------------------------------------------
# Config from environment
# -----------------------------------------------------------------------------
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
MAIL_FROM = os.getenv("MAIL_FROM", "Walton Time Off <no-reply@example.com>")

EMAIL_MARK = os.getenv("EMAIL_MARK", "mw@walton.fr")
EMAIL_NHAN = os.getenv("EMAIL_NHAN", "nhan@walton.fr")
EMAIL_ANH = os.getenv("EMAIL_ANH", "anh@walton.fr")
ALWAYS_CC = os.getenv("ALWAYS_CC", "mw@walton.fr")

BASE_URL = os.getenv("BASE_URL", "https://walton-holidays.onrender.com")

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")

# -----------------------------------------------------------------------------
# Helpers: Google Sheet
# -----------------------------------------------------------------------------
def get_sheet():
    # Return first worksheet of the configured Google Sheet.
    if not SPREADSHEET_ID or not GOOGLE_SERVICE_ACCOUNT_JSON:
        raise RuntimeError("Google Sheet credentials missing")

    info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).sheet1
    return sheet

# -----------------------------------------------------------------------------
# Helpers: business days (Mon–Fri)
# -----------------------------------------------------------------------------
def business_days(start: date, end: date) -> float:
    # Return number of business days between start and end (inclusive).
    # Returns -1 if end < start.
    if end < start:
        return -1
    days = 0
    cur = start
    while cur <= end:
        if cur.weekday() < 5:  # 0=Mon .. 4=Fri
            days += 1
        cur += timedelta(days=1)
    return float(days)

def adjust_half_day(days: float, duration_type: str) -> float:
    # Adjust number of days for half-day option.
    if duration_type == "half":
        if days <= 1:
            return 0.5
        return days - 0.5
    return days

# -----------------------------------------------------------------------------
# Helpers: email
# -----------------------------------------------------------------------------
def send_email(subject: str, to_list, body_html: str, cc_list=None):
    if not SMTP_HOST or not SMTP_USER or not SMTP_PASSWORD:
        raise RuntimeError("SMTP configuration missing")

    if isinstance(to_list, str):
        to_list = [to_list]
    if cc_list is None:
        cc_list = []
    elif isinstance(cc_list, str):
        cc_list = [cc_list]

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = MAIL_FROM
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content("HTML email only")
    msg.add_alternative(body_html, subtype="html")

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(SMTP_USER, SMTP_PASSWORD)
        s.send_message(msg)

# -----------------------------------------------------------------------------
# Routes
# -----------------------------------------------------------------------------
@app.route("/", methods=["GET"])
def index():
    # Main form page.
    return render_template("form.html")

# -----------------------------------------------------------------------------
@app.route("/submit", methods=["POST"])
def submit():
    # Receive request form, send email to approver, log in Google Sheet.
    name       = request.form.get("name", "").strip()
    email      = request.form.get("email", "").strip()
    approver_k = request.form.get("approver")
    start_date = request.form.get("start_date")
    end_date   = request.form.get("end_date")
    reason     = request.form.get("reason", "").strip()
    duration_type = request.form.get("duration_type", "full")  # full | half

    if not (name and email and approver_k and start_date and end_date):
        return Response("Missing fields", 400)

    # Pick approver email
    approver_email = {
        "mark": EMAIL_MARK,
        "nhan": EMAIL_NHAN,
        "anh": EMAIL_ANH,
    }.get(approver_k, EMAIL_MARK)

    # Parse dates
    try:
        d1 = datetime.strptime(start_date, "%Y-%m-%d").date()
        d2 = datetime.strptime(end_date, "%Y-%m-%d").date()
    except Exception:
        return Response("Invalid date format", 400)

    days = business_days(d1, d2)
    if days == -1:
        return Response("End date cannot be before start date.", 400)

    days = adjust_half_day(days, duration_type)

    # Build approval links
    approve_link = (
        f"{BASE_URL}/decision?"
        f"status=approved&email={email}&name={name}"
        f"&sd={start_date}&ed={end_date}&reason={reason}"
        f"&dt={duration_type}"
    )
    reject_link = (
        f"{BASE_URL}/decision?"
        f"status=rejected&email={email}&name={name}"
        f"&sd={start_date}&ed={end_date}&reason={reason}"
        f"&dt={duration_type}"
    )

    # Simple HTML email to approver
    html = f"""
    <p>New time off request:</p>
    <ul>
      <li><b>Name:</b> {name}</li>
      <li><b>Email:</b> {email}</li>
      <li><b>Period:</b> {start_date} → {end_date}</li>
      <li><b>Duration:</b> {days} business day(s) ({duration_type})</li>
      <li><b>Reason:</b> {reason or "—"}</li>
    </ul>
    <p>
      <a href="{approve_link}">✅ Approve</a> |
      <a href="{reject_link}">❌ Reject</a>
    </p>
    """

    send_email(
        subject=f"Time off request – {name}",
        to_list=[approver_email],
        body_html=html,
        cc_list=[ALWAYS_CC] if ALWAYS_CC else None,
    )

    # Log in Sheet (append)
    try:
        sheet = get_sheet()
        sheet.append_row([
            datetime.utcnow().isoformat(timespec="seconds") + "Z",
            name,
            email,
            approver_k,
            start_date,
            end_date,
            days,
            duration_type,
            reason,
            "Pending",  # status
        ])
    except Exception as e:
        print("Error writing to sheet:", e)

    # Show confirmation page
    return render_template(
        "submitted.html",
        approver=approver_k.capitalize(),
        start=start_date,
        end=end_date,
        days=days,
    )

# -----------------------------------------------------------------------------
@app.route("/decision", methods=["GET"])
def decision():
    # Approver clicks approve / reject link.
    status = request.args.get("status")  # approved | rejected
    email  = request.args.get("email", "").strip()
    name   = request.args.get("name", "").strip()
    sd     = request.args.get("sd")
    ed     = request.args.get("ed")
    reason = request.args.get("reason", "")
    duration_type = request.args.get("dt", "full")

    if status not in ("approved", "rejected"):
        return Response("Invalid status", 400)

    # Recompute days
    try:
        d1 = datetime.strptime(sd, "%Y-%m-%d").date()
        d2 = datetime.strptime(ed, "%Y-%m-%d").date()
        days = business_days(d1, d2)
        days = adjust_half_day(days, duration_type)
    except Exception:
        days = 1.0

    # Append decision in Sheet (simple, non-destructive)
    try:
        sheet = get_sheet()
        sheet.append_row([
            datetime.utcnow().isoformat(timespec="seconds") + "Z",
            name,
            email,
            "DECISION",
            sd,
            ed,
            days,
            duration_type,
            reason,
            status.capitalize(),
        ])
    except Exception as e:
        print("Error writing decision to sheet:", e)

    # Email to requester
    decision_text = "approved ✅" if status == "approved" else "rejected ❌"
    html = f"""
    <p>Hi {name},</p>
    <p>Your time off request has been <b>{decision_text}</b>.</p>
    <ul>
      <li>Period: {sd} → {ed}</li>
      <li>Duration: {days} business day(s)</li>
    </ul>
    <p>If you have any questions, please contact HR.</p>
    """

    try:
        send_email(
            subject=f"Time off request – {decision_text}",
            to_list=[email],
            body_html=html,
            cc_list=[ALWAYS_CC] if ALWAYS_CC else None,
        )
    except Exception as e:
        print("Error sending decision email:", e)

    return render_template(
        "decision_result.html",
        name=name,
        email=email,
        start=sd,
        end=ed,
        days=days,
        status=status,
    )

# -----------------------------------------------------------------------------
@app.route("/_smtp_test")
def smtp_test():
    # Quick healthcheck for SMTP.
    try:
        send_email(
            subject="Walton Time Off – SMTP test",
            to_list=[ALWAYS_CC or EMAIL_MARK],
            body_html="<p>SMTP test OK.</p>",
        )
        return jsonify(ok=True)
    except Exception as e:
        print("SMTP test error:", e)
        return jsonify(ok=False, error=str(e)), 500

# -----------------------------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
