import os
import json
import smtplib
from email.message import EmailMessage
from datetime import datetime, date, timedelta

from flask import (
    Flask, render_template, request, jsonify,
    Response
)

import gspread
from google.oauth2.service_account import Credentials

# -----------------------------------------------------------------------------
# Flask
# -----------------------------------------------------------------------------
app = Flask(__name__, static_folder="static")

# -----------------------------------------------------------------------------
# ENV CONFIG
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

APPROVER_EMAILS = {
    "mark": EMAIL_MARK,
    "nhan": EMAIL_NHAN,
    "anh": EMAIL_ANH,
}

APPROVER_LABELS = {
    "mark": "Mark",
    "nhan": "Nhàn",
    "anh": "Anh",
}


# -----------------------------------------------------------------------------
# GOOGLE SHEET
# -----------------------------------------------------------------------------
def get_sheet():
    if not SPREADSHEET_ID or not GOOGLE_SERVICE_ACCOUNT_JSON:
        raise RuntimeError("Google Sheet credentials missing")

    info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).sheet1


# -----------------------------------------------------------------------------
# BUSINESS DAYS
# -----------------------------------------------------------------------------
def business_days(start: date, end: date) -> float:
    if end < start:
        return -1
    days = 0
    cur = start
    while cur <= end:
        if cur.weekday() < 5:  # Monday–Friday
            days += 1
        cur += timedelta(days=1)
    return float(days)


def adjust_half_day(days: float, duration_type: str) -> float:
    if duration_type == "half":
        return 0.5
    return days


# -----------------------------------------------------------------------------
# EMAIL SENDER
# -----------------------------------------------------------------------------
def send_email(subject: str, to_list, body_html: str, cc_list=None):
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
    msg.add_alternative(body_html, subtype="html")

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(SMTP_USER, SMTP_PASSWORD)
        s.send_message(msg)


# -----------------------------------------------------------------------------
# ROUTES
# -----------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("form.html")


# -----------------------------------------------------------------------------
@app.route("/submit", methods=["POST"])
def submit():
    name = request.form.get("name", "").strip()
    email = request.form.get("email", "").strip()
    approver_k = request.form.get("approver")
    start_date = request.form.get("start_date")
    end_date = request.form.get("end_date")
    reason = request.form.get("reason", "").strip()
    duration_type = request.form.get("duration_type", "full")

    if not (name and email and approver_k and start_date and end_date):
        return Response("Missing fields", 400)

    approver_email = APPROVER_EMAILS.get(approver_k, EMAIL_MARK)
    approver_label = APPROVER_LABELS.get(approver_k, approver_k)

    # Parse dates
    try:
        d1 = datetime.strptime(start_date, "%Y-%m-%d").date()
        d2 = datetime.strptime(end_date, "%Y-%m-%d").date()
    except Exception:
        return Response("Invalid date format", 400)

    days = business_days(d1, d2)
    if days == -1:
        return Response("End date cannot be before start date.", 400)

    # Half day – only allowed if same start/end
    if duration_type == "half" and d1 != d2:
        return Response("Half day allowed only when start and end dates match.", 400)

    days = adjust_half_day(days, duration_type)

    # Build approve/reject links
    approve_link = (
        f"{BASE_URL}/decision?status=approved&email={email}"
        f"&name={name}&sd={start_date}&ed={end_date}&reason={reason}"
        f"&dt={duration_type}"
    )
    reject_link = (
        f"{BASE_URL}/decision?status=rejected&email={email}"
        f"&name={name}&sd={start_date}&ed={end_date}&reason={reason}"
        f"&dt={duration_type}"
    )

    # Email approver
    html = f"""
    <p>New time off request:</p>
    <ul>
      <li><b>Name:</b> {name}</li>
      <li><b>Email:</b> {email}</li>
      <li><b>Period:</b> {start_date} → {end_date}</li>
      <li><b>Duration:</b> {days} business day(s)</li>
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

    # Write to sheet (Pending)
    try:
        sheet = get_sheet()
        sheet.append_row([
            datetime.utcnow().isoformat(timespec="seconds"),
            name,           # Name
            email,          # Email
            approver_label, # Approver (Mark / Nhàn / Anh)
            start_date,
            end_date,
            days,
            duration_type,
            reason,
            "Pending",      # Status
        ])
    except Exception as e:
        print("Sheet error on submit:", e)

    return render_template(
        "submitted.html",
        approver=approver_label,
        start=start_date,
        end=end_date,
        days=days,
    )


# -----------------------------------------------------------------------------
@app.route("/decision")
def decision():
    status = request.args.get("status")
    email = request.args.get("email", "")
    name = request.args.get("name", "")
    sd = request.args.get("sd")
    ed = request.args.get("ed")
    reason = request.args.get("reason", "")
    duration_type = request.args.get("dt", "full")

    if status not in ("approved", "rejected"):
        return Response("Invalid status", 400)

    try:
        d1 = datetime.strptime(sd, "%Y-%m-%d").date()
        d2 = datetime.strptime(ed, "%Y-%m-%d").date()
        days = business_days(d1, d2)
        days = adjust_half_day(days, duration_type)
    except Exception:
        days = 1.0

    # Update last Pending row in sheet
    try:
        sheet = get_sheet()
        rows = sheet.get_all_values()

        pending_row = None
        # Start from bottom (most recent)
        for idx in range(len(rows) - 1, 0, -1):
            row = rows[idx]
            if len(row) < 10:
                continue
            row_email = row[2]
            row_start = row[4]
            row_end = row[5]
            row_status = row[9]
            if row_email == email and row_start == sd and row_end == ed and row_status == "Pending":
                pending_row = idx + 1  # gspread is 1-based
                break

        if pending_row:
            sheet.update_cell(pending_row, 10, status.capitalize())
        else:
            # Fallback: append a new row if we didn't find the pending one
            sheet.append_row([
                datetime.utcnow().isoformat(timespec="seconds"),
                name,
                email,
                "Decision",
                sd,
                ed,
                days,
                duration_type,
                reason,
                status.capitalize(),
            ])

    except Exception as e:
        print("Decision sheet error:", e)

    # Email requester
    decision_txt = "approved ✅" if status == "approved" else "rejected ❌"
    html = f"""
    <p>Hi {name},</p>
    <p>Your time off request has been <b>{decision_txt}</b>.</p>
    <ul>
      <li>Period: {sd} → {ed}</li>
      <li>Duration: {days} business day(s)</li>
    </ul>
    """

    try:
        send_email(
            subject=f"Time off request – {decision_txt}",
            to_list=[email],
            body_html=html,
            cc_list=[ALWAYS_CC] if ALWAYS_CC else None,
        )
    except Exception as e:
        print("Decision email error:", e)

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
    app.run(debug=True)
