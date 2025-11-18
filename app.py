import os
import json
import smtplib
from email.message import EmailMessage
from datetime import datetime, date, timedelta
from urllib.parse import quote_plus

from flask import (
    Flask, render_template, request, jsonify,
    Response
)

import gspread
from google.oauth2.service_account import Credentials

# -----------------------------------------------------------------------------
# Flask app
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


# -----------------------------------------------------------------------------
# GOOGLE SHEETS HELPERS
# -----------------------------------------------------------------------------
def get_gspread_client():
    if not SPREADSHEET_ID or not GOOGLE_SERVICE_ACCOUNT_JSON:
        raise RuntimeError("Google Sheet credentials missing")

    info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    client = gspread.authorize(creds)
    return client


def get_workbook():
    client = get_gspread_client()
    return client.open_by_key(SPREADSHEET_ID)


def get_requests_sheet():
    """
    Main sheet used for storing requests.
    Name: 'Requests'
    """
    wb = get_workbook()
    try:
        ws = wb.worksheet("Requests")
    except gspread.WorksheetNotFound:
        ws = wb.add_worksheet(title="Requests", rows=1000, cols=12)
        header = [
            "Timestamp",
            "Name",
            "Email",
            "Approver",
            "Start",
            "End",
            "Days",
            "Duration",
            "Type of leave",
            "Reason",
            "Status",
        ]
        ws.append_row(header)
    return ws


def get_staff_sheet():
    """
    Sheet 'Staff List' with columns: Name | Email
    """
    wb = get_workbook()
    try:
        ws = wb.worksheet("Staff List")
    except gspread.WorksheetNotFound:
        ws = wb.add_worksheet(title="Staff List", rows=100, cols=2)
        ws.append_row(["Name", "Email"])
    return ws


def load_staff_list():
    """
    Returns a list: [{"name": ..., "email": ...}, ...]
    from 'Staff List'.
    """
    ws = get_staff_sheet()
    records = ws.get_all_records()
    staff = []
    for row in records:
        name = row.get("Name") or row.get("name")
        email = row.get("Email") or row.get("email")
        if name and email:
            staff.append({"name": name, "email": email})
    return staff


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
# EMAIL
# -----------------------------------------------------------------------------
def send_email(subject: str, to_list, body_html: str, cc_list=None):
    if not SMTP_USER or not SMTP_PASSWORD:
        raise RuntimeError("SMTP configuration missing")

    if isinstance(to_list, str):
        to_list = [to_list]
    if cc_list is None:
        cc_list = []
    elif isinstance(cc_list, str):
        cc_list = [cc_list]

    # Nettoyer les adresses
    to_list = [str(a).strip() for a in to_list if a]
    cc_list = [str(a).strip() for a in cc_list if a]

    if not to_list:
        raise RuntimeError("No valid recipient in to_list")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = MAIL_FROM
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content("HTML only")
    msg.add_alternative(body_html, subtype="html")

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(SMTP_USER, SMTP_PASSWORD)
        s.send_message(msg)


# -----------------------------------------------------------------------------
# ROUTES
# -----------------------------------------------------------------------------
@app.route("/", methods=["GET"])
def index():
    staff = load_staff_list()
    return render_template("form.html", staff=staff)


@app.route("/submit", methods=["POST"])
def submit():
    form = request.form

    employee_name = form.get("employee_name")
    approver = form.get("approver")
    start_str = form.get("start_date")
    end_str = form.get("end_date")
    duration_type = form.get("duration_type", "full")
    type_of_leave = form.get("type_of_leave")
    reason = form.get("reason", "").strip()

    if not (employee_name and approver and start_str and end_str and type_of_leave):
        return Response("Missing fields", 400)

    # Récupérer email de l’employé depuis Staff List
    staff_map = {m["name"]: m["email"] for m in load_staff_list()}
    employee_email = staff_map.get(employee_name)
    if not employee_email:
        return Response("Employee email not found in Staff List", 400)

    # Dates + calcul jours
    try:
        d1 = datetime.strptime(start_str, "%Y-%m-%d").date()
        d2 = datetime.strptime(end_str, "%Y-%m-%d").date()
    except Exception:
        return Response("Invalid date format", 400)

    days = business_days(d1, d2)
    if days == -1:
        return Response("End date cannot be before start date.", 400)

    # Half day seulement si même date
    if duration_type == "half" and d1 != d2:
        return Response("Half day allowed only when start and end dates match.", 400)

    days = adjust_half_day(days, duration_type)

    # Enregistrement dans Requests
    sheet = get_requests_sheet()
    timestamp = datetime.utcnow().isoformat(timespec="seconds")

    row = [
        timestamp,
        employee_name,
        employee_email,
        approver,
        start_str,
        end_str,
        days,
        duration_type,
        type_of_leave,
        reason,
        "Pending",
    ]
    sheet.append_row(row)

    # Emails
    try:
        approver_email_map = {
            "Mark": EMAIL_MARK,
            "Nhàn": EMAIL_NHAN,
            "Nhan": EMAIL_NHAN,
            "Anh": EMAIL_ANH,
        }
        approver_email = approver_email_map.get(approver, EMAIL_MARK)

        # Liens approve / reject
        approve_link = (
            f"{BASE_URL}/decision"
            f"?status=approved"
            f"&email={quote_plus(employee_email)}"
            f"&name={quote_plus(employee_name)}"
            f"&sd={start_str}"
            f"&ed={end_str}"
            f"&reason={quote_plus(reason)}"
            f"&dt={duration_type}"
        )
        reject_link = (
            f"{BASE_URL}/decision"
            f"?status=rejected"
            f"&email={quote_plus(employee_email)}"
            f"&name={quote_plus(employee_name)}"
            f"&sd={start_str}"
            f"&ed={end_str}"
            f"&reason={quote_plus(reason)}"
            f"&dt={duration_type}"
        )

        # 1) mail à l’approbateur
        approver_body = f"""
        <p>Hello {approver},</p>
        <p>You have a new time off request:</p>
        <ul>
          <li><b>Employee</b>: {employee_name} ({employee_email})</li>
          <li><b>Dates</b>: {start_str} → {end_str}</li>
          <li><b>Days</b>: {days}</li>
          <li><b>Duration</b>: {duration_type}</li>
          <li><b>Type of leave</b>: {type_of_leave}</li>
          <li><b>Reason</b>: {reason or "—"}</li>
        </ul>
        <p>
          <a href="{approve_link}">✅ Approve</a> |
          <a href="{reject_link}">❌ Reject</a>
        </p>
        """

        send_email(
            subject="Walton Time Off – New request",
            to_list=[approver_email],
            body_html=approver_body,
        )

        # 2) mail à l’employé
        employee_body = f"""
        <p>Hello {employee_name},</p>
        <p>Your time off request has been received.</p>
        <ul>
          <li><b>Dates</b>: {start_str} → {end_str}</li>
          <li><b>Days</b>: {days}</li>
          <li><b>Duration</b>: {duration_type}</li>
          <li><b>Type of leave</b>: {type_of_leave}</li>
        </ul>
        <p>You will receive an update once a decision is made.</p>
        """

        send_email(
            subject="Walton Time Off – Request received",
            to_list=[employee_email],
            body_html=employee_body,
        )

     

    except Exception as e:
        print("EMAIL ERROR in /submit:", e)

    return render_template(
        "submitted.html",
        approver=approver,
        start=start_str,
        end=end_str,
        days=days,
    )


@app.route("/decision", methods=["GET"])
def decision():
    status = request.args.get("status")  # approved / rejected
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

    # Met à jour la dernière ligne Pending correspondante
    try:
        sheet = get_requests_sheet()
        rows = sheet.get_all_values()

        pending_row = None
        for idx in range(len(rows) - 1, 0, -1):
            row = rows[idx]
            if len(row) < 11:
                continue
            row_email = row[2]
            row_start = row[4]
            row_end = row[5]
            row_status = row[10]
            if row_email == email and row_start == sd and row_end == ed and row_status == "Pending":
                pending_row = idx + 1  # gspread 1-based
                break

        if pending_row:
            sheet.update_cell(pending_row, 11, status.capitalize())
        else:
            sheet.append_row([
                datetime.utcnow().isoformat(timespec="seconds"),
                name,
                email,
                "Decision",
                sd,
                ed,
                days,
                duration_type,
                "",         # Type of leave inconnu ici
                reason,
                status.capitalize(),
            ])

    except Exception as e:
        print("Decision sheet error:", e)

    # Mail à l’employé
    decision_txt = "approved ✅" if status == "approved" else "rejected ❌"
    body = f"""
    <p>Hi {name},</p>
    <p>Your time off request has been <b>{decision_txt}</b>.</p>
    <ul>
      <li>Period: {sd} → {ed}</li>
      <li>Duration: {days} business day(s)</li>
    </ul>
    """

    try:
        send_email(
            subject=f"Walton Time Off – Request {status}",
            to_list=[email],
cc_list=[ALWAYS_CC],
            body_html=body,
        )
    except Exception as e:
        print("EMAIL ERROR in /decision:", e)

    return render_template(
        "decision_result.html",
        name=name,
        email=email,
        start=sd,
        end=ed,
        days=days,
        status=status,
    )


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


@app.route("/admin/reset-sheet", methods=["GET"])
def reset_sheet():
    """
    Clear the Requests sheet and recreate the header.
    """
    try:
        sheet = get_requests_sheet()
        sheet.clear()

        header = [
            "Timestamp",
            "Name",
            "Email",
            "Approver",
            "Start",
            "End",
            "Days",
            "Duration",
            "Type of leave",
            "Reason",
            "Status",
        ]
        sheet.append_row(header)

        return jsonify(ok=True, message="Sheet reset and header recreated.")
    except Exception as e:
        print("reset-sheet error:", e)
        return jsonify(ok=False, error=str(e)), 500


# -----------------------------------------------------------------------------
if __name__ == "__main__":
    app.run(debug=True)
