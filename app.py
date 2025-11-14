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

def get_staff_sheet():
    """
    Return the 'Staff List' worksheet (must exist in the same spreadsheet).
    Columns: Name | Email
    """
    sh = get_sheet().spreadsheet  # on récupère le classeur déjà utilisé
    try:
        ws = sh.worksheet("Staff List")
    except gspread.WorksheetNotFound:
        # Si la feuille n’existe pas, on la crée avec les colonnes
        ws = sh.add_worksheet(title="Staff List", rows=100, cols=2)
        ws.append_row(["Name", "Email"])
    return ws


def load_staff_list():
    """
    Returns a list of dicts: [{"name": ..., "email": ...}, ...]
    pulled from 'Staff List' sheet.
    """
    ws = get_staff_sheet()
    records = ws.get_all_records()  # lit toutes les lignes sous l’en-tête
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
# EMAIL SENDER
# -----------------------------------------------------------------------------
def send_email(subject: str, to_list, body_html: str, cc_list=None):
    # Vérifier la config SMTP
    if not SMTP_USER or not SMTP_PASSWORD:
        raise RuntimeError("SMTP configuration missing")

    # Normaliser les listes
    if isinstance(to_list, str):
        to_list = [to_list]
    if cc_list is None:
        cc_list = []
    elif isinstance(cc_list, str):
        cc_list = [cc_list]

    # Filtrer les adresses vides / None
    to_list = [str(addr).strip() for addr in to_list if addr]
    cc_list = [str(addr).strip() for addr in cc_list if addr]

    if not to_list:
        raise RuntimeError("No valid recipient email in to_list")

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



# -----------------------------------------------------------------------------
@app.route("/submit", methods=["POST"])
def submit():
    form = request.form

    employee_name = form.get("employee_name")  # vient de la liste déroulante
    approver      = form.get("approver")
    start_str     = form.get("start_date")
    end_str       = form.get("end_date")
    duration_type = form.get("duration_type")  # "full" ou "half"
    type_of_leave = form.get("type_of_leave")  # nouvelle info
    reason        = form.get("reason", "").strip()

    # On recalcule l’email à partir de la Staff List (sécurité)
    staff = {m["name"]: m["email"] for m in load_staff_list()}
    employee_email = staff.get(employee_name)

    if not employee_email:
        # fallback si problème dans la liste (optionnel)
        return "Employee email not found in Staff List", 400

    # Convertir les dates + calcul du nombre de jours ouvrés (ton code existant)
    d1 = datetime.strptime(start_str, "%Y-%m-%d").date()
    d2 = datetime.strptime(end_str, "%Y-%m-%d").date()
    days = business_days(d1, d2)

    # Half-day → 0.5
    if duration_type == "half":
        days = 0.5

    # Ajout de la ligne dans le Google Sheet 'Requests'
    sheet = get_sheet()
    timestamp = datetime.utcnow().isoformat()

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
    # -------------------------
    # Send notification emails
    # -------------------------
      # -------------------------
    # Send notification emails
    # -------------------------
    try:
        # Déterminer l'email de l'approbateur
        approver_email_map = {
            "Mark": EMAIL_MARK,
            "Nhàn": EMAIL_NHAN,
            "Nhan": EMAIL_NHAN,
            "Anh": EMAIL_ANH,
        }
        approver_email = approver_email_map.get(approver, EMAIL_MARK)

        # Construire les liens Approve / Reject
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

        # 1) Email à l'approbateur avec les liens
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

        # 2) Email de confirmation à l'employé
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

        # 3) Copie systématique à ALWAYS_CC (si défini)
        if ALWAYS_CC:
            send_email(
                subject="Walton Time Off – Copy of request",
                to_list=[ALWAYS_CC],
                body_html=approver_body,
            )

    except Exception as e:
        print("EMAIL ERROR in /submit:", e)
        # On ne casse pas la réponse utilisateur, on log juste l'erreur



    # Email, etc. → tu peux garder ton code existant ici

    return render_template(
        "submitted.html",
        approver=approver,
        start=start_str,
        end=end_str,
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
@app.route("/admin/reset-sheet", methods=["GET"])
def reset_sheet():
    try:
        sheet = get_sheet()
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
