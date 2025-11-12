import os
import json
import smtplib
import ssl
from email.message import EmailMessage
from datetime import datetime, date, timedelta

from flask import Flask, render_template, request, jsonify, Response

# (Google Sheets)
import gspread
from google.oauth2.service_account import Credentials

# -----------------------------------------------------------------------------
# Flask
# -----------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")

app = Flask(__name__, template_folder=TEMPLATE_DIR)

# -----------------------------------------------------------------------------
# Config approbateurs via variables d'env (avec valeurs par défaut)
# -----------------------------------------------------------------------------
APPROVERS = {
    "mark": os.environ.get("EMAIL_MARK", "mw@walton.fr"),
    "nhan": os.environ.get("EMAIL_NHAN", "nhan@walton.fr"),
    "anh":  os.environ.get("EMAIL_ANH",  "anh@walton.fr"),
}

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def _env(key, default=None):
    return os.environ.get(key, default)

def business_days(d1: date, d2: date) -> int:
    """Nombre de jours ouvrés inclusifs (lun–ven), sans jours fériés."""
    if d2 < d1:
        d1, d2 = d2, d1
    days = 0
    cur = d1
    while cur <= d2:
        if cur.weekday() < 5:  # 0=Mon .. 6=Sun
            days += 1
        cur += timedelta(days=1)
    return days

def send_mail(to_addrs, subject, html, cc_addrs=None) -> bool:
    """Envoi SMTP via Gmail (App Password)."""
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
    """
    Retourne un client gspread en utilisant :
    - soit GOOGLE_SERVICE_ACCOUNT_JSON (contenu JSON),
    - soit GOOGLE_APPLICATION_CREDENTIALS (chemin vers un fichier JSON sur le disque).
    """
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

SHEET_HEADERS = ["Timestamp", "Nom", "Email", "Début", "Fin", "Raison", "Jours", "Approbateur", "Statut"]

def write_decision_to_sheet(email:str, name:str, start_date:str, end_date:str,
                            reason:str, days:int, approver_email:str, status:str):
    """Ajoute ou met à jour la ligne pour l'email donné dans Google Sheet."""
    sheet_id = _env("SPREADSHEET_ID")
    if not sheet_id:
        app.logger.error("❌ SPREADSHEET_ID manquant")
        return False

    gc = get_sheets_client()
    sh = gc.open_by_key(sheet_id)
    ws = sh.sheet1

    # Initialise l'entête si la feuille est vide
    values = ws.get_all_values()
    if not values:
        ws.append_row(SHEET_HEADERS)
        values = [SHEET_HEADERS]

    # Cherche une ligne existante pour cet email (colonne "Email")
    try:
        header = values[0]
    except IndexError:
        header = []
    try:
        email_idx = header.index("Email") + 1  # gspread est 1-based
    except ValueError:
        # Si les entêtes ne sont pas conformes, on les impose
        ws.clear()
        ws.append_row(SHEET_HEADERS)
        email_idx = SHEET_HEADERS.index("Email") + 1
        values = [SHEET_HEADERS]

    # trouve la première ligne où la colonne Email == email
    target_row = None
    for i, row in enumerate(values[1:], start=2):  # ligne 1 = entêtes
        if len(row) >= email_idx and row[email_idx - 1].strip().lower() == email.strip().lower():
            target_row = i
            break

    row_data = [
        datetime.utcnow().isoformat() + "Z",
        name,
        email,
        start_date,
        end_date,
        reason or "—",
        days,
        approver_email,
        "APPROUVÉ" if status == "approved" else "REFUSÉ",
    ]

    if target_row:
        # Met à jour toute la ligne (9 colonnes)
        ws.update(f"A{target_row}:I{target_row}", [row_data])
        app.logger.info(f"✅ Sheet updated row {target_row} for {email}")
    else:
        ws.append_row(row_data)
        app.logger.info(f"✅ Sheet appended new row for {email}")

    return True

# -----------------------------------------------------------------------------
# Routes
# -----------------------------------------------------------------------------
@app.route("/", methods=["GET"])
def form():
    return render_template("form.html")

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
        return "Approver inconnu", 400

    # Calcul jours ouvrés
    try:
        d1 = datetime.strptime(start_date, "%Y-%m-%d").date()
        d2 = datetime.strptime(end_date, "%Y-%m-%d").date()
        days = business_days(d1, d2)
    except Exception:
        days = 0

    base_url = _env("BASE_URL", request.host_url.rstrip("/"))
    approve_link = f"{base_url}/decision?status=approved&email={email}&name={name}&sd={start_date}&ed={end_date}&reason={reason}"
    reject_link  = f"{base_url}/decision?status=rejected&email={email}&name={name}&sd={start_date}&ed={end_date}&reason={reason}"

    html = f"""
      <h2>Nouvelle demande de congés</h2>
      <p><b>Demandeur :</b> {name} &lt;{email}&gt;</p>
      <p><b>Période :</b> {start_date} → {end_date} ({days} jours ouvrés)</p>
      <p><b>Raison :</b><br>{(reason or '—').replace('\n','<br>')}</p>
      <p>
        <a href="{approve_link}" style="display:inline-block;padding:10px 14px;background:#16a34a;color:#fff;border-radius:8px;text-decoration:none;">✅ Approuver</a>
        &nbsp;&nbsp;
        <a href="{reject_link}"  style="display:inline-block;padding:10px 14px;background:#dc2626;color:#fff;border-radius:8px;text-decoration:none;">❌ Refuser</a>
      </p>
    """

    cc = [_env("ALWAYS_CC", "mw@walton.fr")]
    ok = send_mail([approver_email], f"[Walton] Demande de congés — {name}", html, cc_addrs=cc)

    if ok:
        ack = f"""
          <p>Bonjour {name},</p>
          <p>Votre demande ({start_date} → {end_date}, {days} jours ouvrés) a été envoyée à {approver_email}.</p>
          <p>Vous recevrez un email dès décision.</p>
        """
        send_mail([email], "[Walton] Votre demande a été envoyée", ack, cc_addrs=cc)
        return render_template("submitted.html", name=name, approver=approver_email, start=start_date, end=end_date, days=days)
    else:
        return "Erreur d'envoi email ❌ — vérifiez les logs.", 500

@app.route("/decision", methods=["GET"])
def decision():
    status = request.args.get("status")  # "approved" | "rejected"
    email  = request.args.get("email","").strip()
    name   = request.args.get("name","").strip()
    sd     = request.args.get("sd")
    ed     = request.args.get("ed")
    reason = request.args.get("reason","")

    # Recalcule les jours ouvrés à l’écriture
    try:
        d1 = datetime.strptime(sd, "%Y-%m-%d").date()
        d2 = datetime.strptime(ed, "%Y-%m-%d").date()
        days = business_days(d1, d2)
    except Exception:
        days = 0

    # Écrit dans Google Sheet si approuvé
    if status == "approved":
        approver_email = _env("EMAIL_MARK", "mw@walton.fr") if "mark" in (request.args.get("approver","") or "").lower() else None
        # Si on n'a pas l'info "approver" en paramètre, note simplement l'email du décideur comme ALWAYS_CC
        approver_email = approver_email or _env("ALWAYS_CC", "mw@walton.fr")
        try:
            write_decision_to_sheet(email=email, name=name, start_date=sd, end_date=ed,
                                    reason=reason, days=days, approver_email=approver_email, status=status)
        except Exception as e:
            app.logger.exception(f"❌ Sheet write failed: {e}")

    # Notifie le demandeur de la décision
    note = f"approuvée ✅ ({days} jours ouvrés)" if status == "approved" else "refusée ❌"
    send_mail(
        [email],
        f"[Walton] Décision — {status}",
        f"<p>Bonjour {name},</p><p>Votre demande de congés {sd} → {ed} est {note}.</p>",
        cc_addrs=[_env("ALWAYS_CC", "mw@walton.fr")]
    )
    return render_template("decision_result.html", name=name, email=email, status=status, start=sd, end=ed, days=days)

# -----------------------------------------------------------------------------
# Diagnostics
# -----------------------------------------------------------------------------
@app.get("/_smtp_test")
def smtp_test():
    """Teste l'envoi SMTP avec les variables d'environnement actuelles."""
    to = _env("SMTP_USER")
    ok = send_mail([to], "[Walton] SMTP Test", "<p>Test d'envoi réussi ✅</p>")
    if ok:
        return jsonify(ok=True, to=to)
    return jsonify(ok=False, error="SMTP send failed (voir logs Render)"), 500

@app.get("/_sheet_test")
def sheet_test():
    """Teste l'accès au Google Sheet et initialise l'entête si vide."""
    try:
        sheet_id = _env("SPREADSHEET_ID")
        if not sheet_id:
            return jsonify(ok=False, error="SPREADSHEET_ID manquant"), 400
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

# -----------------------------------------------------------------------------
# Entrée
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(_env("PORT", "10000")))
