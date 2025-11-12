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
# Config approbateurs (via ENV)
# -------------------------------
APPROVERS = {
    "mark": os.environ.get("EMAIL_MARK", "mw@walton.fr"),
    "nhan": os.environ.get("EMAIL_NHAN", "nhan@walton.fr"),
    "anh":  os.environ.get("EMAIL_ANH",  "anh@walton.fr"),
}


# -------------------------------
# Helpers
# -------------------------------
def _env(key, default=None):
    return os.environ.get(key, default)

def business_days(d1: date, d2: date) -> int:
    """Jours ouvrés inclusifs (lun–ven), sans jours fériés."""
    if d2 < d1:
        d1, d2 = d2, d1
    days = 0
    cur = d1
    while cur <= d2:
        if cur.weekday() < 5:
            days += 1
        cur += timedelta(days=1)
    return days

def send_mail(to_addrs, subject, html, cc_addrs=None) -> bool:
    """Envoi SMTP (Gmail App Password)."""
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
    """Crée un client gspread à partir de GOOGLE_SERVICE_ACCOUNT_JSON ou GOOGLE_APPLICATION_CREDENTIALS."""
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
    """Toujours AJOUTER une nouvelle ligne (ne jamais écraser)."""
    sheet_id = _env("SPREADSHEET_ID")
    if not sheet_id:
        app.logger.error("❌ SPREADSHEET_ID manquant")
        return False

    gc = get_sheets_client()
    sh = gc.open_by_key(sheet_id)
    ws = sh.sheet1

    # Entêtes si vide
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
        "APPROUVÉ" if status == "approved" else "REFUSÉ",
    ]

    ws.append_row(row_data)
    app.logger.info(f"✅ Sheet appended new row for {email}")
    return True


# -------------------------------
# Routes publiques
# -------------------------------
@app.route("/", methods=["GET"])
def form():
    html = render_template("form.html")
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
        html2 = render_template("submitted.html", name=name, approver=approver_email, start=start_date, end=end_date, days=days)
        return Response(html2, status=200, mimetype="text/html; charset=utf-8")
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

    # Recalcule les jours ouvrés
    try:
        d1 = datetime.strptime(sd, "%Y-%m-%d").date()
        d2 = datetime.strptime(ed, "%Y-%m-%d").date()
        days = business_days(d1, d2)
    except Exception:
        days = 0

    # Ecriture Google Sheet uniquement si approuvé
    if status == "approved":
        approver_email = _env("ALWAYS_CC", "mw@walton.fr")  # à défaut
        try:
            write_decision_to_sheet(
                email=email, name=name, start_date=sd, end_date=ed,
                reason=reason, days=days, approver_email=approver_email, status=status
            )
        except Exception as e:
            app.logger.exception(f"❌ Sheet write failed: {e}")

    # Mail de décision
    note = f"approuvée ✅ ({days} jours ouvrés)" if status == "approved" else "refusée ❌"
    send_mail(
        [email],
        f"[Walton] Décision — {status}",
        f"<p>Bonjour {name},</p><p>Votre demande de congés {sd} → {ed} est {note}.</p>",
        cc_addrs=[_env("ALWAYS_CC", "mw@walton.fr")]
    )

    html = render_template("decision_result.html", name=name, email=email, status=status, start=sd, end=ed, days=days)
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
    """Affiche les demandes depuis Google Sheets (triées par date desc)."""
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
        <!doctype html><html lang='fr'><head>
          <meta charset='utf-8'>
          <meta name='viewport' content='width=device-width, initial-scale=1'>
          <title>Admin - Walton Holidays</title>
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
            <h1>Admin — Demandes de congés</h1>
            <p>Tri : du plus récent au plus ancien. Total : {len(data)} demandes.</p>
            <table>
              <thead>
                <tr>
                  <th>Timestamp</th><th>Nom</th><th>Email</th><th>Début</th><th>Fin</th><th>Raison</th><th>Jours</th><th>Approbateur</th><th>Statut</th>
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
        return jsonify(ok=False, error=str(e)), 500


# -------------------------------
# Diagnostics
# -------------------------------
@app.get("/_smtp_test")
def smtp_test():
    """Test envoi SMTP."""
    to = _env("SMTP_USER")
    ok = send_mail([to], "[Walton] SMTP Test", "<p>Test d'envoi réussi ✅</p>")
    if ok:
        return jsonify(ok=True, to=to)
    return jsonify(ok=False, error="SMTP send failed (voir logs Render)"), 500

@app.get("/_sheet_test")
def sheet_test():
    """Test accès Google Sheet + init entête si vide."""
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


# -------------------------------
# Run
# -------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(_env("PORT", "10000")))
