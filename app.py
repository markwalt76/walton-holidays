import os
import smtplib
import ssl
from email.message import EmailMessage
from flask import Flask, render_template, request, jsonify, Response

app = Flask(__name__, template_folder="templates")

# --- CONFIG APPROVERS ---
APPROVERS = {
    "mark": os.environ.get("EMAIL_MARK", "mw@walton.fr"),
    "nhan": os.environ.get("EMAIL_NHAN", "nhan@walton.fr"),
    "anh": os.environ.get("EMAIL_ANH", "anh@walton.fr"),
}

# --- HELPER : ENVOI EMAIL ---
def send_mail(to_addrs, subject, html, cc_addrs=None):
    host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    port = int(os.environ.get("SMTP_PORT", "587"))
    user = os.environ.get("SMTP_USER")
    pwd = os.environ.get("SMTP_PASSWORD")
    mail_from = os.environ.get("MAIL_FROM", user)

    if not user or not pwd:
        app.logger.error("SMTP credentials missing")
        return False

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = mail_from or user
    if isinstance(to_addrs, str):
        to_addrs = [to_addrs]
    msg["To"] = ", ".join(to_addrs)
    if cc_addrs:
        if isinstance(cc_addrs, str):
            cc_addrs = [cc_addrs]
        msg["Cc"] = ", ".join(cc_addrs)
    msg.set_content("HTML version required")
    msg.add_alternative(html, subtype="html")

    try:
        with smtplib.SMTP(host, port, timeout=30) as s:
            s.ehlo()
            s.starttls(context=ssl.create_default_context())
            s.login(user, pwd)
            s.send_message(msg)
        app.logger.info(f"Mail sent to {to_addrs} (cc={cc_addrs})")
        return True
    except Exception as e:
        app.logger.exception(f"SMTP send failed: {e}")
        return False

# --- ROUTES ---
@app.route("/")
def form():
    return render_template("form.html")

@app.route("/submit", methods=["POST"])
def submit():
    name = request.form.get("name", "").strip()
    email = request.form.get("email", "").strip()
    approver_k = request.form.get("approver")
    start_date = request.form.get("start_date")
    end_date = request.form.get("end_date")
    reason = request.form.get("reason", "").strip()

    approver_email = APPROVERS.get(approver_k)
    if not approver_email:
        return "Approver inconnu", 400

    base_url = os.environ.get("BASE_URL", request.host_url.rstrip("/"))
    approve_link = f"{base_url}/decision?status=approved&email={email}&name={name}&sd={start_date}&ed={end_date}"
    reject_link = f"{base_url}/decision?status=rejected&email={email}&name={name}&sd={start_date}&ed={end_date}"

    html = f"""
      <h2>Nouvelle demande de congés</h2>
      <p><b>Demandeur :</b> {name} &lt;{email}&gt;</p>
      <p><b>Période :</b> {start_date} → {end_date}</p>
      <p><b>Raison :</b><br>{(reason or '—').replace('\n','<br>')}</p>
      <p>
        <a href="{approve_link}" style="display:inline-block;padding:10px 14px;background:#16a34a;color:#fff;border-radius:8px;text-decoration:none;">✅ Approuver</a>
        &nbsp;&nbsp;
        <a href="{reject_link}" style="display:inline-block;padding:10px 14px;background:#dc2626;color:#fff;border-radius:8px;text-decoration:none;">❌ Refuser</a>
      </p>
    """

    cc = [os.environ.get("ALWAYS_CC", "mw@walton.fr")]
    ok = send_mail(
        [approver_email],
        f"[Walton] Demande de congés — {name}",
        html,
        cc_addrs=cc
    )

    if ok:
        ack = f"""
          <p>Bonjour {name},</p>
          <p>Votre demande ({start_date} → {end_date}) a été envoyée à {approver_email}.</p>
          <p>Vous recevrez un email dès décision.</p>
        """
        send_mail([email], "[Walton] Votre demande a été envoyée", ack, cc_addrs=cc)
        return "Demande envoyée ✅", 200
    else:
        return "Erreur d'envoi email ❌ — vérifiez les logs.", 500

@app.route("/decision")
def decision():
    status = request.args.get("status")
    email = request.args.get("email")
    name = request.args.get("name", "")
    if status == "approved":
        note = f"La demande de {name} ({email}) est approuvée ✅"
    else:
        note = f"La demande de {name} ({email}) est refusée ❌"
    send_mail(
        [email],
        f"[Walton] Décision — {status}",
        f"<p>{note}</p>",
        cc_addrs=[os.environ.get("ALWAYS_CC", "mw@walton.fr")]
    )
    return f"{note}", 200

# --- ROUTE DE TEST SMTP ---
@app.get("/_smtp_test")
def smtp_test():
    """Test l'envoi SMTP avec les variables d'environnement actuelles."""
    to = os.getenv("SMTP_USER")
    try:
        send_mail(to, "[Walton] SMTP Test", "<p>Test d'envoi réussi ✅</p>")
        return jsonify(ok=True, to=to)
    except Exception as e:
        app.logger.exception(f"SMTP test failed: {e}")
        return jsonify(ok=False, error=str(e)), 500

@app.route("/ping")
def ping():
    return "pong", 200

@app.route("/healthz")
def health():
    return "ok", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
