import os
from flask import Flask, render_template, make_response, jsonify

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")

app = Flask(__name__, template_folder=TEMPLATE_DIR)

@app.route("/healthz")
def health():
    return "ok", 200

@app.route("/_debug_files")
def debug_files():
    try:
        return jsonify({
            "cwd": os.getcwd(),
            "base_dir": BASE_DIR,
            "template_dir": TEMPLATE_DIR,
            "root_files": sorted(os.listdir(BASE_DIR)),
            "templates_files": sorted(os.listdir(TEMPLATE_DIR)) if os.path.isdir(TEMPLATE_DIR) else "templates dir not found"
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/", methods=["GET", "HEAD"])
def form():
    # HEAD de santé
    if os.environ.get("REQUEST_METHOD") == "HEAD":
        resp = make_response("", 200)
        resp.headers["Content-Length"] = "0"
        return resp

    try:
        return render_template("form.html")
    except Exception as e:
        html = f"""
        <h1>Walton Holidays</h1>
        <p>Template <code>{TEMPLATE_DIR}/form.html</code> introuvable ou erreur Jinja.</p>
        <pre>{type(e).__name__}: {e}</pre>
        """
        return html, 200

@app.route("/ping")
def ping():
    return "pong", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))

from flask import request

@app.route("/submit", methods=["POST"])
def submit():
    # on valide que la page fonctionne; l’envoi d’emails/Sheets viendra après
    name = request.form.get("name")
    email = request.form.get("email")
    return f"Demande reçue pour {name} ({email}).", 200
