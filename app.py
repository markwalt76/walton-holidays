import os
from flask import Flask, render_template, make_response

app = Flask(__name__, template_folder="templates")

@app.route("/healthz")
def health():
    return "ok", 200

@app.route("/", methods=["GET","HEAD"])
def form():
    # HEAD: ne renvoie que les headers pour les checks Render
    if os.environ.get("REQUEST_METHOD") == "HEAD":
        resp = make_response("", 200)
        resp.headers["Content-Length"] = "0"
        return resp

    # Essaie d'afficher le template; sinon fallback minimal
    try:
        return render_template("form.html")
    except Exception as e:
        html = f"""
        <h1>Walton Holidays</h1>
        <p>Template <code>templates/form.html</code> introuvable ou erreur Jinja.</p>
        <pre>{type(e).__name__}: {e}</pre>
        """
        return html, 200

# Optionnel: page simple pour tester vite
@app.route("/ping")
def ping():
    return "pong", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
