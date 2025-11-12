import os
from flask import Flask, render_template, Response, request, jsonify

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")

app = Flask(__name__, template_folder=TEMPLATE_DIR)

@app.route("/healthz")
def health():
    return "ok", 200

@app.route("/_debug_files")
def debug_files():
    return jsonify({
        "cwd": os.getcwd(),
        "base_dir": BASE_DIR,
        "template_dir": TEMPLATE_DIR,
        "root_files": sorted(os.listdir(BASE_DIR)),
        "templates_files": sorted(os.listdir(TEMPLATE_DIR)) if os.path.isdir(TEMPLATE_DIR) else "templates dir not found"
    })

@app.route("/", methods=["GET", "HEAD"])
def form():
    if request.method == "HEAD":
        return Response("", status=200, mimetype="text/html")
    html = render_template("form.html")
    return Response(html, status=200, mimetype="text/html; charset=utf-8")

from flask import redirect
@app.route("/submit", methods=["POST"])
def submit():
    # placeholder pour éviter 500 pendant les tests
    return "Demande reçue.", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
