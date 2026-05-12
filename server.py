from flask import Flask, request, send_file, jsonify, render_template
import io
import os
from werkzeug.utils import secure_filename

from ppt_generator import generate_ppt

app = Flask(__name__)

app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024


# ======================
# HOME PAGE
# ======================
@app.route("/")
def index():
    return render_template("index.html")


# ======================
# GENERATE PPT
# ======================
@app.route("/generate", methods=["POST"])
def generate():

    if "file" not in request.files:
        return jsonify(error="No file uploaded"), 400

    file = request.files["file"]
    filename = secure_filename(file.filename)

    if not filename.endswith(".xlsx"):
        return jsonify(error="Upload .xlsx file only"), 400

    try:
        ppt_bytes = generate_ppt(file.read())

        output_name = filename.replace(".xlsx", "_report.pptx")

        return send_file(
            io.BytesIO(ppt_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=output_name
        )

    except Exception as e:
        import traceback
        return jsonify(
            error=str(e),
            detail=traceback.format_exc()
        ), 500


# ======================
# LOCAL RUN
# ======================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
