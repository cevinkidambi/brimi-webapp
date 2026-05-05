"""
app.py — Flask web app for BRIMI Fund Performance Automation.

Upload 3 files → process → download output.
"""

import os
import tempfile
import threading
import uuid
from datetime import datetime

from flask import Flask, render_template, request, jsonify, send_file, abort

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB max upload

# In-memory store for processing state per request
processing_state = {}

TMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tmp")
os.makedirs(TMP_DIR, exist_ok=True)


@app.route("/")
def index():
    today = datetime.now().strftime("%d %B %Y")
    return render_template("index.html", today=today)


@app.route("/process", methods=["POST"])
def process():
    request_id = str(uuid.uuid4())[:8]
    logs = []

    def log_callback(msg):
        logs.append(msg)

    try:
        # Validate uploads
        files = request.files
        missing = []
        for key, label in [
            ("historicalnav_t1", "HistoricalNAV (T-1)"),
            ("historicalnav_t2", "HistoricalNAV (T-2)"),
            ("bloomberg", "INDEKS Bloomberg"),
        ]:
            if key not in files or files[key].filename == "":
                missing.append(label)

        if missing:
            return jsonify({
                "status": "error",
                "message": f"Missing files: {', '.join(missing)}"
            }), 400

        # Save uploads to temp files
        saved_paths = {}
        for key in ["historicalnav_t1", "historicalnav_t2", "bloomberg"]:
            f = files[key]
            ext = os.path.splitext(f.filename)[1] or ".xlsx"
            tmp_path = os.path.join(TMP_DIR, f"upload_{request_id}_{key}{ext}")
            f.save(tmp_path)
            saved_paths[key] = tmp_path

        # Run pipeline
        from brimi_engine import run_pipeline
        output_path, nav_date = run_pipeline(
            saved_paths["historicalnav_t1"],
            saved_paths["historicalnav_t2"],
            saved_paths["bloomberg"],
            log_callback=log_callback,
        )

        # Rename output with date for clean download
        download_name = f"BRIMI_Output_{nav_date.replace(' ', '_')}.xlsx"
        download_path = os.path.join(TMP_DIR, f"output_{request_id}_{download_name}")
        import shutil
        shutil.copy2(output_path, download_path)

        # Clean up uploaded files
        for p in saved_paths.values():
            try:
                os.remove(p)
            except OSError:
                pass

        return jsonify({
            "status": "ok",
            "download_url": f"/download?file={os.path.basename(download_path)}",
            "logs": logs,
        })

    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        logs.append(f"\nERROR: {e}")
        logs.append(tb)
        return jsonify({
            "status": "error",
            "message": str(e),
            "logs": logs,
        }), 500


@app.route("/download")
def download():
    filename = request.args.get("file")
    if not filename:
        abort(400)
    filepath = os.path.join(TMP_DIR, filename)
    if not os.path.exists(filepath):
        abort(404)
    return send_file(
        filepath,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)
