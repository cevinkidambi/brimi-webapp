"""
api/index.py — Vercel entrypoint for BRIMI Flask app.
All HTML inlined — no templates folder (Vercel only deploys api/ files).
"""
import os
import sys
import uuid
from datetime import datetime
from flask import Flask, request, jsonify, send_file, abort

# Add project root to sys.path for pipeline modules
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

# Use /tmp for uploads — /var/task is read-only on Vercel
TMP_DIR = "/tmp"

# ─── Inlined HTML (replaces templates/index.html) ──────────────────────────
def serve_page(today):
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BRIMI BRIMI! HAHAHA!</title>
    <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&family=Roboto:wght@400;500&display=swap" rel="stylesheet">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <style>
        :root {{
            --md-sys-color-primary: #FFB4C2;
            --md-sys-color-primary-dark: #FF8FA3;
            --md-sys-color-on-primary: #000000;
            --md-sys-color-surface: #FFFFFF;
            --md-sys-color-surface-variant: #F5F5F5;
            --md-sys-color-outline: #E0E0E0;
            --md-sys-color-on-surface: #1C1B1F;
            --md-sys-color-success: #4CAF50;
            --md-sys-color-error: #F44336;
            --md-ref-typeface-brand: 'Google Sans', sans-serif;
            --md-ref-typeface-plain: 'Roboto', sans-serif;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: var(--md-ref-typeface-plain);
            background: var(--md-sys-color-surface);
            color: var(--md-sys-color-on-surface);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
        }}
        .container {{
            max-width: 640px;
            width: 100%;
            padding: 24px 16px;
        }}
        header {{ text-align: center; margin-bottom: 32px; }}
        header h1 {{
            font-family: var(--md-ref-typeface-brand);
            font-size: 28px;
            font-weight: 700;
            margin-bottom: 4px;
        }}
        header .subtitle {{ font-size: 14px; color: #666; }}
        header .date {{ font-size: 13px; color: #888; margin-top: 8px; font-weight: 500; }}
        .card {{
            background: var(--md-sys-color-surface);
            border: 1px solid var(--md-sys-color-outline);
            border-radius: 12px;
            padding: 24px;
            margin-bottom: 24px;
        }}
        .upload-field {{ margin-bottom: 20px; }}
        .upload-field label {{ display: block; font-weight: 500; font-size: 14px; margin-bottom: 8px; }}
        .upload-field .hint {{ font-size: 12px; color: #888; margin-bottom: 6px; }}
        .upload-field input[type="file"] {{
            width: 100%;
            padding: 12px;
            border: 2px dashed var(--md-sys-color-outline);
            border-radius: 8px;
            background: var(--md-sys-color-surface-variant);
            cursor: pointer;
            font-size: 14px;
        }}
        .upload-field input[type="file"]:hover {{ border-color: var(--md-sys-color-primary); }}
        .process-btn {{
            width: 100%;
            padding: 14px 24px;
            background: var(--md-sys-color-primary);
            color: var(--md-sys-color-on-primary);
            border: none;
            border-radius: 8px;
            font-family: var(--md-ref-typeface-brand);
            font-size: 16px;
            font-weight: 500;
            cursor: pointer;
        }}
        .process-btn:hover {{ background: var(--md-sys-color-primary-dark); }}
        .process-btn:disabled {{ background: #E0E0E0; color: #999; cursor: not-allowed; }}
        .loading {{ display: none; text-align: center; padding: 24px; }}
        .loading.active {{ display: block; }}
        .spinner {{
            display: inline-block;
            width: 40px;
            height: 40px;
            border: 4px solid var(--md-sys-color-outline);
            border-top-color: var(--md-sys-color-primary);
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            margin-bottom: 12px;
        }}
        @keyframes spin {{ to {{ transform: rotate(360deg); }} }}
        .loading p {{ font-size: 14px; color: #666; }}
        .status {{ display: none; margin-top: 24px; }}
        .status.active {{ display: block; }}
        .status.success {{ padding: 16px; background: #E8F5E9; border: 1px solid #4CAF50; border-radius: 8px; }}
        .status.error {{ padding: 16px; background: #FFEBEE; border: 1px solid #F44336; border-radius: 8px; }}
        .status .title {{ font-weight: 500; font-size: 15px; margin-bottom: 8px; }}
        .status.success .title {{ color: #2E7D32; }}
        .status.error .title {{ color: #C62828; }}
        .download-btn {{
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 12px 24px;
            background: #4CAF50;
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            text-decoration: none;
        }}
        .download-btn:hover {{ background: #43A047; }}
        details {{ margin-top: 12px; border: 1px solid var(--md-sys-color-outline); border-radius: 8px; }}
        details summary {{ padding: 12px 16px; cursor: pointer; font-size: 13px; font-weight: 500; color: #666; }}
        .log-content {{
            padding: 12px 16px;
            font-family: monospace;
            font-size: 12px;
            line-height: 1.6;
            max-height: 300px;
            overflow-y: auto;
            white-space: pre-wrap;
            word-break: break-all;
        }}
        footer {{ text-align: center; padding: 24px 16px; font-size: 12px; color: #999; margin-top: auto; }}
        @media (max-width: 480px) {{
            .container {{ padding: 16px 12px; }}
            header h1 {{ font-size: 24px; }}
            .card {{ padding: 16px; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>BRIMI BRIMI! HAHAHA!</h1>
            <p class="subtitle">Fund Performance Automation</p>
            <p class="date">{today}</p>
        </header>
        <div class="card">
            <form id="uploadForm" enctype="multipart/form-data">
                <div class="upload-field">
                    <label for="historicalnav_t1">HistoricalNAV (T-1)</label>
                    <p class="hint">BRI daily AUM — most recent date</p>
                    <input type="file" id="historicalnav_t1" name="historicalnav_t1" accept=".xlsx" required>
                </div>
                <div class="upload-field">
                    <label for="historicalnav_t2">HistoricalNAV (T-2)</label>
                    <p class="hint">BRI daily AUM — previous date</p>
                    <input type="file" id="historicalnav_t2" name="historicalnav_t2" accept=".xlsx" required>
                </div>
                <div class="upload-field">
                    <label for="bloomberg">INDEKS Bloomberg</label>
                    <p class="hint">Bloomberg global indexes</p>
                    <input type="file" id="bloomberg" name="bloomberg" accept=".xlsx" required>
                </div>
                <button type="submit" class="process-btn" id="processBtn">Process Files</button>
            </form>
            <div class="loading" id="loading">
                <div class="spinner"></div>
                <p>Processing files... (this may take 5–15 seconds)</p>
            </div>
            <div class="status" id="status">
                <div class="title" id="statusTitle"></div>
                <div id="statusBody"></div>
            </div>
        </div>
    </div>
    <footer>&copy; 2026 Saya akan LAWANNN!. All rights reserved.</footer>
    <script>
        const form = document.getElementById('uploadForm');
        const loading = document.getElementById('loading');
        const status = document.getElementById('status');
        const statusTitle = document.getElementById('statusTitle');
        const statusBody = document.getElementById('statusBody');
        const processBtn = document.getElementById('processBtn');

        form.addEventListener('submit', async (e) => {{
            e.preventDefault();
            const t1 = document.getElementById('historicalnav_t1').files[0];
            const t2 = document.getElementById('historicalnav_t2').files[0];
            const bb = document.getElementById('bloomberg').files[0];
            if (!t1 || !t2 || !bb) {{ alert('Please select all 3 files.'); return; }}

            loading.classList.add('active');
            status.classList.remove('active', 'success', 'error');
            processBtn.disabled = true;
            processBtn.textContent = 'Processing...';

            const formData = new FormData();
            formData.append('historicalnav_t1', t1);
            formData.append('historicalnav_t2', t2);
            formData.append('bloomberg', bb);

            try {{
                const response = await fetch('/process', {{ method: 'POST', body: formData }});
                const data = await response.json();
                handleResponse(data);
            }} catch (err) {{
                loading.classList.remove('active');
                status.classList.add('active', 'error');
                statusTitle.textContent = 'Network error';
                statusBody.textContent = err.message;
            }}
            processBtn.disabled = false;
            processBtn.textContent = 'Process Files';
        }});

        function handleResponse(data) {{
            loading.classList.remove('active');
            status.classList.add('active');
            processBtn.disabled = false;
            processBtn.textContent = 'Process Files';
            if (data.status === 'ok') {{
                status.classList.add('success');
                statusTitle.textContent = 'Success!';
                statusBody.innerHTML = '<a href="' + data.download_url + '" class="download-btn"><span class="material-icons">download</span> Download Output</a><details><summary>Processing Log</summary><div class="log-content">' + escapeHtml(data.logs.join('\\\\n')) + '</div></details>';
            }} else {{
                status.classList.add('error');
                statusTitle.textContent = 'Error: ' + escapeHtml(data.message);
                statusBody.innerHTML = '<details><summary>Full Log</summary><div class="log-content">' + escapeHtml(data.logs.join('\\\\n')) + '</div></details>';
            }}
        }}
        function escapeHtml(s) {{ const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }}
    </script>
</body>
</html>"""


@app.route("/")
def index():
    today = datetime.now().strftime("%d %B %Y")
    return serve_page(today)


@app.route("/process", methods=["POST"])
def process():
    request_id = str(uuid.uuid4())[:8]
    logs = []

    def log_callback(msg):
        logs.append(msg)

    try:
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

        saved_paths = {}
        for key in ["historicalnav_t1", "historicalnav_t2", "bloomberg"]:
            f = files[key]
            ext = os.path.splitext(f.filename)[1] or ".xlsx"
            tmp_path = os.path.join(TMP_DIR, f"upload_{request_id}_{key}{ext}")
            f.save(tmp_path)
            saved_paths[key] = tmp_path

        from brimi_engine import run_pipeline
        output_path, nav_date = run_pipeline(
            saved_paths["historicalnav_t1"],
            saved_paths["historicalnav_t2"],
            saved_paths["bloomberg"],
            log_callback=log_callback,
        )

        download_name = f"BRIMI_Output_{nav_date.replace(' ', '_')}.xlsx"
        download_path = os.path.join(TMP_DIR, f"output_{request_id}_{download_name}")
        import shutil
        shutil.copy2(output_path, download_path)

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
