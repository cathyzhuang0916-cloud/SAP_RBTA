"""
Flask server for the APAC KPI Dashboard.
Run:  python app.py
Then: browser opens automatically at http://localhost:5050
"""

import io
import threading
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, send_file, request, make_response

import kpi_engine

BASE_DIR  = Path(__file__).parent
HTML_FILE = BASE_DIR / "dashboard.html"

# Auto-detect template: prefer _UPDATED, fall back to _LATEST
def _find_template():
    for name in [
        "APAC – Role Transformation KPIs_UPDATED.pptx",
        "APAC – Role Transformation KPIs_LATEST.pptx",
    ]:
        p = BASE_DIR / name
        if p.exists():
            return p
    raise FileNotFoundError("No PPTX template found. Expected APAC – Role Transformation KPIs_UPDATED.pptx")

app = Flask(__name__, static_folder=str(BASE_DIR), static_url_path="")

_cache = {}

def _load_default():
    try:
        data = kpi_engine.compute_kpis(str(BASE_DIR / "Impact KPI.xlsx"))
        _cache["data"] = kpi_engine.kpi_data_to_json(data)
        _cache["raw"]  = data
        print(f"  Loaded KPI data: timestamp={_cache['data']['timestamp']}, prior_ts={_cache['data'].get('prior_ts','?')}")
    except Exception as e:
        print(f"  Warning: could not pre-load KPI data: {e}")
        _cache["data"] = {}
        _cache["raw"]  = {}


@app.route("/")
def index():
    return HTML_FILE.read_text(encoding="utf-8")


@app.route("/data")
def data():
    return jsonify(_cache.get("data", {}))


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"error": "Please upload an .xlsx file"}), 400

    tmp_path = BASE_DIR / "_uploaded_impact_kpi.xlsx"
    f.save(str(tmp_path))

    try:
        raw = kpi_engine.compute_kpis(str(tmp_path))
        _cache["data"] = kpi_engine.kpi_data_to_json(raw)
        _cache["raw"]  = raw
        return jsonify({
            "status":    "ok",
            "timestamp": _cache["data"]["timestamp"],
            "prior_ts":  _cache["data"].get("prior_ts", ""),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


@app.route("/export", methods=["POST"])
def export_ppt():
    if not _cache.get("raw"):
        return jsonify({"error": "No KPI data loaded. Please upload a file first."}), 400
    try:
        template = _find_template()
        out_buf  = io.BytesIO()
        kpi_engine.build_pptx(_cache["raw"], str(template), out_buf)
        size = out_buf.tell()
        out_buf.seek(0)

        ts_safe  = _cache["data"]["timestamp"].replace("/", "-")
        filename = f"APAC_Role_Transformation_KPIs_{ts_safe}.pptx"

        resp = make_response(out_buf.read())
        resp.headers["Content-Type"]           = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        resp.headers["Content-Disposition"]    = f'attachment; filename="{filename}"'
        resp.headers["Content-Length"]         = str(size)
        resp.headers["X-Content-Type-Options"] = "nosniff"
        return resp
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("Loading initial KPI data...")
    _load_default()
    try:
        tmpl = _find_template()
        print(f"  Using template: {tmpl.name}")
    except FileNotFoundError as e:
        print(f"  Warning: {e}")
    print("Starting dashboard at http://localhost:5050")
    threading.Timer(1.2, lambda: webbrowser.open("http://localhost:5050")).start()
    app.run(host="0.0.0.0", port=5050, debug=False)
