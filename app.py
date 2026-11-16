"""
app.py — TVS CPS Deck Generator Web App
Deploy: Render (render.yaml included) or run locally with `py app.py`
"""
import os, sys, uuid, threading, time, zipfile, shutil, tempfile
from flask import Flask, request, render_template, send_file, jsonify

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE_DIR, "tvs_cps_automation"))

import extract_data as ed
import run_pipeline as rp
import validate as val

# Use system temp dir on hosted servers, local dirs otherwise
UPLOAD_DIR  = os.path.join(tempfile.gettempdir(), "tvs_uploads")
OUTPUT_DIR  = os.path.join(tempfile.gettempdir(), "tvs_outputs")
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.pptx")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 150 * 1024 * 1024  # 150 MB

jobs: dict = {}

MONTH_NAMES = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]


def _detect_month(filename: str) -> int:
    """Detect report month index (1=Apr..12=Mar) from filename."""
    lower = filename.lower()
    abbrs = ["apr","may","jun","jul","aug","sep","oct","nov","dec","jan","feb","mar"]
    for i, abbr in enumerate(abbrs):
        if abbr in lower:
            return i + 1
    return 7  # default Oct


def _cleanup(*paths):
    for p in paths:
        if p and os.path.exists(p):
            try: os.remove(p)
            except: pass


def run_job(job_id, curr_path, prev_path, template_path, month_idx):
    unpack_dir = os.path.join(OUTPUT_DIR, f"{job_id}_unpacked")
    out_path   = os.path.join(OUTPUT_DIR, f"{job_id}_TVS_CPS_{MONTH_NAMES[month_idx-1]}.pptx")
    try:
        jobs[job_id]["msg"] = "Reading Excel files…"
        extracted = ed.extract_all(
            curr_xlsb=curr_path,
            prev_xlsb=prev_path,
            report_month_idx=month_idx,
            plan_leads_month=0,
            plan_retails_month=0,
            plan_leads_ytd=0,
            plan_retails_ytd=0,
        )
        jobs[job_id]["msg"] = "Copying template deck…"
        shutil.copy2(template_path, out_path)
        rp.unpack(out_path, unpack_dir)

        jobs[job_id]["msg"] = "Updating charts and slides…"
        validation_result = {}

        class Args:
            old_month_label = ""
            month_label = ""

        def run_val():
            time.sleep(1)
            r = val.validate_deck(unpack_dir, extracted)
            validation_result.update(r)

        vt = threading.Thread(target=run_val, daemon=True)
        vt.start()

        rp.update_charts(unpack_dir, extracted, Args())
        rp.update_slides(unpack_dir, extracted, Args())
        rp.repack(unpack_dir, out_path)
        vt.join(timeout=30)

        passed = validation_result.get("passed", "?")
        total  = validation_result.get("total",  "?")
        jobs[job_id].update({
            "status":     "done",
            "file":       out_path,
            "filename":   f"TVS_CPS_{MONTH_NAMES[month_idx-1]}.pptx",
            "msg":        f"Done — {passed}/{total} validation checks passed",
            "validation": validation_result,
        })
    except Exception as e:
        import traceback
        jobs[job_id].update({"status": "error", "msg": str(e), "trace": traceback.format_exc()})
    finally:
        _cleanup(curr_path, prev_path)
        if os.path.exists(unpack_dir):
            shutil.rmtree(unpack_dir, ignore_errors=True)
        # Clean up template if it was uploaded (not the stored one)
        if template_path != TEMPLATE_PATH:
            _cleanup(template_path)


@app.route("/")
def index():
    has_template = os.path.exists(TEMPLATE_PATH)
    return render_template("index.html", has_template=has_template)


@app.route("/generate", methods=["POST"])
def generate():
    def save(field):
        f = request.files.get(field)
        if not f or not f.filename:
            return None, None
        ext = os.path.splitext(f.filename)[1]
        path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}{ext}")
        f.save(path)
        return path, f.filename

    curr_path, curr_name = save("curr_xlsb")
    prev_path, _         = save("prev_xlsb")
    tmpl_path, _         = save("template_pptx")

    if not tmpl_path:
        tmpl_path = TEMPLATE_PATH if os.path.exists(TEMPLATE_PATH) else None

    errors = []
    if not curr_path: errors.append("Current FY Excel file is required.")
    if not prev_path: errors.append("Previous FY Excel file is required.")
    if not tmpl_path: errors.append("No template deck found on server. Please upload the base PPTX.")
    if errors:
        _cleanup(curr_path, prev_path)
        return jsonify({"error": " ".join(errors)}), 400

    # Save uploaded template as the stored one if requested
    if request.form.get("save_template") == "1" and tmpl_path != TEMPLATE_PATH:
        shutil.copy2(tmpl_path, TEMPLATE_PATH)

    month_idx = _detect_month(curr_name or "")
    try:
        month_idx = int(request.form.get("month_idx") or month_idx)
    except (ValueError, TypeError):
        pass

    job_id = str(uuid.uuid4())[:8]
    jobs[job_id] = {"status": "running", "msg": "Starting…"}

    threading.Thread(
        target=run_job,
        args=(job_id, curr_path, prev_path, tmpl_path, month_idx),
        daemon=True,
    ).start()

    return jsonify({"job_id": job_id, "month_idx": month_idx})


@app.route("/status/<job_id>")
def status(job_id):
    return jsonify(jobs.get(job_id, {"status": "unknown", "msg": "Job not found"}))


@app.route("/download/<job_id>")
def download(job_id):
    j = jobs.get(job_id, {})
    if j.get("status") != "done" or not j.get("file"):
        return "Not ready", 404
    return send_file(j["file"], as_attachment=True, download_name=j["filename"])


@app.route("/has_template")
def has_template():
    return jsonify({"has_template": os.path.exists(TEMPLATE_PATH)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)
