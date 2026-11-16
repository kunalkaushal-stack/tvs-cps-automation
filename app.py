"""
app.py — TVS CPS Deck Generator Web App
Run:  py app.py
Then open http://localhost:5000
"""
import os, sys, uuid, threading, time, zipfile, shutil
from flask import Flask, request, render_template, send_file, jsonify, url_for

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "tvs_cps_automation"))
import extract_data as ed
import chart_updater as cu
import slide_updater as su
import validate as val
import run_pipeline as rp

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "outputs")
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.pptx")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB

# Job state: { job_id: {"status": "running"|"done"|"error", "msg": str, "file": path} }
jobs = {}


def run_job(job_id, curr_path, prev_path, template_path, month_idx, month_label,
            old_month_label, plan_lm, plan_rm, plan_ly, plan_ry):
    try:
        jobs[job_id]["msg"] = "Extracting data from Excel files…"
        extracted = ed.extract_all(
            curr_xlsb=curr_path,
            prev_xlsb=prev_path,
            report_month_idx=month_idx,
            plan_leads_month=plan_lm,
            plan_retails_month=plan_rm,
            plan_leads_ytd=plan_ly,
            plan_retails_ytd=plan_ry,
        )

        out_name = f"TVS_CPS_{month_label.replace(' ', '_')}.pptx"
        out_path = os.path.join(OUTPUT_DIR, f"{job_id}_{out_name}")
        unpack_dir = out_path + "_unpacked"

        jobs[job_id]["msg"] = "Copying template deck…"
        shutil.copy2(template_path, out_path)
        rp.unpack(out_path, unpack_dir)

        jobs[job_id]["msg"] = "Updating charts and slides…"
        validation_result = {}

        class FakeArgs:
            old_month_label = ""
            month_label = ""

        fa = FakeArgs()
        fa.old_month_label = old_month_label
        fa.month_label = month_label

        def run_val():
            time.sleep(1)
            r = val.validate_deck(unpack_dir, extracted)
            validation_result.update(r)

        vt = threading.Thread(target=run_val, daemon=True)
        vt.start()

        rp.update_charts(unpack_dir, extracted, fa)
        rp.update_slides(unpack_dir, extracted, fa)
        rp.repack(unpack_dir, out_path)
        vt.join(timeout=30)

        passed = validation_result.get("passed", "?")
        total = validation_result.get("total", "?")
        jobs[job_id].update({
            "status": "done",
            "file":   out_path,
            "filename": out_name,
            "msg": f"Done! Validation: {passed}/{total} checks passed.",
            "validation": validation_result,
        })
    except Exception as e:
        import traceback
        jobs[job_id].update({"status": "error", "msg": str(e), "trace": traceback.format_exc()})
    finally:
        # Clean up uploads
        for p in [curr_path, prev_path]:
            if p and os.path.exists(p) and p.startswith(UPLOAD_DIR):
                try: os.remove(p)
                except: pass
        if template_path != TEMPLATE_PATH and os.path.exists(template_path):
            try: os.remove(template_path)
            except: pass


@app.route("/")
def index():
    has_template = os.path.exists(TEMPLATE_PATH)
    return render_template("index.html", has_template=has_template)


@app.route("/generate", methods=["POST"])
def generate():
    def save(field):
        f = request.files.get(field)
        if not f or f.filename == "":
            return None
        ext = os.path.splitext(f.filename)[1]
        path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}{ext}")
        f.save(path)
        return path

    curr_path = save("curr_xlsb")
    prev_path = save("prev_xlsb")
    tmpl_path = save("template_pptx") or (TEMPLATE_PATH if os.path.exists(TEMPLATE_PATH) else None)

    errors = []
    if not curr_path: errors.append("Current FY Excel file is required.")
    if not prev_path: errors.append("Previous FY Excel file is required.")
    if not tmpl_path: errors.append("Template PPTX is required (upload once or save as template.pptx).")
    if errors:
        return jsonify({"error": " ".join(errors)}), 400

    try:
        month_idx      = int(request.form.get("month_idx", 1))
        month_label    = request.form.get("month_label", "").strip() or "New Month"
        old_month_label= request.form.get("old_month_label", "").strip()
        plan_lm        = int(request.form.get("plan_leads_month", 0) or 0)
        plan_rm        = int(request.form.get("plan_retails_month", 0) or 0)
        plan_ly        = int(request.form.get("plan_leads_ytd", 0) or 0)
        plan_ry        = int(request.form.get("plan_retails_ytd", 0) or 0)
    except ValueError as e:
        return jsonify({"error": f"Invalid number: {e}"}), 400

    # Save uploaded template as the stored template if checkbox ticked
    if request.form.get("save_template") == "1" and tmpl_path != TEMPLATE_PATH:
        shutil.copy2(tmpl_path, TEMPLATE_PATH)

    job_id = str(uuid.uuid4())[:8]
    jobs[job_id] = {"status": "running", "msg": "Starting…"}

    t = threading.Thread(
        target=run_job,
        args=(job_id, curr_path, prev_path, tmpl_path,
              month_idx, month_label, old_month_label,
              plan_lm, plan_rm, plan_ly, plan_ry),
        daemon=True,
    )
    t.start()
    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def status(job_id):
    j = jobs.get(job_id, {"status": "unknown", "msg": "Job not found"})
    return jsonify(j)


@app.route("/download/<job_id>")
def download(job_id):
    j = jobs.get(job_id, {})
    if j.get("status") != "done" or not j.get("file"):
        return "Not ready", 404
    return send_file(j["file"], as_attachment=True, download_name=j["filename"])


if __name__ == "__main__":
    app.run(debug=False, port=5000)
