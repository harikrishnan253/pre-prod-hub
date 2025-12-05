import os
import time
import zipfile
import shutil
import threading
from datetime import datetime
from pathlib import Path

from flask import (
    Blueprint,
    request,
    jsonify,
    send_file,
    current_app,
    render_template,
    session,
)
from config import ROUTE_PERMISSIONS, UPLOAD_FOLDER
from auth_utils import role_required
from jinja2 import Template
from utils import log_errors      # ✅ REQUIRED FIX
import chardet
import re

import sys
sys.path.append(os.path.dirname(os.path.dirname(__file__)))


ppd_bp = Blueprint("ppd", __name__)


# ---------------------------------------------------------
#   GLOBAL FUNCTION (importable)
# ---------------------------------------------------------
def html_to_excel_no_images(html_path, output_dir):
    """Convert HTML file to XLS by removing <img> tags."""
    try:
        with open(html_path, "rb") as f:
            raw_data = f.read()

        detected = chardet.detect(raw_data)
        encoding = detected.get("encoding") or "utf-8"
        html_content = raw_data.decode(encoding, errors="ignore")

        # Remove image tags
        html_no_images = re.sub(r"<img\b[^>]*>", "", html_content, flags=re.IGNORECASE)
        html_no_images = re.sub(
            r"url\(\s*data:[^)]+\)", "url()", html_no_images, flags=re.IGNORECASE
        )

        base = Path(html_path).stem
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_file = os.path.join(output_dir, f"{base}_{timestamp}.xls")

        with open(out_file, "w", encoding="utf-8") as f:
            f.write(html_no_images)

        return out_file

    except Exception as e:
        print("Excel conversion error:", e)
        return None


# ---------------------------------------------------------
#   MAIN PPD ROUTE
# ---------------------------------------------------------
@ppd_bp.route("/ppd", methods=["GET", "POST"])
@role_required(ROUTE_PERMISSIONS.get("ppd", ["ADMIN"]))
def ppd_route():
    if request.method == "GET":
        return render_template("ppd.html")

    uploaded = request.files.getlist("docfiles")
    if not uploaded:
        return jsonify({"error": "No files uploaded"}), 400

    import tempfile

    try:
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        tmpdir = tempfile.mkdtemp(prefix="s4c_ppd_", dir=UPLOAD_FOLDER)
    except Exception as e:
        current_app.logger.error(f"Failed to create temp directory: {e}")
        return jsonify({"error": f"Failed to create temp directory: {str(e)}"}), 500
    saved = []

    # Save uploaded Word files
    for f in uploaded:
        fn = f.filename
        if not fn.lower().endswith((".doc", ".docx")):
            continue
        path = os.path.join(tmpdir, fn)
        f.save(path)
        saved.append(path)

    if not saved:
        return jsonify({"error": "No valid .doc/.docx files uploaded"}), 400

    username = session.get("username", "Analyst")

    job_id = str(int(time.time() * 1000))
    current_app.config.setdefault("PROGRESS_DATA", {})
    current_app.config["PROGRESS_DATA"][job_id] = {
        "total": len(saved),
        "current": 0,
        "status": "Starting",
    }

    # ---------------------------------------------------------
    #   BACKGROUND PROCESSOR (THREAD)
    # ---------------------------------------------------------
    app = current_app._get_current_object()

    def process_job(app, username):
        with app.app_context():
            try:
                from word_analyzer import (
                    CitationAnalyzer,
                    extract_with_word,
                    extract_with_docx,
                    generate_formatting_html,
                    generate_multilingual_html,
                    build_comments_html,
                    build_export_highlight_html,
                    build_detailed_summary_table,
                    DASHBOARD_CSS,
                    DASHBOARD_JS,
                    HTML_WRAPPER,
                    HAS_WIN32COM,
                )
            except Exception as e:
                log_errors([f"Import Error in word_analyzer: {e}"])
                current_app.config["PROGRESS_DATA"][job_id]["status"] = "Failed"
                return

            results = []

            for i, path in enumerate(saved, 1):

                fname = os.path.basename(path)
                current_app.config["PROGRESS_DATA"][job_id].update(
                    {"current": i, "status": f"Processing {fname}"}
                )

                try:
                    # Extract doc data
                    if os.name == "nt" and HAS_WIN32COM:
                        paras, comments, imgs, foot, end = extract_with_word(path)
                    else:
                        paras, comments, imgs, foot, end = extract_with_docx(path)

                    CitationAnalyzer().remove_tags_keep_formatting_docx(path)
                    analyzer = CitationAnalyzer()

                    doc_data = [(t, p, c) for (t, p, c, _) in paras]
                    dtypes = analyzer.analyze_document_citations(doc_data)
                    table_count = len(dtypes.get("Table", {}).get("Caption", {}))

                    fmt_html = generate_formatting_html(path, used_word=False)
                    spec_html = generate_multilingual_html(path)
                    com_html = build_comments_html(comments)

                    summary_html = build_detailed_summary_table(
                        dtypes,
                        imgs,
                        table_count,
                        foot,
                        end,
                        fmt_html,
                        spec_html,
                        com_html,
                    )

                    msr_html = analyzer.build_citation_tables_html(dtypes, fname)
                    exp_html = build_export_highlight_html(paras)

                    wc = sum(len(t.split()) for (t, _, _, _) in paras)

                    # Render dashboard HTML
                    template = Template(HTML_WRAPPER)
                    html = template.render(
                        doc_name=fname,
                        pages=(len(paras) // 40) + 1,
                        words=wc,
                        ce_pages=(wc // 250) + 1,
                        date=datetime.now().strftime("%d-%m-%Y"),
                        analyst=username,
                        detailed_summary=summary_html,
                        msr_content=msr_html,
                        fmt_content=fmt_html,
                        spec_content=spec_html,
                        comment_content=com_html,
                        export_highlight=exp_html,
                        images=imgs,
                        footnotes=foot,
                        endnotes=end,
                        css=DASHBOARD_CSS,
                        js=DASHBOARD_JS,
                        logo_path="",
                    )

                    out_html = os.path.join(
                        tmpdir, Path(path).stem + "_Dashboard.html"
                    )
                    with open(out_html, "w", encoding="utf-8") as f:
                        f.write(html)

                    results.append(out_html)

                    # Excel Output
                    excel_output = html_to_excel_no_images(out_html, tmpdir)
                    if excel_output:
                        results.append(excel_output)
                    else:
                        log_errors(
                            [f"Excel conversion FAILED for file: {out_html}"]
                        )

                except Exception as e:
                    current_app.logger.error(f"Failed processing {fname}: {e}")
                    log_errors([f"Exception PPD processing {fname}: {e}"])
                    current_app.config["PROGRESS_DATA"][job_id][
                        "status"
                    ] = f"Failed: {e}"
                    break

            # ZIP results
            zip_path = os.path.join(tmpdir, "PPD_Results.zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                for f in results + saved:
                    z.write(f, arcname=os.path.basename(f))

            current_app.config["PROGRESS_DATA"][job_id].update(
                {"status": "Completed", "zip_path": zip_path, "current": len(saved)}
            )

    # Start thread
    threading.Thread(target=process_job, args=(app, username,), daemon=True).start()
    return jsonify({"job_id": job_id})


# ---------------------------------------------------------
#   STATUS ENDPOINTS
# ---------------------------------------------------------
@ppd_bp.route("/progress/<job_id>")
def progress(job_id):
    return jsonify(current_app.config.get("PROGRESS_DATA", {}).get(job_id, {}))


@ppd_bp.route("/download_zip/<job_id>")
def download_zip(job_id):
    data = current_app.config.get("PROGRESS_DATA", {}).get(job_id)
    if not data or "zip_path" not in data:
        return "Not ready", 404
    return send_file(
        data["zip_path"],
        as_attachment=True,
        download_name="PPD_Results.zip",
    )
