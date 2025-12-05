import os
import shutil
from flask import Blueprint, render_template, redirect, url_for, session, request, send_from_directory, flash, jsonify, g, current_app
from database import get_db, init_db
from utils import log_activity, log_errors
from config import ROUTE_MACROS, REPORT_FOLDER, UPLOAD_FOLDER, DATABASE
from shared_state import download_tokens
from auth_utils import admin_required
from werkzeug.utils import secure_filename

main_bp = Blueprint('main', __name__)

@main_bp.route('/')
def index():
    if 'user_id' in session:
        return redirect(url_for('main.dashboard'))
    else:
        return redirect(url_for('auth.login'))

@main_bp.route("/dashboard", strict_slashes=False)
def dashboard():
    if 'user_id' not in session:
        return redirect(url_for('auth.login'))

    with get_db() as db:
        if session.get('is_admin'):
            recent_files = db.execute('''SELECT f.*, u.username 
                                       FROM files f 
                                       JOIN users u ON f.user_id = u.id 
                                       ORDER BY f.upload_date DESC LIMIT 5''').fetchall()

            recent_macro = db.execute('''SELECT m.*, u.username 
                                       FROM macro_processing m
                                       JOIN users u ON m.user_id = u.id 
                                       ORDER BY m.processing_date DESC LIMIT 5''').fetchall()

            # Route-specific stats
            route_stats = {}
            for route_type in ROUTE_MACROS.keys():
                count = db.execute("SELECT COUNT(*) FROM macro_processing WHERE route_type = ?",
                                   (route_type,)).fetchone()[0]
                route_stats[route_type] = count

            admin_stats = {
                'total_users': db.execute("SELECT COUNT(*) FROM users").fetchone()[0],
                'total_files': db.execute("SELECT COUNT(*) FROM files").fetchone()[0],
                'total_validations': db.execute("SELECT COUNT(*) FROM validation_results").fetchone()[0],
                'total_macro': db.execute("SELECT COUNT(*) FROM macro_processing").fetchone()[0],
                'route_stats': route_stats
            }
        else:
            recent_files = db.execute('''SELECT * FROM files 
                                       WHERE user_id=? 
                                       ORDER BY upload_date DESC LIMIT 5''',
                                      (session['user_id'],)).fetchall()

            recent_macro = db.execute('''SELECT * FROM macro_processing 
                                       WHERE user_id=? 
                                       ORDER BY processing_date DESC LIMIT 5''',
                                      (session['user_id'],)).fetchall()

            # User-specific route stats
            route_stats = {}
            for route_type in ROUTE_MACROS.keys():
                count = db.execute("SELECT COUNT(*) FROM macro_processing WHERE user_id = ? AND route_type = ?",
                                   (session['user_id'], route_type)).fetchone()[0]
                route_stats[route_type] = count

            admin_stats = {'route_stats': route_stats}

    return render_template("dashboard.html",
                           recent_files=recent_files,
                           recent_macro=recent_macro,
                           admin_stats=admin_stats,
                           route_macros=ROUTE_MACROS)

@main_bp.route('/history', strict_slashes=False)
def file_history():
    if not g.user:
        flash("Please log in to view file history", "error")
        return redirect(url_for('auth.login'))

    page = int(request.args.get('page', 1))
    per_page = 10
    offset = (page - 1) * per_page
    route_filter = request.args.get('route', 'all')

    with get_db() as conn:
        cursor = conn.cursor()

        # Filter logic
        filter_condition = ""
        params = []

        if route_filter != "all":
            if route_filter == "validation":
                filter_condition = "WHERE type = 'validation'"
            else:
                filter_condition = "WHERE type = 'macro' AND route_type = ?"
                params.append(route_filter)

        # Admin vs User-specific
        if session.get("is_admin"):
            user_condition = ""
        else:
            user_condition = "AND user_id = ?" if filter_condition else "WHERE user_id = ?"
            params.append(g.user["id"])

        # Unified query
        query = f"""
            SELECT * FROM (
                SELECT f.id,
                       f.original_filename AS original_filename,
                       f.upload_date AS date,
                       f.report_filename,
                       v.total_references,
                       v.total_citations,
                       u.username,
                       'validation' AS type,
                       '' AS route_type,
                       '' AS token,
                       '' AS selected_tasks,
                       '' AS original_filenames,
                       f.user_id
                FROM files f
                LEFT JOIN validation_results v ON f.id = v.file_id
                JOIN users u ON f.user_id = u.id

                UNION ALL

                SELECT m.id,
                       '' AS original_filename,
                       m.processing_date AS date,
                       '' AS report_filename,
                       0 AS total_references,
                       0 AS total_citations,
                       u.username,
                       'macro' AS type,
                       m.route_type AS route_type, 
                       m.token,
                       m.selected_tasks,
                       m.original_filenames,
                       m.user_id
                FROM macro_processing m
                JOIN users u ON m.user_id = u.id
            ) combined
            {filter_condition}
            {user_condition}
            ORDER BY date DESC
            LIMIT ? OFFSET ?
        """

        params.extend([per_page, offset])
        cursor.execute(query, params)
        history = cursor.fetchall()

        # Count total records for pagination
        count_query = f"""
            SELECT COUNT(*) FROM (
                SELECT f.id, f.user_id, 'validation' AS type
                FROM files f
                UNION ALL
                SELECT m.id, m.user_id, 'macro' AS type
                FROM macro_processing m
            ) combined
            {filter_condition}
            {user_condition}
        """
        cursor.execute(count_query, params[:-2])  # exclude LIMIT/OFFSET
        total_records = cursor.fetchone()[0]

    total_pages = (total_records + per_page - 1) // per_page

    return render_template(
        "file_history.html",
        history=history,
        page=page,
        total_pages=total_pages,
        route_filter=route_filter,
        route_macros=ROUTE_MACROS
    )

@main_bp.route("/report/<filename>")
def download_report(filename):
    if 'user_id' not in session:
        return redirect(url_for('auth.login'))

    safe_filename = secure_filename(filename)
    if safe_filename != filename:
        flash("Invalid filename", "error")
        return redirect(url_for('main.dashboard'))

    with get_db() as db:
        if session.get('is_admin'):
            file_exists = db.execute('SELECT 1 FROM files WHERE report_filename=?',
                                     (filename,)).fetchone()
        else:
            file_exists = db.execute('SELECT 1 FROM files WHERE report_filename=? AND user_id=?',
                                     (filename, session['user_id'])).fetchone()

        if not file_exists:
            flash("No permission to access this report", "error")
            return redirect(url_for('main.dashboard'))

    report_path = os.path.join(REPORT_FOLDER, filename)
    if not os.path.exists(report_path):
        flash("Report file not found", "error")
        return redirect(url_for('main.dashboard'))

    try:
        return send_from_directory(REPORT_FOLDER, filename, as_attachment=True, download_name=f"report_{filename}")
    except FileNotFoundError:
        flash("Report file could not be downloaded", "error")
        return redirect(url_for('main.dashboard'))

@main_bp.route("/download/<filename>")
def download_file(filename):
    if 'user_id' not in session:
        return redirect(url_for('auth.login'))

    safe_filename = secure_filename(filename)
    file_path = os.path.join(UPLOAD_FOLDER, safe_filename)
    
    if not os.path.exists(file_path):
        flash("File not found", "error")
        return redirect(url_for('main.dashboard'))

    return send_from_directory(UPLOAD_FOLDER, safe_filename, as_attachment=True)

@main_bp.route('/macro-reset', methods=['POST'])
def macro_reset_application():
    try:
        if os.path.exists(UPLOAD_FOLDER):
            shutil.rmtree(UPLOAD_FOLDER)
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        download_tokens.clear()
        return jsonify({"success": True, "message": "All files are cleared"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})

@main_bp.route("/reset-db", methods=["POST"])
@admin_required
def reset_database():
    if os.path.exists(DATABASE):
        os.remove(DATABASE)
    init_db()
    return "Database reset successfully! New admin created: username='admin', password='admin123'"
