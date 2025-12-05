import os
import json
import traceback
import sqlite3
from flask import Blueprint, render_template, request, redirect, url_for, flash, session
from werkzeug.security import generate_password_hash
from database import get_db
from utils import log_activity, log_errors
from config import ROUTE_MACROS, UPLOAD_FOLDER, REPORT_FOLDER
from auth_utils import admin_required

admin_bp = Blueprint('admin', __name__)

@admin_bp.route("/admin", strict_slashes=False)
@admin_required
def admin_dashboard():
    with get_db() as db:
        route_stats = {}
        for route_type in ROUTE_MACROS.keys():
            count = db.execute(
                "SELECT COUNT(*) FROM macro_processing WHERE route_type = ?",
                (route_type,)
            ).fetchone()[0]
            route_stats[route_type] = count

        # totals
        total_users = db.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        total_files = db.execute("SELECT COUNT(*) FROM files").fetchone()[0]
        total_validations = db.execute("SELECT COUNT(*) FROM validation_results").fetchone()[0]
        total_macro = db.execute("SELECT COUNT(*) FROM macro_processing").fetchone()[0]

        # roles
        try:
            role_counts = db.execute(
                "SELECT role, COUNT(*) as count FROM users GROUP BY role"
            ).fetchall()
            role_stats = {
                (r["role"] if r["role"] else "USER"): r["count"] for r in role_counts
            }
        except sqlite3.OperationalError as e:
            log_errors([f"Role stats query failed: {e}"])
            role_stats = {}

        admin_stats = {
            'total_users': total_users,
            'total_files': total_files,
            'total_validations': total_validations,
            'total_macro': total_macro,
            'route_stats': route_stats
        }

    return render_template(
        "admin_dashboard.html",
        admin_stats=admin_stats,
        route_macros=ROUTE_MACROS,
        role_stats=role_stats
    )

@admin_bp.route("/admin/user/<int:user_id>/change-role", methods=["POST"], strict_slashes=False)
@admin_required
def admin_change_role(user_id):
    new_role = request.form.get('role', '').upper()
    if not new_role:
        flash("No role provided", "error")
        return redirect(url_for('admin.admin_users'))

    if user_id == session.get('user_id'):
        flash("Cannot change your own role", "error")
        return redirect(url_for('admin.admin_users'))

    with get_db() as db:
        user = db.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        if not user:
            flash("User not found", "error")
            return redirect(url_for('admin.admin_users'))

        db.execute("UPDATE users SET role=? WHERE id=?", (new_role, user_id))
        db.commit()
        flash("User role updated", "success")
        log_activity(session['username'], 'CHANGE_ROLE', f"user:{user['username']} -> {new_role}")
    return redirect(url_for('admin.admin_users'))

@admin_bp.route("/admin/users", strict_slashes=False)
@admin_required
def admin_users():
    with get_db() as db:
        users = db.execute(
            'SELECT id, username, email, is_admin, role, created_at FROM users ORDER BY created_at DESC').fetchall()
    return render_template("admin_users.html", users=users)

@admin_bp.route("/admin/create-user", methods=["GET", "POST"], strict_slashes=False)
@admin_required
def admin_create_user():
    if request.method == "POST":
        username = request.form['username']
        password = request.form['password']
        email = request.form.get('email', '')
        is_admin = 'is_admin' in request.form
        role = request.form.get('role', 'USER').upper()

        with get_db() as db:
            try:
                hashed = generate_password_hash(password, method='pbkdf2:sha256')
                db.execute("INSERT INTO users (username,password,email,is_admin,role) VALUES (?,?,?,?,?)",
                           (username, hashed, email, is_admin, role))
                db.commit()
                flash("User created successfully", "success")
                return redirect(url_for('admin.admin_users'))
            except sqlite3.IntegrityError:
                db.rollback()
                flash("Username/email exists", "error")

    return render_template("admin_create_user.html")

@admin_bp.route('/admin/change_password/<int:user_id>', methods=['GET', 'POST'], strict_slashes=False)
@admin_required
def admin_change_password(user_id):
    with get_db() as db:
        user = db.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()

        if not user:
            flash("User not found.", "error")
            return redirect(url_for('admin.admin_users'))

        if request.method == 'POST':
            new_password = request.form['new_password']
            hashed = generate_password_hash(new_password, method='pbkdf2:sha256')
            db.execute("UPDATE users SET password = ? WHERE id = ?", (hashed, user_id))
            db.commit()
            flash(f"Password updated for {user['username']}.", "success")
            return redirect(url_for('admin.admin_users'))

    return render_template("admin_change_password.html", user=user)

@admin_bp.route("/admin/user/<int:user_id>/toggle-admin", methods=["POST"], strict_slashes=False)
@admin_required
def admin_toggle_admin(user_id):
    if user_id == session.get('user_id'):
        flash("Cannot change your own admin status", "error")
        return redirect(url_for('admin.admin_users'))

    with get_db() as db:
        user = db.execute("SELECT is_admin FROM users WHERE id=?", (user_id,)).fetchone()
        if not user:
            flash("User not found", "error")
            return redirect(url_for('admin.admin_users'))

        new_status = not bool(user['is_admin'])
        db.execute("UPDATE users SET is_admin=? WHERE id=?", (new_status, user_id))
        db.commit()
        status_text = "granted" if new_status else "revoked"
        flash(f"Admin privileges {status_text}", "success")

    return redirect(url_for('admin.admin_users'))

@admin_bp.route("/admin/user/<int:user_id>/delete", methods=["POST"], strict_slashes=False)
@admin_required
def admin_delete_user(user_id):
    if user_id == session.get('user_id'):
        flash("Cannot delete your own account", "error")
        return redirect(url_for('admin.admin_users'))

    try:
        with get_db() as db:
            macro_count = db.execute(
                "SELECT COUNT(*) FROM macro_processing WHERE user_id=?",
                (user_id,)
            ).fetchone()[0]

            if macro_count > 0:
                flash("Cannot delete user with macro history", "error")
                return redirect(url_for('admin.admin_users'))

            user_files = db.execute(
                "SELECT COUNT(*) FROM files WHERE user_id=?",
                (user_id,)
            ).fetchone()[0]

            if user_files > 0:
                flash("Cannot delete user with files", "error")
                return redirect(url_for('admin.admin_users'))

            try:
                db.execute("DELETE FROM validation_results WHERE file_id IN (SELECT id FROM files WHERE user_id=?)", (user_id,))
            except Exception:
                pass

            try:
                db.execute("DELETE FROM files WHERE user_id=?", (user_id,))
            except Exception:
                pass

            db.execute("DELETE FROM macro_processing WHERE user_id=?", (user_id,))
            db.execute("DELETE FROM users WHERE id=?", (user_id,))
            db.commit()

            flash("User deleted successfully", "success")
            log_activity(session.get('username', 'system'), "DELETE_USER", f"user_id:{user_id}")
            return redirect(url_for('admin.admin_users'))

    except Exception as e:
        log_errors([f"Error deleting user {user_id}: {e}", traceback.format_exc()])
        flash("An error occurred while deleting the user", "error")
        return redirect(url_for('admin.admin_users'))

@admin_bp.route("/admin/files")
@admin_required
def admin_files():
    page = request.args.get('page', 1, type=int)
    per_page = 10
    offset = (page - 1) * per_page

    with get_db() as db:
        files = db.execute('''SELECT f.*, u.username, v.total_references, v.total_citations
                           FROM files f
                           JOIN users u ON f.user_id = u.id
                           LEFT JOIN validation_results v ON f.id = v.file_id
                           ORDER BY f.upload_date DESC LIMIT ? OFFSET ?''',
                           (per_page, offset)).fetchall()

        total_count = db.execute("SELECT COUNT(*) FROM files").fetchone()[0]
        total_pages = (total_count + per_page - 1) // per_page

    return render_template("admin_files.html", files=files, page=page, total_pages=total_pages)

@admin_bp.route("/admin/file/<int:file_id>/delete", methods=["POST"])
@admin_required
def admin_delete_file(file_id):
    with get_db() as db:
        file = db.execute("SELECT * FROM files WHERE id=?", (file_id,)).fetchone()
        if not file:
            flash("File not found", "error")
            return redirect(url_for('admin.admin_files'))

        try:
            file_path = os.path.join(UPLOAD_FOLDER, file['stored_filename'])
            if os.path.exists(file_path):
                os.remove(file_path)

            if file['report_filename']:
                report_path = os.path.join(REPORT_FOLDER, file['report_filename'])
                if os.path.exists(report_path):
                    os.remove(report_path)
        except Exception as e:
            flash(f"Error deleting file: {str(e)}", "error")
            return redirect(url_for('admin.admin_files'))

        db.execute("DELETE FROM validation_results WHERE file_id=?", (file_id,))
        db.execute("DELETE FROM files WHERE id=?", (file_id,))
        db.commit()

        flash("File deleted successfully", "success")
        return redirect(url_for('admin.admin_files'))

@admin_bp.route("/admin/stats")
@admin_required
def admin_stats():
    with get_db() as db:
        recent_files = db.execute('''SELECT f.*, u.username 
                                   FROM files f 
                                   JOIN users u ON f.user_id = u.id 
                                   ORDER BY f.upload_date DESC LIMIT 10''').fetchall()

        user_stats = db.execute('''SELECT u.username, COUNT(f.id) as file_count
                                FROM users u
                                LEFT JOIN files f ON u.id = f.user_id
                                GROUP BY u.id
                                ORDER BY file_count DESC''').fetchall()

        total_users = db.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        total_files = db.execute("SELECT COUNT(*) FROM files").fetchone()[0]
        total_validations = db.execute("SELECT COUNT(*) FROM validation_results").fetchone()[0]
        total_macro = db.execute("SELECT COUNT(*) FROM macro_processing").fetchone()[0]

        role_counts = db.execute("SELECT role, COUNT(*) as count FROM users GROUP BY role").fetchall()
        role_stats = {r["role"]: r["count"] for r in role_counts}

    return render_template(
        "admin_stats.html",
        recent_files=recent_files,
        user_stats=user_stats,
        admin_stats={
            'total_users': total_users,
            'total_files': total_files,
            'total_validations': total_validations,
            'total_macro': total_macro
        },
        role_stats=role_stats
    )

@admin_bp.route("/admin/macro-stats")
@admin_required
def admin_macro_stats():
    with get_db() as db:
        macro_records = db.execute('''SELECT selected_tasks, processing_date, errors, route_type
                                    FROM macro_processing 
                                    ORDER BY processing_date DESC''').fetchall()

    route_stats = {}
    error_stats = {}
    daily_stats = {}

    for record in macro_records:
        route_type = record['route_type'] or 'unknown'

        if route_type not in route_stats:
            route_stats[route_type] = 0
        route_stats[route_type] += 1

        if record['errors']:
            try:
                error_count = len(json.loads(record['errors']))
                if route_type not in error_stats:
                    error_stats[route_type] = 0
                error_stats[route_type] += error_count
            except:
                pass

        date = record['processing_date'][:10]
        if date not in daily_stats:
            daily_stats[date] = {}
        if route_type not in daily_stats[date]:
            daily_stats[date][route_type] = 0
        daily_stats[date][route_type] += 1

    return render_template("admin_macro_stats.html",
                           route_stats=route_stats,
                           error_stats=error_stats,
                           daily_stats=daily_stats,
                           route_macros=ROUTE_MACROS)

@admin_bp.route("/admin/macro-history")
@admin_required
def admin_macro_history():
    page = request.args.get('page', 1, type=int)
    per_page = 10
    offset = (page - 1) * per_page

    with get_db() as db:
        macro_history = db.execute('''SELECT m.*, u.username
                                    FROM macro_processing m
                                    JOIN users u ON m.user_id = u.id
                                    ORDER BY m.processing_date DESC LIMIT ? OFFSET ?''',
                                   (per_page, offset)).fetchall()

        total_count = db.execute("SELECT COUNT(*) FROM macro_processing").fetchone()[0]
        total_pages = (total_count + per_page - 1) // per_page

    return render_template("admin_macro_history.html",
                           macro_history=macro_history,
                           page=page,
                           total_pages=total_pages,
                           macro_names=ROUTE_MACROS.get('macro_processing', {}).get('macros', []))
