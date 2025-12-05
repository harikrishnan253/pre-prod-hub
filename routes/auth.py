from flask import Blueprint, render_template, request, redirect, url_for, flash, session, current_app
from werkzeug.security import check_password_hash, generate_password_hash
from utils import log_activity
import traceback

auth_bp = Blueprint('auth', __name__)

@auth_bp.route('/login', methods=['GET', 'POST'], strict_slashes=False)
def login():
    if 'user_id' in session:
        return redirect(url_for('main.dashboard'))

    if request.method == "POST":
        username = request.form.get('username', '')
        password = request.form.get('password', '')

        try:
            current_app.logger.info(f"Login attempt for username: {username}")
            from app import db_pool
            with db_pool.get_connection() as db:
                user = db.execute("SELECT id, username, password, is_admin FROM users WHERE username = ?", (username,)).fetchone()
                current_app.logger.info(f"User found: {user is not None}")
                
                if user:
                    stored_hash = user['password']
                    # Check for legacy or non-standard hash format
                    if stored_hash.startswith('$'):
                        stored_hash = stored_hash[1:]

                    if check_password_hash(stored_hash, password):
                        current_app.logger.info(f"Password check passed for {username}")
                        session['user_id'] = user['id']
                        session['username'] = user['username']
                        session['is_admin'] = bool(user['is_admin'])
                        log_activity(username, "LOGIN")
                        flash("Login successful", "success")
                        current_app.logger.info(f"Redirecting to dashboard for {username}")
                        return redirect(url_for('main.dashboard'))
                    else:
                        current_app.logger.info(f"Password check failed for {username}")

                flash("Invalid username or password", "error")
        except Exception as e:
            current_app.logger.error(f"Login error: {str(e)}")
            current_app.logger.error(traceback.format_exc())
            flash("An error occurred during login", "error")

    return render_template('login.html')


@auth_bp.route("/register", methods=["GET", "POST"], strict_slashes=False)
def register():
    if request.method == "POST":
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        email = request.form.get('email', '')

        try:
            from app import db_pool
            with db_pool.get_connection() as db:
                existing = db.execute("SELECT id FROM users WHERE username = ?", (username,)).fetchone()
                if existing:
                    flash("Username already exists", "error")
                else:
                    hashed = generate_password_hash(password, method='pbkdf2:sha256')
                    db.execute("INSERT INTO users (username, password, email) VALUES (?, ?, ?)",
                               (username, hashed, email))
                    db.commit()
                    flash("Registration successful", "success")
                    return redirect(url_for('auth.login'))
        except Exception as e:
            current_app.logger.error(f"Registration error: {e}")
            flash(f"Registration failed: {str(e)}", "error")

    return render_template("register.html")


@auth_bp.route('/logout', strict_slashes=False)
def logout():
    user = session.get('username')
    if user:
        log_activity(user, "LOGOUT")
    session.clear()
    flash("Logged out successfully.")
    return redirect(url_for('auth.login'))