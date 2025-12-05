from functools import wraps
from flask import session, flash, redirect, url_for, g

def get_user_role():
    return session.get('role') or (g.user.get('role') if hasattr(g, 'user') and g.user else None)

def has_role(*roles):
    role = get_user_role()
    if role is None:
        return False
    # Convert role to string in case it's stored as an integer
    role_str = str(role).upper() if role else None
    return role_str is not None and role_str in [str(r).upper() for r in roles]

def role_required(allowed_roles):
    def decorator(f):
        @wraps(f)
        def wrapped(*args, **kwargs):
            if 'user_id' not in session:
                flash("Please log in to continue.")
                return redirect(url_for('auth.login'))
            if not has_role(*allowed_roles) and not session.get('is_admin'):
                flash("You don't have permission to access this page.", "error")
                return redirect(url_for('main.dashboard'))
            return f(*args, **kwargs)
        return wrapped
    return decorator

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('is_admin'):
            flash("Admin privileges required", "error")
            return redirect(url_for('main.dashboard'))
        return f(*args, **kwargs)

    return decorated_function
