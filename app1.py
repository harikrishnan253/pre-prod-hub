from flask import Flask, render_template, g, session
from flask_wtf.csrf import CSRFProtect
import os
import threading
import time
import json
from waitress import serve

# Import local modules
from config import ROUTE_MACROS
from database import init_db, get_db
from utils import setup_logging, log_errors, cleanup_expired_tokens, get_ip_address
from auth_utils import get_user_role
from models import db, User

# Import Blueprints
from routes.auth import auth_bp
from routes.main import main_bp
from routes.macros import macros_bp
from routes.ppd import ppd_bp
from routes.validation import validation_bp
from routes.admin import admin_bp
from routes.doi import doi_bp

def start_background_cleanup():
    def cleanup_worker():
        while True:
            try:
                cleanup_expired_tokens()
                time.sleep(300)  # Run every 5 minutes
            except Exception as e:
                log_errors([f"Background cleanup error: {str(e)}"])
    threading.Thread(target=cleanup_worker, daemon=True).start()

def create_app():
    app = Flask(__name__)
    app.secret_key = os.environ.get('SECRET_KEY') or os.urandom(24)
    
    # Configure Database
    app.config['SQLALCHEMY_DATABASE_URI'] = f"sqlite:///{os.path.abspath('reference_validator.db')}"
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
    
    db.init_app(app)
    CSRFProtect(app)

    # Register Blueprints
    app.register_blueprint(auth_bp)
    app.register_blueprint(main_bp)
    app.register_blueprint(macros_bp)
    app.register_blueprint(ppd_bp)
    app.register_blueprint(validation_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(doi_bp)

    # Before Request Handler
    @app.before_request
    def load_user():
        g.user = None
        if 'user_id' in session:
            # Use Session.get() to avoid SQLAlchemy legacy Query.get() warning
            try:
                user = db.session.get(User, session['user_id'])
            except Exception:
                # Fallback to Query.get if session.get isn't available for some reason
                user = User.query.get(session['user_id'])
            if user:
                g.user = {
                    'id': user.id,
                    'username': user.username,
                    'is_admin': user.is_admin,
                    'role': user.role
                }

    # Context Processors
    @app.context_processor
    def inject_user_role():
        current_role = get_user_role()
        return dict(get_user_role=get_user_role, current_role=current_role)

    @app.context_processor
    def inject_route_macros():
        return dict(ROUTE_MACROS=ROUTE_MACROS)

    # Template Filters
    @app.template_filter('datetime')
    def format_datetime(value):
        if value:
            return value.replace('T', ' ')[:19]
        return ''

    @app.template_filter('from_json')
    def from_json_filter(value):
        if value:
            try:
                return json.loads(value)
            except (json.JSONDecodeError, TypeError):
                return []
        return []

    @app.template_filter('format_date')
    def format_date(value):
        """Return a formatted date string (YYYY-MM-DD)."""
        if not value:
            return ''
        if isinstance(value, str):
            return value.split('T')[0]
        try:
            return value.strftime('%Y-%m-%d')
        except Exception:
            return str(value)

    # Error Handlers
    @app.errorhandler(404)
    def page_not_found(e):
        return render_template('404.html'), 404

    @app.errorhandler(500)
    def internal_server_error(e):
        return render_template('500.html'), 500

    @app.errorhandler(Exception)
    def handle_unexpected_error(error):
        app.logger.error(f'Unexpected error: {error}')
        if app.debug:
            return str(error), 500
        return 'An unexpected error occurred', 500

    # Initialize
    setup_logging(app)
    with app.app_context():
        init_db()
    start_background_cleanup()

    # Load PPD macros if available
    try:
        import PPD_Final as ppd
        if hasattr(ppd, 'macro_names') and isinstance(ppd.macro_names, (list, tuple)):
            ROUTE_MACROS['ppd']['macros'] = ppd.macro_names
            app.logger.info("Loaded PPD macros successfully")
    except ImportError:
        app.logger.info("PPD_Final module not found, skipping PPD macro loading")
    except Exception as e:
        log_errors([f"Failed to load PPD macro names: {e}"])

    # Initialize PPD Progress Data
    app.config["PROGRESS_DATA"] = {}

    app.logger.info("Application initialized with modular routes")
    return app

# Create global app instance for Waitress
app = create_app()

if __name__ == '__main__':
    print("=== S4C APPLICATION STARTUP ===")
    host_ip = get_ip_address()
    print(f"Your IP address: {host_ip}")
    port = 5001
    print("\nAccess URLs:")
    print(f"Local: http://localhost:{port}")
    print(f"Network: http://{host_ip}:{port}")
    print("=================================\n")
    serve(app, host="0.0.0.0", port=port, threads=4)