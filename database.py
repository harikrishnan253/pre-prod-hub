from models import db
from flask import current_app

def init_db():
    """Initialize the database."""
    with current_app.app_context():
        db.create_all()
        
        # Create default admin if not exists
        from models import User
        from werkzeug.security import generate_password_hash
        
        admin = User.query.filter_by(username='admin').first()
        if not admin:
            hashed_password = generate_password_hash("admin123", method='pbkdf2:sha256')
            new_admin = User(username='admin', password=hashed_password, email='admin@example.com', is_admin=True, role='ADMIN')
            db.session.add(new_admin)
            db.session.commit()
            print("Created default admin user.")

def get_db():
    """
    Legacy support for raw connection.
    Ideally, use db.session in new code.
    """
    # This is a bit hacky for legacy support, but allows gradual migration.
    # We return a context manager that yields a connection.
    from contextlib import contextmanager
    import sqlite3
    
    @contextmanager
    def db_context():
        # We can't easily get a raw sqlite3 connection from SQLAlchemy that behaves exactly like the old one
        # (especially with row_factory=sqlite3.Row).
        # So for legacy parts, we might still want to open a raw connection or try to wrap SQLAlchemy connection.
        # Given the "Next Level" goal, let's try to use the engine's connection.
        
        # However, to be safe and avoid breaking everything immediately:
        # We will open a fresh connection using the config path.
        from config import DATABASE
        conn = sqlite3.connect(DATABASE, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
        finally:
            conn.close()

    return db_context()

def migrate_add_role_column(app_logger):
    # Handled by SQLAlchemy models/migrations usually, but keeping for safety
    pass

def migrate_add_route_type_column(app_logger):
    # Handled by SQLAlchemy models/migrations usually, but keeping for safety
    pass


