from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
from flask_login import UserMixin

db = SQLAlchemy()

class User(UserMixin, db.Model):
    __tablename__ = 'users'
    
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    email = db.Column(db.String(150))
    is_admin = db.Column(db.Boolean, default=False)
    role = db.Column(db.String(50), default='USER')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    # Relationships
    files = db.relationship('File', backref='user', lazy=True)
    macro_processes = db.relationship('MacroProcessing', backref='user', lazy=True)

class File(db.Model):
    __tablename__ = 'files'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    original_filename = db.Column(db.String(255), nullable=False)
    stored_filename = db.Column(db.String(255), nullable=False)
    report_filename = db.Column(db.String(255))
    upload_date = db.Column(db.DateTime, default=datetime.utcnow)

    # Relationships
    validation_results = db.relationship('ValidationResult', backref='file', uselist=False, lazy=True)

class ValidationResult(db.Model):
    __tablename__ = 'validation_results'
    
    id = db.Column(db.Integer, primary_key=True)
    file_id = db.Column(db.Integer, db.ForeignKey('files.id'), nullable=False)
    total_references = db.Column(db.Integer)
    total_citations = db.Column(db.Integer)
    missing_references = db.Column(db.Text)  # Stored as JSON string or text
    unused_references = db.Column(db.Text)   # Stored as JSON string or text
    sequence_issues = db.Column(db.Text)     # Stored as JSON string or text

class MacroProcessing(db.Model):
    __tablename__ = 'macro_processing'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    token = db.Column(db.String(100), unique=True, nullable=False)
    original_filenames = db.Column(db.Text, nullable=False)
    processed_filenames = db.Column(db.Text, nullable=False)
    selected_tasks = db.Column(db.Text, nullable=False)
    processing_date = db.Column(db.DateTime, default=datetime.utcnow)
    errors = db.Column(db.Text)
    route_type = db.Column(db.String(50), default='general')
