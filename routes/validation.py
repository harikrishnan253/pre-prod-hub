import os
import json
from datetime import datetime
from flask import Blueprint, render_template, request, redirect, url_for, flash, session, jsonify
from werkzeug.utils import secure_filename
from database import get_db
from utils import log_errors, allowed_file, save_uploaded_file
from config import UPLOAD_FOLDER, REPORT_FOLDER
from validator import ReferenceValidator

validation_bp = Blueprint('validation', __name__)

@validation_bp.route('/validate', methods=['GET', 'POST'], strict_slashes=False)
def validate_file():
    if 'user_id' not in session:
        flash("Please log in to continue.")
        return redirect(url_for('auth.login'))

    if request.method == 'POST':
        # Check if this is an AJAX request
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        
        # Handle both 'file' and 'files' field names
        files = request.files.getlist('files') or [request.files.get('file')]
        files = [f for f in files if f and f.filename]
        
        if not files:
            if is_ajax:
                return jsonify({'success': False, 'error': 'No files uploaded'}), 400
            flash('No file selected')
            return redirect(request.url)

        processed_files = []
        errors = []

        for file in files:
            if not allowed_file(file.filename):
                errors.append(f"{file.filename}: Invalid file type")
                continue

            file_path, error = save_uploaded_file(file, UPLOAD_FOLDER)
            if error:
                errors.append(f"{file.filename}: {error}")
                continue

            try:
                validator = ReferenceValidator(file_path)
                
                # Define renumbered path
                renumbered_filename = f"renumbered_{file.filename}"
                renumbered_path = os.path.join(UPLOAD_FOLDER, renumbered_filename)
                
                with validator:
                    # Enable auto-renumbering
                    results = validator.validate(auto_renumber=True, save_path=renumbered_path)

                # Check if renumbering occurred
                if results.get('renumber_attempt', {}).get('renumbered'):
                    renumbered_url = url_for('main.download_file', filename=renumbered_filename)
                    processed_files.append({
                        'filename': file.filename,
                        'renumbered_url': renumbered_url,
                        'message': 'File was renumbered due to sequence issues.'
                    })

                # Save results to DB
                with get_db() as db:
                    # Insert file record
                    cursor = db.execute(
                        "INSERT INTO files (user_id, original_filename, stored_filename) VALUES (?, ?, ?)",
                        (session['user_id'], file.filename, os.path.basename(file_path))
                    )
                    file_id = cursor.lastrowid

                    # Insert validation results
                    db.execute('''INSERT INTO validation_results 
                                (file_id, total_references, total_citations, missing_references, 
                                 unused_references, sequence_issues) 
                                VALUES (?, ?, ?, ?, ?, ?)''',
                               (file_id,
                                results['total_references'],
                                results['total_citations'],
                                json.dumps(results['missing_references']),
                                json.dumps(results['unused_references']),
                                json.dumps(results['sequence_issues'])))
                    db.commit()

                # Generate HTML report
                os.makedirs(REPORT_FOLDER, exist_ok=True)
                report_filename = f"validation_{file_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
                report_path = os.path.join(REPORT_FOLDER, report_filename)
                
                # Create HTML report
                html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>Validation Report - {file.filename}</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }}
        .container {{ max-width: 1000px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
        h1 {{ color: #333; border-bottom: 3px solid #4361ee; padding-bottom: 10px; }}
        h2 {{ color: #4361ee; margin-top: 30px; }}
        .summary {{ background: #e8f4f8; padding: 15px; border-radius: 5px; margin: 20px 0; }}
        .stat {{ display: inline-block; margin: 10px 20px 10px 0; }}
        .stat-label {{ font-weight: bold; color: #666; }}
        .stat-value {{ font-size: 24px; color: #4361ee; }}
        table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
        th {{ background: #4361ee; color: white; padding: 12px; text-align: left; }}
        td {{ padding: 10px; border-bottom: 1px solid #ddd; }}
        tr:hover {{ background: #f8f9fa; }}
        .success {{ color: #06d6a0; }}
        .warning {{ color: #ffd60a; }}
        .error {{ color: #ef476f; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>Reference Validation Report</h1>
        <p><strong>File:</strong> {file.filename}</p>
        <p><strong>Date:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        
        <div class="summary">
            <h2>Summary</h2>
            <div class="stat">
                <div class="stat-label">Total References</div>
                <div class="stat-value">{results['total_references']}</div>
            </div>
            <div class="stat">
                <div class="stat-label">Total Citations</div>
                <div class="stat-value">{results['total_citations']}</div>
            </div>
            <div class="stat">
                <div class="stat-label">Missing References</div>
                <div class="stat-value error">{len(results['missing_references'])}</div>
            </div>
            <div class="stat">
                <div class="stat-label">Unused References</div>
                <div class="stat-value warning">{len(results['unused_references'])}</div>
            </div>
        </div>
        
        <h2>Missing References</h2>
        <table>
            <thead>
                <tr><th>#</th><th>Citation</th></tr>
            </thead>
            <tbody>
                {"".join(f"<tr><td>{i+1}</td><td>{ref}</td></tr>" for i, ref in enumerate(results['missing_references'])) if results['missing_references'] else "<tr><td colspan='2'>No missing references</td></tr>"}
            </tbody>
        </table>
        
        <h2>Unused References</h2>
        <table>
            <thead>
                <tr><th>#</th><th>Reference</th></tr>
            </thead>
            <tbody>
                {"".join(f"<tr><td>{i+1}</td><td>{ref}</td></tr>" for i, ref in enumerate(results['unused_references'])) if results['unused_references'] else "<tr><td colspan='2'>No unused references</td></tr>"}
            </tbody>
        </table>
        
        <h2>Sequence Issues</h2>
        <table>
            <thead>
                <tr><th>#</th><th>Issue</th></tr>
            </thead>
            <tbody>
                {"".join(f"<tr><td>{i+1}</td><td>{issue}</td></tr>" for i, issue in enumerate(results['sequence_issues'])) if results['sequence_issues'] else "<tr><td colspan='2'>No sequence issues</td></tr>"}
            </tbody>
        </table>
    </div>
</body>
</html>
"""
                
                with open(report_path, 'w', encoding='utf-8') as f:
                    f.write(html_content)
                
                # Update database with report filename
                with get_db() as db:
                    db.execute("UPDATE files SET report_filename = ? WHERE id = ?", 
                              (report_filename, file_id))
                    db.commit()

                # Generate report URL
                report_url = url_for('main.download_report', filename=report_filename)
                processed_files.append({
                    'filename': file.filename,
                    'report_url': report_url
                })

            except Exception as e:
                log_errors([f"Validation failed for {file.filename}: {str(e)}"])
                errors.append(f"{file.filename}: {str(e)}")

        # Return JSON for AJAX requests
        if is_ajax:
            return jsonify({
                'success': len(processed_files) > 0,
                'processed_files': processed_files,
                'errors': errors
            })

        # Return HTML for regular form submissions
        if processed_files:
            flash(f"Successfully validated {len(processed_files)} file(s)", "success")
        if errors:
            for error in errors:
                flash(error, "error")
        
        return redirect(url_for('validation.validate_file'))

    return render_template('upload.html')
