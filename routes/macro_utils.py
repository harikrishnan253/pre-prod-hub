import os
import uuid
import json
import traceback
from datetime import datetime
from flask import request, redirect, url_for, flash, session, render_template
from werkzeug.utils import secure_filename
from database import get_db
from utils import log_activity, log_errors, allowed_file, save_uploaded_file
from routes.ppd import html_to_excel_no_images
from config import ROUTE_MACROS, UPLOAD_FOLDER, TOKEN_TTL
from shared_state import download_tokens, download_tokens_lock, WORD_LOCK
from word_processor import OptimizedDocumentProcessor
import sys

sys.path.append(os.path.dirname(os.path.dirname(__file__)))

def process_macro_request(route_type, redirect_endpoint):
    """
    Generic handler for macro routes.
    """
    word_files = request.files.getlist('word_files[]')
    selected_tasks = request.form.getlist('tasks[]')
    user_id = session.get('user_id')
    username = session.get('username', 'unknown')

    if not word_files or not selected_tasks:
        flash("Please upload files and select at least one task.")
        return redirect(url_for(redirect_endpoint))

    token = uuid.uuid4().hex
    unique_folder = os.path.join(UPLOAD_FOLDER, token)
    os.makedirs(unique_folder, exist_ok=True)

    # Register download token (thread-safe)
    with download_tokens_lock:
        download_tokens[token] = {
            'path': unique_folder,
            'expires': datetime.now() + TOKEN_TTL,
            'user': username,
            'route_type': route_type
        }

    word_paths = []
    original_filenames = []

    for f in word_files:
        if f and allowed_file(f.filename):
            filename = secure_filename(f.filename)
            save_path = os.path.join(unique_folder, filename)
            try:
                f.save(save_path)
                word_paths.append(save_path)
                original_filenames.append(filename)
            except Exception as e:
                log_errors([f"Error saving uploaded file {filename}: {str(e)}"])

    if not word_paths:
        flash("No valid Word files uploaded.")
        return redirect(url_for(redirect_endpoint))

    all_errors = []

    try:
        with WORD_LOCK:
            with OptimizedDocumentProcessor() as processor:
                try:
                    batch_errors = processor.process_documents_batch(word_paths, selected_tasks, route_type)
                    if batch_errors:
                        all_errors.extend(batch_errors)
                except Exception as e:
                    all_errors.append(f"Batch processing failed: {str(e)}")
                    log_errors([traceback.format_exc()])

                for doc_path in word_paths:
                    log_activity(username, f"MACRO_PROCESS_{route_type.upper()}",
                                 details=os.path.basename(doc_path))

    except Exception as e:
        all_errors.append(f"Processing failed: {str(e)}")
        log_errors([traceback.format_exc()])

    # PPD-specific processing
    if route_type.lower() == 'ppd':
        try:
            if os.path.exists(unique_folder):
                html_files = [f for f in os.listdir(unique_folder) if f.lower().endswith(".html")]
            else:
                html_files = []

            for file in html_files:
                html_path = os.path.join(unique_folder, file)
                html_to_excel_no_images(html_path, unique_folder)
        except Exception as e:
            error_msg = f"HTML to Excel conversion failed: {str(e)}"
            all_errors.append(error_msg)
            log_errors([error_msg])

    # Store in database
    try:
        selected_macro_names = []
        route_macros = ROUTE_MACROS.get(route_type, {}).get('macros', [])
        for task_idx in selected_tasks:
            try:
                idx = int(task_idx)
                if 0 <= idx < len(route_macros):
                    selected_macro_names.append(route_macros[idx])
            except Exception:
                pass

        processed_filenames = list(original_filenames)
        try:
            if os.path.exists(unique_folder):
                for root, _, files in os.walk(unique_folder):
                    for fn in files:
                        if fn not in processed_filenames:
                            processed_filenames.append(fn)
        except Exception:
            pass

        with get_db() as db:
            db.execute('''INSERT INTO macro_processing 
                          (user_id, token, original_filenames, processed_filenames, selected_tasks, errors, route_type)
                          VALUES (?, ?, ?, ?, ?, ?, ?)''',
                       (user_id, token,
                        json.dumps(original_filenames),
                        json.dumps(processed_filenames),
                        json.dumps({
                            'route_type': route_type,
                            'task_indices': selected_tasks,
                            'macro_names': selected_macro_names
                        }),
                        json.dumps(all_errors) if all_errors else None,
                        route_type))
            db.commit()
    except Exception as e:
        log_errors([f"Error saving macro processing: {str(e)}"])

    route_name = ROUTE_MACROS.get(route_type, {}).get('name', 'Processing')
    if all_errors:
        flash(f"{route_name} completed with some errors. Check log for details.")
        log_errors(all_errors)
    else:
        flash(f"{route_name} completed successfully!")

    return redirect(url_for(redirect_endpoint, download_token=token))

def handle_macro_route(route_type, template_name, redirect_endpoint):
    if 'user_id' not in session:
        flash("Please log in to continue.")
        return redirect(url_for('auth.login'))

    if request.method == 'POST':
        return process_macro_request(route_type, redirect_endpoint)

    download_token = request.args.get('download_token')
    route_config = ROUTE_MACROS.get(route_type, {})

    return render_template(template_name,
                           download_token=download_token,
                           route_config=route_config,
                           macro_names=route_config.get('macros', []))
