import os
import shutil
import io
import zipfile
from datetime import datetime
from flask import Blueprint, request, redirect, url_for, flash, send_file
from utils import log_errors
from config import ROUTE_PERMISSIONS, ROUTE_MACROS
from shared_state import download_tokens
from auth_utils import role_required
from routes.macro_utils import handle_macro_route

macros_bp = Blueprint('macros', __name__)

@macros_bp.route('/language', methods=['GET', 'POST'], strict_slashes=False)
@role_required(ROUTE_PERMISSIONS.get('language', ['ADMIN']))
def language():
    return handle_macro_route('language', 'language_edit.html', 'macros.language')

@macros_bp.route('/technical', methods=['GET', 'POST'])
@role_required(ROUTE_PERMISSIONS.get('technical', ['ADMIN']))
def technical():
    return handle_macro_route('technical', 'technical_edit.html', 'macros.technical')

@macros_bp.route('/macro_processing', methods=['GET', 'POST'])
@role_required(ROUTE_PERMISSIONS.get('macro_processing', ['ADMIN']))
def macro_processing():
    return handle_macro_route('macro_processing', 'macro_processing.html', 'macros.macro_processing')

@macros_bp.route('/macro-download', strict_slashes=False)
def macro_download():
    token = request.args.get('token')
    if not token:
        flash("Invalid download request.")
        return redirect(url_for('main.dashboard'))

    token_data = download_tokens.get(token)
    if not token_data:
        flash("Invalid or expired download token.")
        return redirect(url_for('main.dashboard'))

    if datetime.now() > token_data['expires']:
        flash("Download token has expired.")
        try:
            if os.path.exists(token_data['path']):
                shutil.rmtree(token_data['path'])
            del download_tokens[token]
        except Exception as e:
            log_errors([f"Error cleaning expired token: {str(e)}"])
        return redirect(url_for('main.dashboard'))

    user_folder = token_data['path']
    route_type = token_data.get('route_type', 'general')

    if not os.path.exists(user_folder):
        flash("No files found for this download token.")
        return redirect(url_for('main.dashboard'))

    try:
        memory_file = io.BytesIO()
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zipf:
            for root, _, files in os.walk(user_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    if os.path.getsize(file_path) < 50 * 1024 * 1024:
                        arcname = os.path.relpath(file_path, user_folder)
                        zipf.write(file_path, arcname)

        memory_file.seek(0)

        try:
            shutil.rmtree(user_folder)
            del download_tokens[token]
        except Exception as e:
            log_errors([f"Cleanup error: {str(e)}"])

        route_name = ROUTE_MACROS.get(route_type, {}).get('name', 'Processed')
        zip_filename = f"{route_name.replace(' ', '_')}_Documents_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

        return send_file(memory_file,
                         mimetype='application/zip',
                         as_attachment=True,
                         download_name=zip_filename)

    except Exception as e:
        flash(f"Download failed: {str(e)}")
        log_errors([f"Download error for token {token}: {str(e)}"])
        return redirect(url_for('main.dashboard'))
