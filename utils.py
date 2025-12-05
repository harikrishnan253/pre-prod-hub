import os
import socket
import logging
import subprocess
import shutil
import re
import json
import base64
import chardet
from datetime import datetime
from pathlib import Path
from logging.handlers import RotatingFileHandler
from werkzeug.utils import secure_filename
from config import LOG_FILE, ALLOWED_EXTENSIONS
from shared_state import download_tokens

def get_ip_address():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip_address = s.getsockname()[0]
        s.close()
        return ip_address
    except Exception:
        return "127.0.0.1"

def log_activity(username, action, details=""):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"{timestamp} - {username} - {action} - {details}\n")

def log_errors(error_list):
    with open(LOG_FILE, "a", encoding="utf-8") as log_file:
        for err in error_list:
            log_file.write(f"{datetime.now().isoformat()} - ERROR - {err}\n")

def allowed_file(filename):
    return any(filename.lower().endswith(ext) for ext in ALLOWED_EXTENSIONS)

def setup_logging(app):
    if not app.debug:
        file_handler = RotatingFileHandler('logs/s4c.log', maxBytes=10240000, backupCount=10)
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
        ))
        file_handler.setLevel(logging.INFO)
        app.logger.addHandler(file_handler)
        app.logger.setLevel(logging.INFO)

def cleanup_expired_tokens():
    current_time = datetime.now()
    # Create a copy to iterate safely, though we should probably lock if we were being 100% safe.
    # Since this is a background task, we might want to use the lock from shared_state if we export it.
    # For now, we'll just iterate a copy.
    expired = [t for t, data in download_tokens.items() if current_time > data['expires']]

    for token in expired:
        try:
            token_data = download_tokens[token]
            path = token_data['path']
            route_type = token_data.get('route_type', 'unknown')

            if os.path.exists(path):
                shutil.rmtree(path)

            log_activity(token_data.get('user', 'system'),
                         f"TOKEN_CLEANUP_{route_type.upper()}",
                         details=f"Token: {token[:8]}...")

            del download_tokens[token]

        except Exception as e:
            log_errors([f"Error cleaning expired token {token}: {str(e)}"])

def kill_word_processes():
    try:
        subprocess.run(["taskkill", "/f", "/im", "winword.exe"],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass

def save_uploaded_file(file, folder):
    try:
        filename = secure_filename(file.filename)
        os.makedirs(folder, exist_ok=True)
        base, ext = os.path.splitext(filename)
        file_path = os.path.join(folder, filename)

        try:
            with open(file_path, 'wb') as f:
                file.save(f)
            return file_path, None
        except PermissionError as pe:
            # Try to relax permissions if file exists
            try:
                if os.path.exists(file_path):
                    os.chmod(file_path, 0o666)
                    with open(file_path, 'wb') as f:
                        file.save(f)
                    return file_path, None
            except Exception:
                pass

            # Fallback: save with a unique filename to avoid locks/permissions issues
            import time, uuid
            alt_name = f"{base}_{int(time.time())}_{uuid.uuid4().hex[:8]}{ext}"
            alt_path = os.path.join(folder, alt_name)
            try:
                with open(alt_path, 'wb') as f:
                    file.save(f)
                return alt_path, None
            except Exception as e2:
                return None, f"Permission denied saving file: {e2}"
        except Exception as e:
            return None, str(e)
    except Exception as e:
        return None, str(e)

def get_base64_logo(static_folder):
    logo_path = os.path.join(static_folder, "images", "S4c.png")
    try:
        with open(logo_path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    except Exception as e:
        return ""


