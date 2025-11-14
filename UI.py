fron werkzeug.utils import secure_filename

from flask import Flask, render template, request, redirect, url for, send file, session, flash

import os

import uvid

import json

import re

from pathlib import Path

from src.Helpers import main

import Api Call

import zipfile

import io

import threading

fron flask import jsonify

Import logging

import logging.config

import getpass

from datetime import datetime

import configparser

from dotenv import Load_dotenv

app Flask (_name__)

app.secret key = '46215442c98b1176996ee4ab24b6b5a1ecf8707cc37f118890b5a51d4e6a4d63'

cwd=os.getcwd()

project_root os.path.dirname(cwd)

input_folder os.path.join(project_root, 'input_folders')

input_folder2 os.path.join(project_root, 'validation_input')

#Save Invoice/PO contents here

#Save the single root-level .xlsx here

Path(input_folder).mkdir(parents=True, exist_ok=True)

Path(input_folder2).mkdir(parents=True, exist_ok=True)

input_folder3 os.path.join(project_root, 'src')

email_file_path = os.path.join(input_folder3, "email_id.txt")

report_dir_path = os.path.join(cwd, 'Output_File', 'Report_Files')

report_dir_path1 = os.path.join(cwd, 'Output_File', 'Data_Files')




Path(report_dir_path).mkdir(parents=True, exist_ok=True)

#pjvg

def init logging():

#Load.env so os.environ contains ENV

    load_dotenv()

    user name getpass.getuser()



    #BASE_LOG_DIR = fr"C:\Users\{user_name}\OneDrive WBA WBS\PSP\Capital_Project"

    one_drive_path = os.path.join(os.path.expanduser("-"), "OneDrive WBA", "WBS", "PSP", "Capital Project")

    BASE_LOG_DIR = one_drive_path
    DEVELOPER_LOG_DIR = os.path.join(BASE_LOG_DIR, "developer")
    USER_LOG_DIR = os.path.join(BASE_LOG_DIR, "user")

    os.makedirs(DEVELOPER_LOG_DIR, exist_ok=True)
    os.makedirs(USER_LOG_DIR, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d-%H-%M-%S")
    developer_log_file = os.path.join(DEVELOPER_LOG_DIR, f"developer_{timestamp}.log")
    user_log_file = os.path.join(USER_LOG_DIR, f"user_{timestamp}.log")

    CONFIG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..\config")
    log_configs = {"dev": "logging.dev.ini", "prod": "logging.prod.ini","uat":"logging.uat.ini"}

    env = os.environ.get("ENV", "dev").lower()
    config = log_configs.get(env, "logging.dev.ini")

    config_path = os.path.join(CONFIG_DIR, config)
    config_parser = configparser.ConfigParser()
    config_parser.read(config_path)

    LOGGING_URL = config_parser.get("log_config", "LOGGING_URL", fallback="")
    os.environ["LOGGING_URL"] = LOGGING_URL
    logging.config.fileConfig(
        config_path,
        disable_existing_loggers=False,
        defaults={
            "logfilename": developer_log_file.replace("\\", "/"),
            "userLogfilename": user_log_file.replace("\\", "/"),
        },
    )

    developer_logger = logging.getLogger("file")
    user_logger = logging.getLogger("user")

    # developer_logger.info(f"Tower Name: PSP")
    # user_logger.info(f"Tower Name: PSP")
    # developer_logger.info(f"Sub Function: Capital Projects")
    # user_logger.info(f"Sub Function: Capital Projects")
    # developer_logger.info(f"Use Case: Automation of invoice pending approval")
    # user_logger.info(f"Use Case: Automation of invoice pending approval")

    return developer_logger, user_logger

developer_logger, user_logger = init_logging()

# === Utilities ===
def empty_directory(dir_path: str):
    """Delete all files and subdirs inside dir_path, but keep dir_path."""
    if not os.path.exists(dir_path):
        return
    for root, dirs, files in os.walk(dir_path, topdown=False):
        for name in files:
            try:
                os.remove(os.path.join(root, name))
            except Exception:
                pass
        for name in dirs:
            try:
                os.rmdir(os.path.join(root, name))
            except OSError:
                pass


def secure_path_components(rel_path: str):
    """
    Split a relative path into secure components, stripping ., .. and unsafe chars.
    Use only for saving, not for detection (which needs human-readable names).
    """
    parts = rel_path.replace("\\", "/").split("/")
    safe_parts = []
    for p in parts:
        if not p or p in (".", ".."):
            continue
        safe_parts.append(secure_filename(p))
    return safe_parts


def safe_join(base: str, rel_path: str):
    """
    Join base with a sanitized relative path (secure_filename per component),
    and ensure the result remains within base (prevents traversal).
    """
    parts = secure_path_components(rel_path)
    clean_rel = os.path.normpath(os.path.join(*parts)) if parts else ""
    clean_rel = clean_rel.lstrip(os.sep)
    candidate = os.path.abspath(os.path.join(base, clean_rel))
    base_abs = os.path.abspath(base)
    if not (candidate == base_abs or candidate.startswith(base_abs + os.sep)):
        raise ValueError("Unsafe path traversal detected")
    return candidate


def norm_seg(s: str) -> str:
    """Lowercase and strip non-alphanumerics to match folder names robustly."""
    return re.sub(r'[^a-z0-9]+', '', s.lower())


# === Routes ===
@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = (request.form.get('email') or '').strip().lower()

        # Clear previous inputs on login to avoid mixing runs
        empty_directory(input_folder)
        empty_directory(input_folder2)

        if email.endswith("@walgreens.com"):
            session['logged_in'] = True
            session['email'] = email
            session['session_id'] = str(uuid.uuid4())
            with open(email_file_path, "w") as f:
                f.write(email)
            return redirect(url_for('upload'))
        else:
            return render_template('login.html', error="Invalid email. Please use your Walgreens Email ID.")

    # Optional: remove this if you don't want to auto-logout when opening login page
    session.clear()
    return render_template('login.html')


@app.route('/upload', methods=['GET'])
def upload():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    download_link = url_for('download_file') if session.get('process_complete') and session.get('output_file') else None
    return render_template(
        'upload.html',
        email=session['email'],
        download_link=download_link,
        error=None
    )

@app.route('/upload_and_run', methods=['POST'])
def upload_and_run():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    session.pop('process_complete', None)
    session.pop('output_file', None)

    empty_directory(input_folder)
    empty_directory(input_folder2)

    files = request.files.getlist('files')
    if not files:
        return jsonify({"error": "Please select a folder to upload."}), 400

    found_invoice_files = False
    found_purchase_order_files = False
    root_xlsx_count = 0

    for file in files:
        rel_path_raw = file.filename
        if not rel_path_raw:
            continue

        raw_parts = rel_path_raw.replace("\\", "/").split("/")
        if len(raw_parts) < 2:
            continue

        norm_parts = [norm_seg(p) for p in raw_parts]

        try:
            idx = next(i for i, p in enumerate(norm_parts) if p in ('invoicefiles', 'purchaseorderfiles'))
            sub_rel = "/".join(raw_parts[idx:])
            dest_path = safe_join(input_folder, sub_rel)
            Path(os.path.dirname(dest_path)).mkdir(parents=True, exist_ok=True)
            file.save(dest_path)

            folder_key = norm_seg(raw_parts[idx])
            if folder_key == 'invoicefiles':
                found_invoice_files = True
            elif folder_key == 'purchaseorderfiles':
                found_purchase_order_files = True
            continue
        except StopIteration:
            pass

        if len(raw_parts) == 2 and raw_parts[-1].lower().endswith(".xlsx"):
            dest_path = safe_join(input_folder2, raw_parts[-1])
            Path(os.path.dirname(dest_path)).mkdir(parents=True, exist_ok=True)
            file.save(dest_path)
            root_xlsx_count += 1
            continue

    errors = []
    if not found_invoice_files:
        errors.append("Missing required subfolder: 'Invoice Files'.")
    if not found_purchase_order_files:
        errors.append("Missing required subfolder: 'Purchase Order Files'.")
    if root_xlsx_count == 0:
        errors.append("Missing root-level Vender Master file.")
    elif root_xlsx_count > 1:
        errors.append("Multiple root-level Vender Master files found.")

    if errors:
        return jsonify({"error": " ".join(errors)}), 400

    return jsonify({"status": "uploaded"}), 200


@app.route('/start_processing', methods=['POST'])
def start_processing():
    if not session.get('logged_in'):
        return jsonify({"error": "Not logged in"}), 403

    # Capture session data before starting the thread
    email_address = session.get('email')
    session_id = session.get('session_id')

    # Use a global or app-level variable to track status
    app.config['PROCESS_COMPLETE'] = False
    app.config['PROCESS_ERROR'] = False

    def background_task(email, session_id):
        try:
            session_id_json = json.dumps({"id": session_id})
            Api_Call.post_log("Capital_Projects", email, "Started", session_id_json)

            main.cap_pro()

            Api_Call.post_log("Capital_Projects", email, "Completed", session_id_json)

            # Find output files
            output_file = max(
                (f for f in os.listdir(report_dir_path) if f.lower().endswith(".xlsx")),
                key=lambda f: os.path.getmtime(os.path.join(report_dir_path, f)),
                default=None
            )
            data_file = max(
                (f for f in os.listdir(report_dir_path1) if f.lower().endswith(".xlsx")),
                key=lambda f: os.path.getmtime(os.path.join(report_dir_path1, f)),
                default=None
            )

            if output_file and data_file:
                app.config['OUTPUT_FILE'] = output_file
                app.config['DATA_FILE'] = data_file
                app.config['PROCESS_COMPLETE'] = True
            else:
                app.config['PROCESS_ERROR'] = True
        except Exception as e:
            developer_logger.error(f"Error during processing: {e}", exc_info=True)
            app.config['PROCESS_ERROR'] = True

    threading.Thread(target=background_task, args=(email_address, session_id)).start()
    return jsonify({"status": "processing started"})

@app.route('/check_status')
def check_status():
    if app.config.get('PROCESS_ERROR'):
        return jsonify({"complete": "error"})
    return jsonify({"complete": app.config.get('PROCESS_COMPLETE', False)})

@app.route('/download')
def download_file():
    if not session.get('logged_in'):
        return redirect(url_for('upload'))

    output_file = app.config.get('OUTPUT_FILE')
    data_file = app.config.get('DATA_FILE')

    if not output_file or not data_file:
        flash("One or more output files not found.", "error")
        return redirect(url_for('upload'))

    file1_path = os.path.join(report_dir_path, output_file)
    file2_path = os.path.join(report_dir_path1, data_file)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        zip_file.write(file1_path, arcname=os.path.basename(file1_path))
        zip_file.write(file2_path, arcname=os.path.basename(file2_path))

    zip_buffer.seek(0)
    return send_file(zip_buffer, as_attachment=True, download_name="Capital_Projects_Files.zip", mimetype='application/zip')



if __name__ == '__main__':
    # Consider setting host/port as per deployment
    CONFIG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "../config")
    log_configs = {"dev": "logging.dev.ini", "prod": "logging.prod.ini", "uat": "logging.uat.ini"}

    env = os.environ.get("ENV", "dev").lower()
    config = log_configs.get(env, "logging.dev.ini")
    config_path = os.path.join(CONFIG_DIR, config)
    config_parser = configparser.ConfigParser()
    config_parser.read(config_path)
    host = config_parser.get("host_name", "Host_name", fallback="127.0.0.1")
    print(host)
    port = config_parser.getint("Port_Number", "port", fallback=5000)
    app.run(host=host, port=port, debug=True)
    # app.run(debug=True)
