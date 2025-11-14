import os
import Api_Call


def clear_folder(folder_path):
    """Helper function to clear all files in a folder."""
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)

def log_process_start(email_address, session_id):
    """Log the start of the process."""
    Api_Call.post_log("Capital_Projects", email_address, "Started", session_id)


def log_process_end(email_address, session_id):
    """Log the completion of the process."""
    Api_Call.post_log("Capital_Projects", email_address, "Completed", session_id)


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
