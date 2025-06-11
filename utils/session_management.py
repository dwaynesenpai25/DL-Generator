import threading
import uuid
import os
from pathlib import Path
from utils.config import OUTPUT_DIR, logger
from utils.ftp_operations import FTPConnection, download_ftp_template
from utils.config import FTP_CONFIG

# User-specific session management
USER_SESSIONS = {}
SESSION_LOCK = threading.RLock()

def get_fresh_session_state():
    """Create a fresh session state for a user"""
    return {
        'selected_mode': '',
        'selected_folder': '',
        'selected_dl_type': '',
        'selected_template': '',
        'base_template': None,
        'template_path': None,
        'header_footer_template_path': None,
        'transmittal_template_path': None,
        'placeholders': [],
        'download_completed': False,
        'files_to_cleanup': [],
        'zip_path': None,
        'template_combined': False,
        'output_format': 'zip',
        'user_output_dir': None,
        'processing_lock': threading.Lock()
    }

def get_user_session_state(user_email: str):
    """Get or create session state for a specific user"""
    with SESSION_LOCK:
        if user_email not in USER_SESSIONS:
            USER_SESSIONS[user_email] = get_fresh_session_state()
            # Create user-specific output directory
            user_output_dir = OUTPUT_DIR / f"user_{uuid.uuid4().hex[:8]}_{user_email.replace('@', '_').replace('.', '_')}"
            os.makedirs(user_output_dir, exist_ok=True)
            USER_SESSIONS[user_email]['user_output_dir'] = user_output_dir
        return USER_SESSIONS[user_email]

def reset_user_session_state(user_email: str):
    """Reset session state for a specific user"""
    with SESSION_LOCK:
        if user_email in USER_SESSIONS:
            session_state = USER_SESSIONS[user_email]
            # Clean up files
            cleanup_files(session_state.get('files_to_cleanup', []))
            cleanup_user_directory(session_state['user_output_dir'])
            # Reset to fresh state but keep the same output directory
            user_output_dir = session_state.get('user_output_dir')
            USER_SESSIONS[user_email] = get_fresh_session_state()
            USER_SESSIONS[user_email]['user_output_dir'] = user_output_dir
            logger.info(f"Session state reset for user: {user_email}")

def cleanup_files(file_paths):
    """Enhanced cleanup function that handles all types of files"""
    for file_path in file_paths:
        try:
            if file_path and Path(file_path).exists():
                path_obj = Path(file_path)
                if path_obj.is_file():
                    path_obj.unlink()
                    logger.debug(f"Deleted file: {file_path}")
                elif path_obj.is_dir():
                    import shutil
                    shutil.rmtree(path_obj)
                    logger.debug(f"Deleted directory: {file_path}")
        except Exception as e:
            logger.warning(f"Failed to delete {file_path}: {e}")

def cleanup_user_directory(user_dir):
    """Enhanced cleanup function for user directories"""
    try:
        if user_dir and Path(user_dir).exists():
            total_size = 0
            file_count = 0
            
            # Calculate total size and count before deletion
            for item in Path(user_dir).rglob('*'):
                if item.is_file():
                    try:
                        total_size += item.stat().st_size
                        file_count += 1
                    except (OSError, FileNotFoundError):
                        pass
            
            # Delete all contents
            for item in Path(user_dir).iterdir():
                try:
                    if item.is_file():
                        item.unlink()
                        logger.debug(f"Deleted user file: {item}")
                    elif item.is_dir():
                        import shutil
                        shutil.rmtree(item)
                        logger.debug(f"Deleted user directory: {item}")
                except Exception as e:
                    logger.warning(f"Failed to delete user item {item}: {e}")
            
            # Log cleanup summary
            size_mb = total_size / (1024 * 1024)
            logger.info(f"Cleaned up user directory {user_dir}: {file_count} files, {size_mb:.2f} MB freed")
            
    except Exception as e:
        logger.error(f"Error cleaning user directory {user_dir}: {e}")

def set_processing_mode(mode: str, user_email: str):
    """Set processing mode for user"""
    if mode not in ["DL Only", "DL w/ Transmittal", "Transmittal Only"]:
        from fastapi import HTTPException
        raise HTTPException(status_code=400, detail="Invalid mode")
    
    session_state = get_user_session_state(user_email)
    
    # Reset session state before setting new mode to prevent conflicts
    if session_state.get('selected_mode') and session_state.get('selected_mode') != mode:
        logger.info(f"Mode changing from {session_state.get('selected_mode')} to {mode} for user {user_email}, resetting state")
        reset_user_session_state(user_email)
        session_state = get_user_session_state(user_email)
    
    session_state['selected_mode'] = mode
    template_status = {}

    if mode in ["DL Only", "DL w/ Transmittal"]:
        if session_state.get('base_template') and session_state.get('header_footer_template_path') and session_state.get('template_path'):
            template_status['dl_template'] = "DL and header/footer templates are ready"
        else:
            template_status['dl_template'] = "DL or header/footer template not loaded"

    if mode in ["DL w/ Transmittal", "Transmittal Only"]:
        try:
            with FTPConnection(FTP_CONFIG["hostname"], FTP_CONFIG["port"], FTP_CONFIG["username"], FTP_CONFIG["password"]) as ftp_conn:
                ftp = ftp_conn.connect()
                if not ftp:
                    from fastapi import HTTPException
                    raise HTTPException(status_code=500, detail="Failed to connect to FTP server")
                transmittal_template = "Transmittal QRCODE.docx"
                transmittal_template_path = download_ftp_template(ftp, None, transmittal_template, is_transmittal=True)
                if not transmittal_template_path:
                    from fastapi import HTTPException
                    raise HTTPException(status_code=500, detail=f"Failed to download transmittal template {transmittal_template}")
                session_state['files_to_cleanup'].append(transmittal_template_path)
                session_state['transmittal_template_path'] = transmittal_template_path
                template_status['transmittal_template'] = "Transmittal template is ready"
        except Exception as e:
            logger.error(f"Failed to set mode {mode} for user {user_email}: {e}")
            template_status['transmittal_template'] = f"Failed to load transmittal template: {str(e)}"
            from fastapi import HTTPException
            raise HTTPException(status_code=500, detail=f"Failed to set mode: {str(e)}")

    return {
        "success": True,
        "mode": mode,
        "template_status": template_status
    }
