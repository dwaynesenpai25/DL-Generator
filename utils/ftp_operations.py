from ftplib import FTP
from utils.config import logger
import os
from tempfile import NamedTemporaryFile
from datetime import datetime, timedelta

class FTPConnection:
    def __init__(self, hostname, port, username, password):
        self.ftp = None
        self.hostname = hostname
        self.port = port
        self.username = username
        self.password = password

    def connect(self):
        try:
            self.ftp = FTP()
            self.ftp.connect(self.hostname, self.port)
            self.ftp.login(self.username, self.password)
            logger.info("Connected to FTP server")
            return self.ftp
        except Exception as e:
            logger.error(f"Failed to connect to FTP: {e}")
            return None

    def close(self):
        if self.ftp:
            try:
                self.ftp.quit()
                logger.info("Closed FTP connection")
            except Exception as e:
                logger.warning(f"Error closing FTP connection: {e}")
            self.ftp = None

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

def get_ftp_folders_list(ftp):
    try:
        ftp_path = "/DL AUTOMATION/Template DL V2/Content"
        ftp.cwd(ftp_path)
        items = ftp.nlst()
        folders = []
        current_dir = ftp.pwd()
        for item in items:
            try:
                ftp.cwd(item)
                folders.append(item)
                ftp.cwd(current_dir)
            except:
                continue
        return sorted(folders)
    except Exception as e:
        logger.error(f"Failed to retrieve folders from FTP: {e}")
        return []

def get_ftp_templates_list(ftp, folder_name):
    try:
        ftp_path = f"/DL AUTOMATION/Template DL V2/Content/{folder_name}"
        ftp.cwd(ftp_path)
        templates = [item for item in ftp.nlst() if item.lower().endswith('.docx')]
        return sorted(templates)
    except Exception as e:
        logger.error(f"Failed to retrieve templates from FTP folder {folder_name}: {e}")
        return []

def download_ftp_template(ftp, folder_name, template_name, is_header_footer=False, is_transmittal=False):
    try:
        if is_transmittal:
            ftp_path = "/DL AUTOMATION/Template Transmittal V2"
        elif is_header_footer:
            ftp_path = "/DL AUTOMATION/Template DL V2/Letter Head"
        else:
            ftp_path = f"/DL AUTOMATION/Template DL V2/Content/{folder_name}"
        ftp.cwd(ftp_path)
        with NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            ftp.retrbinary(f"RETR {template_name}", tmp.write)
            tmp_path = tmp.name
        if os.path.exists(tmp_path):
            logger.info(f"Downloaded template {template_name} to: {tmp_path}")
            return tmp_path
        raise Exception(f"Failed to download template {template_name}")
    except Exception as e:
        logger.error(f"Failed to download template {template_name}: {e}")
        return None

def fetch_signature_from_ftp(ftp):
    # Try today and yesterday's date
    for days_ago in [0, 1]:
        date_str = (datetime.today() - timedelta(days=days_ago)).strftime('%m-%d-%Y')
        # ftp_path = f"field/DL/ATTY SIGNATURE/{date_str}"
        ftp_path = f"field/DL/ATTY SIGNATURE/05-29-2025"
        try:
            ftp.cwd(ftp_path)
            with NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                ftp.retrbinary("RETR attySignature.PNG", tmp.write)
                tmp_path = tmp.name
            if os.path.exists(tmp_path):
                return tmp_path
        except Exception as e:
            logger.warning(f"Failed to fetch signature for {date_str}: {e}")
    logger.error("Failed to fetch signature from FTP for today or yesterday.")
    return None
