from ftplib import FTP, error_perm
from utils.config import logger
import os
from tempfile import NamedTemporaryFile
from datetime import datetime, timedelta
import asyncio # For asyncio.to_thread

class FTPConnection: # Remains synchronous
    def __init__(self, hostname, port, username, password):
        self.ftp = None
        self.hostname = hostname
        self.port = port
        self.username = username
        self.password = password

    def connect(self):
        try:
            self.ftp = FTP()
            self.ftp.connect(self.hostname, self.port, timeout=30) # Added timeout
            self.ftp.login(self.username, self.password)
            logger.info(f"Connected to FTP server: {self.hostname}")
            return self.ftp
        except Exception as e:
            logger.error(f"Failed to connect to FTP server {self.hostname}: {e}")
            return None

    def close(self):
        if self.ftp:
            try:
                self.ftp.quit()
                logger.info(f"Closed FTP connection to {self.hostname}")
            except Exception as e: # Catch error_perm or other exceptions on quit
                logger.warning(f"Error closing FTP connection to {self.hostname}: {e}")
            self.ftp = None

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

# Functions below are synchronous but will be called with asyncio.to_thread from async endpoints

def get_ftp_folders_list(ftp: FTP): # ftp object is from sync ftplib
    try:
        ftp_path = "/DL AUTOMATION/Template DL V2/Content"
        ftp.cwd(ftp_path)
        items = ftp.nlst()
        folders = []
        current_dir = ftp.pwd() # Store current directory to return to
        for item_name in items:
            try:
                # Check if item is a directory by trying to CWD into it
                ftp.cwd(item_name)  # Try to change to item_name
                folders.append(item_name) # If successful, it's a folder
                ftp.cwd(current_dir)    # Go back to original directory
            except error_perm as e: # If CWD fails, it's likely not a directory or not accessible
                if "550" in str(e): # 550 usually means "File unavailable" or "Not a directory"
                    logger.debug(f"Item '{item_name}' is not a folder or not accessible: {e}")
                else:
                    logger.warning(f"FTP permission error checking item '{item_name}': {e}")
                continue # Skip to next item
            except Exception as e_other:
                logger.warning(f"Unexpected error checking item '{item_name}' as folder: {e_other}")
                continue
        return sorted(folders)
    except error_perm as e:
        logger.error(f"FTP permission error retrieving folders from '{ftp_path}': {e}")
        return []
    except Exception as e:
        logger.error(f"Failed to retrieve folders from FTP path '{ftp_path}': {e}")
        return []


def get_ftp_templates_list(ftp: FTP, folder_name: str):
    try:
        ftp_path = f"/DL AUTOMATION/Template DL V2/Content/{folder_name}"
        ftp.cwd(ftp_path)
        templates = [item for item in ftp.nlst() if item.lower().endswith('.docx')]
        return sorted(templates)
    except error_perm as e:
        logger.error(f"FTP permission error retrieving templates from '{ftp_path}': {e}")
        return []
    except Exception as e:
        logger.error(f"Failed to retrieve templates from FTP folder '{folder_name}': {e}")
        return []

def download_ftp_template(ftp: FTP, folder_name: str, template_name: str, is_header_footer=False, is_transmittal=False):
    try:
        if is_transmittal:
            ftp_path = "/DL AUTOMATION/Template Transmittal V2"
        elif is_header_footer:
            ftp_path = "/DL AUTOMATION/Template DL V2/Letter Head"
        else:
            if not folder_name: # Ensure folder_name is provided for content templates
                 logger.error("Folder name is required for content templates.")
                 return None
            ftp_path = f"/DL AUTOMATION/Template DL V2/Content/{folder_name}"
        
        ftp.cwd(ftp_path)
        # Use NamedTemporaryFile to ensure it's cleaned up if process crashes
        # and to get a unique filename.
        with NamedTemporaryFile(delete=False, suffix=".docx", prefix=f"{template_name.split('.')[0]}_") as tmp:
            tmp_path = tmp.name # Get path before closing on some OS, or before retrbinary
        
        with open(tmp_path, 'wb') as local_file: # Open in binary write mode
            ftp.retrbinary(f"RETR {template_name}", local_file.write)
            
        if os.path.exists(tmp_path) and os.path.getsize(tmp_path) > 0: # Check if file exists and is not empty
            logger.info(f"Downloaded template '{template_name}' from '{ftp_path}' to: {tmp_path}")
            return tmp_path
        else:
            logger.error(f"Failed to download template '{template_name}' or downloaded file is empty. Path: {tmp_path}")
            if os.path.exists(tmp_path): # Clean up empty file
                os.unlink(tmp_path)
            return None
            
    except error_perm as e:
        logger.error(f"FTP permission error downloading template '{template_name}' from '{ftp_path}': {e}")
        return None
    except Exception as e:
        logger.error(f"Failed to download template '{template_name}' from '{ftp_path}': {e}")
        # Clean up temp file if it exists and an error occurred
        if 'tmp_path' in locals() and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except OSError:
                pass # Ignore if already deleted or other issue
        return None


def fetch_signature_from_ftp(ftp: FTP):
    # Try today and yesterday's date
    # Hardcoded path for testing: "field/DL/ATTY SIGNATURE/05-29-2025"
    # Original logic:
    # for days_ago in [0, 1]:
    #     date_str = (datetime.today() - timedelta(days=days_ago)).strftime('%m-%d-%Y')
    #     ftp_path = f"field/DL/ATTY SIGNATURE/{date_str}"
    
    # Using the hardcoded path from the original file for now
    # TODO: Revert to dynamic date logic or make configurable if needed
    # fixed_date_str = "05-29-2025" # As per original code's effective path
    # ftp_path = f"field/DL/ATTY SIGNATURE/{fixed_date_str}"
    # Get the current date
    current_date_str = datetime.now().strftime("%m-%d-%Y")  # Format as MM-DD-YYYY

    # Construct the dynamic FTP path
    ftp_path = f"field/DL/ATTY SIGNATURE/{current_date_str}"
    signature_filename = "attySignature.PNG"

    try:
        ftp.cwd(ftp_path)
        with NamedTemporaryFile(delete=False, suffix=".png", prefix="signature_") as tmp:
            tmp_path = tmp.name
        
        with open(tmp_path, 'wb') as local_file:
            ftp.retrbinary(f"RETR {signature_filename}", local_file.write)

        if os.path.exists(tmp_path) and os.path.getsize(tmp_path) > 0:
            logger.info(f"Fetched signature '{signature_filename}' from '{ftp_path}' to {tmp_path}")
            return tmp_path
        else:
            logger.warning(f"Failed to fetch signature '{signature_filename}' from '{ftp_path}' or file is empty.")
            if os.path.exists(tmp_path): os.unlink(tmp_path)
            return None
            
    except error_perm as e:
        logger.warning(f"FTP permission error fetching signature from '{ftp_path}/{signature_filename}': {e}")
        return None
    except Exception as e:
        logger.error(f"Failed to fetch signature from FTP path '{ftp_path}/{signature_filename}': {e}")
        if 'tmp_path' in locals() and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
        return None
    # If loop was used, the final error log would be outside the loop
    # logger.error("Failed to fetch signature from FTP for relevant dates.")
    # return None
