from fastapi import HTTPException
from utils.ftp_operations import FTPConnection, get_ftp_folders_list, get_ftp_templates_list, download_ftp_template, fetch_signature_from_ftp
from utils.config import FTP_CONFIG, SERVICE_ACCOUNT_JSON, SPREADSHEET_ID, SHEET_NAME, logger
from utils.session_management import get_user_session_state
from utils.document_utils import get_sheet_data, combine_templates, extract_placeholders, extract_transmittal_placeholders
import win32com.client
import pythoncom
import os

def get_ftp_folders(user_clients, user_access):
    """Get FTP folders based on user access"""
    if not all([FTP_CONFIG["hostname"], FTP_CONFIG["username"], FTP_CONFIG["password"]]):
        raise HTTPException(status_code=500, detail="FTP configuration incomplete")
    
    with FTPConnection(FTP_CONFIG["hostname"], FTP_CONFIG["port"], FTP_CONFIG["username"], FTP_CONFIG["password"]) as ftp_conn:
        ftp = ftp_conn.connect()
        if not ftp:
            raise HTTPException(status_code=500, detail="Failed to connect to FTP")
        
        all_folders = get_ftp_folders_list(ftp)
        
        # # If user is admin, return all folders
        if user_access == "modal":
            return all_folders
        
        # If user has assigned clients, return only those folders
        if user_clients:
            available_folders = [folder for folder in all_folders if folder in user_clients]
            return available_folders
        
        # If user has no assigned clients, return empty
        return []

def get_dl_types_for_folder(folder: str, user_clients, user_access):
    """Get DL types for a specific folder"""
    if not folder:
        raise HTTPException(status_code=400, detail="Folder not specified")
    
    # Check if user has access to this folder
    if user_access != "admin":  # Not admin
        if not user_clients or folder not in user_clients:
            raise HTTPException(status_code=403, detail="Access denied to this template folder")
    
    sheet_df = get_sheet_data(SERVICE_ACCOUNT_JSON, SPREADSHEET_ID, SHEET_NAME)
    if sheet_df.empty:
        raise HTTPException(status_code=500, detail="Failed to retrieve Google Sheets data")
    dl_types = sorted(sheet_df[sheet_df["CAMPAIGN"] == folder]["DL TYPE"].dropna().unique().tolist())
    return dl_types

def get_ftp_templates(folder: str, user_clients, user_access):
    """Get templates for a specific folder"""
    if not folder:
        raise HTTPException(status_code=400, detail="Folder not specified")
    
    # Check if user has access to this folder
    if user_access != "admin":  # Not admin
        if not user_clients or folder not in user_clients:
            raise HTTPException(status_code=403, detail="Access denied to this template folder")
    
    with FTPConnection(FTP_CONFIG["hostname"], FTP_CONFIG["port"], FTP_CONFIG["username"], FTP_CONFIG["password"]) as ftp_conn:
        ftp = ftp_conn.connect()
        if not ftp:
            raise HTTPException(status_code=500, detail="Failed to connect to FTP")
        templates = get_ftp_templates_list(ftp, folder)
        return {"message": "Content/Letter Head retrieved successfully", "templates": templates}

def get_placeholders_for_template(request, user_email: str, user_clients, user_access):
    """Get placeholders for a specific template"""
    if not all([request.folder, request.dl_type, request.template]):
        raise HTTPException(status_code=400, detail="Missing parameters")
    
    # Check if user has access to this folder
    if user_access != "admin":  # Not admin
        if not user_clients or request.folder not in user_clients:
            raise HTTPException(status_code=403, detail="Access denied to this template folder")
    
    # Get user-specific session state
    session_state = get_user_session_state(user_email)
    
    # Store selected folder and dl_type for audit trail
    session_state['selected_folder'] = request.folder
    session_state['selected_dl_type'] = request.dl_type
    
    with FTPConnection(FTP_CONFIG["hostname"], FTP_CONFIG["port"], FTP_CONFIG["username"], FTP_CONFIG["password"]) as ftp_conn:
        ftp = ftp_conn.connect()
        if not ftp:
            raise HTTPException(status_code=500, detail="Failed to connect to FTP")
        signature_img_path = fetch_signature_from_ftp(ftp)
        if not signature_img_path:
            raise HTTPException(status_code=500, detail="Failed to fetch signature from file server. Folder or file might not be created yet.")
        session_state['files_to_cleanup'].append(signature_img_path)
        template_path = download_ftp_template(ftp, request.folder, request.template, is_header_footer=False)
        if not template_path:
            raise HTTPException(status_code=500, detail="Failed to download content template")
        session_state['files_to_cleanup'].append(template_path)
        session_state['template_path'] = template_path
        sheet_df = get_sheet_data(SERVICE_ACCOUNT_JSON, SPREADSHEET_ID, SHEET_NAME)
        matching_row = sheet_df[(sheet_df["CAMPAIGN"] == request.folder) & (sheet_df["DL TYPE"] == request.dl_type)]
        if matching_row.empty:
            raise HTTPException(status_code=404, detail=f"No header/footer template for {request.folder}/{request.dl_type}")
        header_footer_filename = matching_row["FILE"].iloc[0]
        if not header_footer_filename.lower().endswith('.docx'):
            header_footer_filename += '.docx'
        header_footer_template_path = download_ftp_template(ftp, None, header_footer_filename, is_header_footer=True)
        if not header_footer_template_path:
            raise HTTPException(status_code=500, detail=f"Failed to download header/footer template {header_footer_filename}")
        session_state['files_to_cleanup'].append(header_footer_template_path)
        session_state['header_footer_template_path'] = header_footer_template_path
        pythoncom.CoInitialize()
        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False
        try:
            base_template = combine_templates(header_footer_template_path, template_path, signature_img_path, word_app)
            if not base_template:
                raise HTTPException(status_code=500, detail="Failed to combine templates")
            session_state['base_template'] = base_template
            placeholders = extract_placeholders(base_template)
            session_state['placeholders'] = placeholders
            
            # Check if template is already combined
            template_combined_message = ""
            if session_state.get('template_combined', False):
                template_combined_message = " (Template already)"
            
            return {
                "message": f"Final template retrieved successfully{template_combined_message}", 
                "placeholders": placeholders,
                "template_combined": session_state.get('template_combined', False)
            }
        finally:
            word_app.Quit()
            pythoncom.CoUninitialize()
            os.system("taskkill /IM WINWORD.EXE /F >nul 2>&1")

def get_transmittal_placeholders(user_email: str, user_clients, user_access, folder: str = None):
    """Get placeholders for transmittal template only"""
    # Get user-specific session state
    session_state = get_user_session_state(user_email)
    
    # If folder is provided, store it for audit trail (for Transmittal Only mode)
    if folder:
        # Check if user has access to this folder
        if user_access != "admin":  # Not admin
            if not user_clients or folder not in user_clients:
                raise HTTPException(status_code=403, detail="Access denied to this template folder")
        
        # Store selected folder for audit trail
        session_state['selected_folder'] = folder
        session_state['selected_dl_type'] = "Transmittal Only"  # Set a default DL type for transmittal only
    
    # Get the transmittal template from session state
    transmittal_template_path = session_state.get('transmittal_template_path')
    
    if not transmittal_template_path:
        raise HTTPException(status_code=404, detail="No transmittal template found. Please set mode first.")
    
    try:
        # Extract placeholders from transmittal template using the specialized function
        # Pass the file path directly to the function
        placeholders = extract_transmittal_placeholders(transmittal_template_path)
        
        session_state['transmittal_placeholders'] = placeholders
        
        return {
            "message": "Transmittal template placeholders retrieved successfully",
            "placeholders": placeholders,
            "template_type": "transmittal",
            "folder": folder if folder else session_state.get('selected_folder')
        }
    except Exception as e:
        logger.error(f"Error extracting transmittal placeholders: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to extract transmittal placeholders: {str(e)}")