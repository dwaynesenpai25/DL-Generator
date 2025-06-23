from fastapi import HTTPException
from utils.ftp_operations import FTPConnection, get_ftp_folders_list, get_ftp_templates_list, download_ftp_template, fetch_signature_from_ftp # These are sync
from utils.config import FTP_CONFIG, SERVICE_ACCOUNT_JSON, SPREADSHEET_ID, SHEET_NAME, logger
from utils.session_management import get_user_session_state # Sync
from utils.document_utils import get_sheet_data, combine_templates, extract_placeholders, extract_transmittal_placeholders # Sync, some use win32com
import asyncio # For asyncio.to_thread
import os

# Attempt to import win32com.client and pythoncom, handling ImportErrors for non-Windows environments
try:
    import win32com.client
    import pythoncom
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    logger.warning("win32com.client or pythoncom not available. DOCX combining/manipulation features will be disabled.")
    win32com = None
    pythoncom = None


def get_ftp_folders(user_clients, user_access): # Sync
    if not all([FTP_CONFIG["hostname"], FTP_CONFIG["username"], FTP_CONFIG["password"]]):
        raise HTTPException(status_code=500, detail="FTP configuration incomplete")
    
    with FTPConnection(FTP_CONFIG["hostname"], FTP_CONFIG["port"], FTP_CONFIG["username"], FTP_CONFIG["password"]) as ftp_conn:
        ftp = ftp_conn.connect() # connect is sync
        if not ftp:
            raise HTTPException(status_code=500, detail="Failed to connect to FTP server")
        
        all_folders = get_ftp_folders_list(ftp) # sync
        
        if user_access == "admin1" or user_access == "modal": # "modal" seems to be used as an admin-like context for fetching all folders
            return all_folders
        
        if user_clients: # user_clients should be a list
            available_folders = [folder for folder in all_folders if folder in user_clients]
            return available_folders
        
        return []

def get_dl_types_for_folder(folder: str, user_clients, user_access): # Sync
    if not folder:
        raise HTTPException(status_code=400, detail="Folder not specified")
    
    if user_access != "admin":
        if not user_clients or folder not in user_clients:
            raise HTTPException(status_code=403, detail="Access denied to this template folder")
    
    sheet_df = get_sheet_data(SERVICE_ACCOUNT_JSON, SPREADSHEET_ID, SHEET_NAME) # sync (Google API call)
    if sheet_df.empty:
        raise HTTPException(status_code=500, detail="Failed to retrieve Google Sheets data or sheet is empty")
    
    # Filter for the campaign (folder) and get unique DL types
    folder_data = sheet_df[sheet_df["CAMPAIGN"] == folder]
    if folder_data.empty:
        logger.warning(f"No DL types found for folder '{folder}' in Google Sheets.")
        return [] # Return empty list if folder not found in sheet
        
    dl_types = sorted(folder_data["DL TYPE"].dropna().unique().tolist())
    return dl_types

def get_ftp_templates(folder: str, user_clients, user_access): # Sync
    if not folder:
        raise HTTPException(status_code=400, detail="Folder not specified")
    
    if user_access != "admin":
        if not user_clients or folder not in user_clients:
            raise HTTPException(status_code=403, detail="Access denied to this template folder")
    
    with FTPConnection(FTP_CONFIG["hostname"], FTP_CONFIG["port"], FTP_CONFIG["username"], FTP_CONFIG["password"]) as ftp_conn:
        ftp = ftp_conn.connect() # sync
        if not ftp:
            raise HTTPException(status_code=500, detail="Failed to connect to FTP server")
        templates = get_ftp_templates_list(ftp, folder) # sync
        return {"message": "Content templates retrieved successfully", "templates": templates}


def get_placeholders_for_template(request_data, user_email: str, user_clients, user_access): # Sync due to win32com
    if not WIN32COM_AVAILABLE:
        raise HTTPException(status_code=501, detail="Word document processing (win32com) is not available on this server.")

    if not all([request_data.folder, request_data.dl_type, request_data.template]):
        raise HTTPException(status_code=400, detail="Missing parameters: folder, dl_type, or template.")
    
    if user_access != "admin":
        if not user_clients or request_data.folder not in user_clients:
            raise HTTPException(status_code=403, detail="Access denied to this template folder.")
    
    session_state = get_user_session_state(user_email) # sync
    session_state['selected_folder'] = request_data.folder
    session_state['selected_dl_type'] = request_data.dl_type
    
    # Initialize COM for the current thread if it's going to be run in a thread by asyncio.to_thread
    pythoncom.CoInitialize()
    word_app = None
    try:
        with FTPConnection(FTP_CONFIG["hostname"], FTP_CONFIG["port"], FTP_CONFIG["username"], FTP_CONFIG["password"]) as ftp_conn:
            ftp = ftp_conn.connect() # sync
            if not ftp:
                raise HTTPException(status_code=500, detail="Failed to connect to FTP server for placeholders.")
            
            signature_img_path = fetch_signature_from_ftp(ftp) # sync
            if not signature_img_path:
                # Allow processing without signature, but log a warning.
                logger.warning("Signature image not found or failed to download. Proceeding without signature.")
                # raise HTTPException(status_code=500, detail="Failed to fetch signature. Folder or file might not exist.")
            if signature_img_path: # Only add to cleanup if path exists
                 session_state['files_to_cleanup'].append(signature_img_path)

            template_path = download_ftp_template(ftp, request_data.folder, request_data.template, is_header_footer=False) # sync
            if not template_path:
                raise HTTPException(status_code=500, detail=f"Failed to download content template: {request_data.template}")
            session_state['files_to_cleanup'].append(template_path)
            session_state['template_path'] = template_path
            
            sheet_df = get_sheet_data(SERVICE_ACCOUNT_JSON, SPREADSHEET_ID, SHEET_NAME) # sync
            matching_row = sheet_df[(sheet_df["CAMPAIGN"] == request_data.folder) & (sheet_df["DL TYPE"] == request_data.dl_type)]
            if matching_row.empty:
                raise HTTPException(status_code=404, detail=f"No header/footer template mapping found for {request_data.folder}/{request_data.dl_type} in Google Sheets.")
            
            header_footer_filename = matching_row["FILE"].iloc[0]
            if not header_footer_filename.lower().endswith('.docx'):
                header_footer_filename += '.docx'
            
            header_footer_template_path = download_ftp_template(ftp, None, header_footer_filename, is_header_footer=True) # sync
            if not header_footer_template_path:
                raise HTTPException(status_code=500, detail=f"Failed to download header/footer template: {header_footer_filename}")
            session_state['files_to_cleanup'].append(header_footer_template_path)
            session_state['header_footer_template_path'] = header_footer_template_path

        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False # wdAlertsNone equivalent

        base_template = combine_templates(header_footer_template_path, template_path, signature_img_path, word_app) # sync
        if not base_template:
            raise HTTPException(status_code=500, detail="Failed to combine templates using win32com.")
        
        session_state['base_template'] = base_template # Store the python-docx Document object
        placeholders = extract_placeholders(base_template) # sync
        session_state['placeholders'] = placeholders
        session_state['template_combined'] = True # Mark as combined

        return {
            "message": "Final template retrieved and placeholders extracted successfully.",
            "placeholders": placeholders,
            "template_combined": True
        }
    except HTTPException: # Re-raise HTTPExceptions directly
        raise
    except Exception as e:
        logger.error(f"Error in get_placeholders_for_template: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"An error occurred while processing templates: {str(e)}")
    finally:
        if word_app:
            try:
                word_app.Quit()
            except Exception as e_quit:
                logger.error(f"Error quitting Word application: {e_quit}")
        pythoncom.CoUninitialize()
        # Consider more targeted process killing if WINWORD.EXE hangs
        # os.system("taskkill /IM WINWORD.EXE /F >nul 2>&1") # This is a bit aggressive


def get_transmittal_placeholders(user_email: str, user_clients, user_access, folder: str = None): # Sync
    session_state = get_user_session_state(user_email) # sync
    
    if folder:
        if user_access != "admin":
            if not user_clients or folder not in user_clients:
                raise HTTPException(status_code=403, detail="Access denied to this template folder for transmittal.")
        session_state['selected_folder'] = folder
        session_state['selected_dl_type'] = "Transmittal Only" 
    
    transmittal_template_path = session_state.get('transmittal_template_path')
    if not transmittal_template_path:
        # Attempt to load it if mode was set to include transmittal but path is missing (e.g., after a reset)
        # This might require re-calling set_processing_mode or a dedicated load function.
        # For now, strict check:
        raise HTTPException(status_code=404, detail="Transmittal template not loaded in session. Please ensure 'DL w/ Transmittal' or 'Transmittal Only' mode was set successfully.")
    
    try:
        placeholders = extract_transmittal_placeholders(transmittal_template_path) # sync
        session_state['transmittal_placeholders'] = placeholders
        
        return {
            "message": "Transmittal template placeholders retrieved successfully.",
            "placeholders": placeholders,
            "template_type": "transmittal",
            "folder": folder if folder else session_state.get('selected_folder', 'N/A')
        }
    except Exception as e:
        logger.error(f"Error extracting transmittal placeholders: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Failed to extract transmittal placeholders: {str(e)}")
