import asyncio
import subprocess # For SumatraPDF
from pathlib import Path
from fastapi import HTTPException
from utils.config import logger
from utils.session_management import get_user_session_state # Sync
from docx import Document # Sync
from docx.shared import Inches # Sync
import os

# Attempt to import win32print and win32com.client, handling ImportErrors
try:
    import win32print
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    logger.warning("win32print or win32com.client not available. Printing features will be limited or disabled.")
    win32print = None
    win32com = None
    pythoncom = None


def get_available_printers(): # Sync
    if not WIN32_AVAILABLE:
        logger.warning("Cannot get printers: win32print not available.")
        return []
    try:
        printers = []
        # EnumPrinters might be slow, consider wrapping if it blocks significantly
        printer_enum = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        default_printer_name = win32print.GetDefaultPrinter()
        for printer_info in printer_enum: # printer_info is a tuple
            printers.append({
                "name": printer_info[2],  # Printer name is at index 2
                "is_default": printer_info[2] == default_printer_name
            })
        return printers
    except Exception as e:
        logger.error(f"Failed to get available printers: {e}")
        return []

def normalize_margins(doc_path_str): # Sync
    doc_path = Path(doc_path_str)
    doc = Document(doc_path) # Sync
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        # section.left_margin = Inches(0.5) # Keep original side margins unless specified
        # section.right_margin = Inches(0.5)
    doc.save(doc_path) # Sync

def print_to_specific_printer_docx(docx_files, printer_name): # Sync due to win32com
    if not WIN32_AVAILABLE:
        logger.error("Cannot print DOCX: win32com.client not available.")
        return False
    
    pythoncom.CoInitialize() # Initialize COM for this thread
    word_app = None
    try:
        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0  # wdAlertsNone

        for docx_file_str in docx_files:
            docx_file_path = Path(docx_file_str).absolute()
            normalize_margins(str(docx_file_path)) # Normalize before opening

            doc = word_app.Documents.Open(str(docx_file_path))
            try:
                word_app.ActivePrinter = printer_name
                doc.PrintOut(Background=True, Copies=1, Collate=True)
                logger.info(f"Sent DOCX file to printer {printer_name}: {docx_file_str}")
            finally:
                doc.Close(SaveChanges=False)
        return True
    except Exception as e:
        logger.error(f"Failed to print DOCX to {printer_name}: {e}", exc_info=True)
        return False
    finally:
        if word_app:
            try:
                word_app.Quit()
            except Exception as e_quit:
                 logger.error(f"Error quitting Word application during printing: {e_quit}")
        pythoncom.CoUninitialize()


async def print_to_specific_printer_pdf_async(pdf_files, printer_name): # Async for subprocess
    """Print PDF files to a specific printer using SumatraPDF asynchronously."""
    try:
        for pdf_file_str in pdf_files:
            pdf_file_path = Path(pdf_file_str)
            cmd = [
                "SumatraPDF.exe", # Ensure SumatraPDF is in PATH or provide full path
                "-print-to", printer_name, 
                "-silent", "-exit-when-done", # Added -exit-when-done
                str(pdf_file_path.absolute())
            ]
            process = await asyncio.create_subprocess_exec(
                *cmd,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            stdout, stderr = await process.communicate()
            if process.returncode != 0:
                logger.error(f"SumatraPDF error printing {pdf_file_str} to {printer_name}: {stderr.decode(errors='ignore')}")
                return False # Stop on first error or collect errors?
            logger.info(f"Sent PDF file to printer {printer_name}: {pdf_file_str}")
        return True
    except FileNotFoundError:
        logger.error("SumatraPDF not found. Please ensure it is installed and in your system's PATH.")
        raise HTTPException(status_code=500, detail="SumatraPDF not found. Printing PDF failed.")
    except Exception as e:
        logger.error(f"Failed to print PDF to {printer_name}: {e}", exc_info=True)
        return False


def print_files_for_area(area: str, printer_name_selected: str, user_email: str): # Sync wrapper
    session_state = get_user_session_state(user_email) # Sync
    user_output_dir = Path(session_state['user_output_dir'])
    
    area_dir = user_output_dir / area
    if not area_dir.exists():
        raise HTTPException(status_code=404, detail=f"Directory for area {area} not found.")
    
    if session_state.get('output_format') == "print":
        docx_files_paths = [str(f) for f in area_dir.glob("*.docx")]
        if not docx_files_paths:
            raise HTTPException(status_code=404, detail=f"No DOCX files found for printing in area {area}.")
        
        target_printer = printer_name_selected
        if not target_printer and WIN32_AVAILABLE: # If no printer selected, use default
            target_printer = win32print.GetDefaultPrinter()
        
        if not target_printer:
             raise HTTPException(status_code=400, detail="No printer selected and no default printer found (or win32print unavailable).")

        success = print_to_specific_printer_docx(docx_files_paths, target_printer) # This is sync
        if not success:
            raise HTTPException(status_code=500, detail=f"Failed to print DOCX files to {target_printer}.")
        return {"success": True, "message": f"Sent {len(docx_files_paths)} DOCX files for area {area} to printer: {target_printer}."}
    
    else: # Assuming PDF printing for other formats (e.g. "zip" if user wants to print from server post-generation)
          # This part of the logic might need review based on actual workflow for "zip" format.
          # Typically, for "zip", user downloads and prints locally.
          # If server-side PDF printing is intended for "zip" too, this is okay.
        pdf_files_paths = [str(f) for f in area_dir.glob("*.pdf")] # Check for merged PDFs or individual ones
        if not pdf_files_paths:
            # Try to find merged PDFs if individual ones are not present
            merged_dl_pdf = area_dir / f"{area}_DL_MERGED.pdf"
            merged_transmittal_pdf = area_dir / f"{area}_TRANSMITTAL_MERGED.pdf"
            if merged_dl_pdf.exists():
                pdf_files_paths.append(str(merged_dl_pdf))
            if merged_transmittal_pdf.exists():
                pdf_files_paths.append(str(merged_transmittal_pdf))

        if not pdf_files_paths:
            raise HTTPException(status_code=404, detail=f"No PDF files found for printing in area {area}.")

        target_printer = printer_name_selected
        if not target_printer and WIN32_AVAILABLE:
             target_printer = win32print.GetDefaultPrinter()
        
        if not target_printer:
            raise HTTPException(status_code=400, detail="No printer selected and no default PDF printer mechanism available.")

        # Run the async PDF printing function. Since this parent function is sync (called by to_thread),
        # we need to run the async part in a new event loop or adapt.
        # For simplicity here, if this function is called via to_thread, direct async call is not straightforward.
        # It's better if the endpoint itself calls the async version.
        # This function `print_files_for_area` should ideally be async if it calls async subprocesses.
        # Let's assume the endpoint will call an async version of this.
        # For now, this will block if called from a sync context.
        # To make it work if called by asyncio.to_thread:
        try:
            loop = asyncio.get_event_loop()
        except RuntimeError: # No event loop in current thread
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        
        success = loop.run_until_complete(print_to_specific_printer_pdf_async(pdf_files_paths, target_printer))

        if not success:
            raise HTTPException(status_code=500, detail=f"Failed to print PDF files to {target_printer}.")
        return {"success": True, "message": f"Sent {len(pdf_files_paths)} PDF files for area {area} to printer: {target_printer}."}
