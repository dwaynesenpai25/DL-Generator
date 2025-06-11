import win32print
import win32com.client
import pythoncom
import subprocess
from pathlib import Path
from fastapi import HTTPException
from utils.config import logger
from utils.session_management import get_user_session_state
from docx import Document
from docx.shared import Inches

def get_available_printers():
    """Get list of available printers on the system"""
    try:
        printers = []
        printer_enum = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        for printer in printer_enum:
            printers.append({
                "name": printer[2],  # Printer name
                "is_default": printer[2] == win32print.GetDefaultPrinter()
            })
        return printers
    except Exception as e:
        logger.error(f"Failed to get available printers: {e}")
        return []

def normalize_margins(doc_path):
    doc = Document(doc_path)
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        # section.left_margin = Inches(0.5)
        # section.right_margin = Inches(0.5)
    doc.save(doc_path)

def print_to_specific_printer_docx(docx_files, printer_name):
    """Print DOCX files to a specific printer using Microsoft Word"""
    try:
        pythoncom.CoInitialize()
        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0  # wdAlertsNone
        try:
            for docx_file in docx_files:
                # Normalize margins before opening in Word
                normalize_margins(docx_file)

                doc = word_app.Documents.Open(str(Path(docx_file).absolute()))
                try:
                    word_app.ActivePrinter = printer_name
                    doc.PrintOut(
                        Background=True,
                        Copies=1,
                        Collate=True
                    )
                    logger.info(f"Printed DOCX file: {docx_file} to {printer_name}")
                finally:
                    doc.Close(SaveChanges=False)
        finally:
            word_app.Quit()
            pythoncom.CoUninitialize()
        return True
    except Exception as e:
        logger.error(f"Failed to print DOCX to {printer_name}: {e}")
        return False

def print_to_specific_printer(pdf_files, printer_name):
    """Print PDF files to a specific printer"""
    try:
        for pdf_file in pdf_files:
            # Use SumatraPDF with specific printer
            subprocess.run([
                "SumatraPDF", 
                "-print-to", printer_name, 
                "-silent",
                str(pdf_file)
            ], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        return True
    except Exception as e:
        logger.error(f"Failed to print to {printer_name}: {e}")
        return False

def print_files_for_area(area: str, printer: str, user_email: str):
    """Print files for a specific area"""
    session_state = get_user_session_state(user_email)
    user_output_dir = session_state['user_output_dir']
    
    area_dir = user_output_dir / f"{area}"
    if not area_dir.exists():
        raise HTTPException(status_code=404, detail=f"Directory for area {area} not found")
    
    # Look for DOCX files if output_format is print
    if session_state.get('output_format') == "print":
        docx_files = list(area_dir.glob("*.docx"))
        if not docx_files:
            raise HTTPException(status_code=404, detail=f"No DOCX files found for area {area}")
        
        try:
            if printer:
                # Print to specific printer
                success = print_to_specific_printer_docx(docx_files, printer)
                if not success:
                    raise HTTPException(status_code=500, detail=f"Failed to print to {printer}")
            else:
                # Use default printer
                default_printer = win32print.GetDefaultPrinter()
                success = print_to_specific_printer_docx(docx_files, default_printer)
                if not success:
                    raise HTTPException(status_code=500, detail=f"Failed to print to default printer")
            
            return {"success": True, "message": f"Printed {len(docx_files)} DOCX files for area {area}" + (f" to {printer}" if printer else " to default printer")}
        except Exception as e:
            logger.error(f"Failed to print DOCX files for area {area} for user {user_email}: {e}")
            raise HTTPException(status_code=500, detail=f"Failed to print DOCX files: {str(e)}")
    else:
        # Existing PDF printing logic
        pdf_files = list(area_dir.glob("*.pdf"))
        if not pdf_files:
            raise HTTPException(status_code=404, detail=f"No PDF files found for area {area}")
        
        try:
            if printer:
                # Print to specific printer
                success = print_to_specific_printer(pdf_files, printer)
                if not success:
                    raise HTTPException(status_code=500, detail=f"Failed to print to {printer}")
            else:
                # Use default printer
                for pdf_file in pdf_files:
                    subprocess.run(["SumatraPDF", "-print-to-default", str(pdf_file)], 
                                  check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            
            return {"success": True, "message": f"Printed {len(pdf_files)} PDF files for area {area}" + (f" to {printer}" if printer else " to default printer")}
        except Exception as e:
            logger.error(f"Failed to print PDF files for area {area} for user {user_email}: {e}")
            raise HTTPException(status_code=500, detail=f"Failed to print PDF files: {str(e)}")
