import pandas as pd
import asyncio
import json
import time
import uuid
import shutil
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from fastapi import HTTPException
from fastapi.responses import FileResponse
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches
# from PyPDF2 import PdfMerger # PyPDF2 is sync, consider alternatives or run in thread if it becomes a bottleneck
from utils.config import logger
from utils.session_management import get_user_session_state # This is sync
from utils.document_utils import generate_barcode, generate_qrcode, amount_to_words, combine_templates, extract_placeholders # Some of these might be sync
from utils.pdf_conversion import batch_convert_libreoffice # Now async
from utils.database import generate_doc_code, add_audit_entry, add_processed_accounts # Now async
import os
import re

# PyPDF2 is synchronous. For async PDF merging, you might need to run it in a thread
# or find an async-native PDF library if this becomes a bottleneck.
# For now, we'll wrap its usage.
from PyPDF2 import PdfMerger 

def get_raw_file(file): # Sync
    try:
        contents = file.file.read()
        df = pd.read_excel(BytesIO(contents), dtype=str)
        return df
    except Exception as e:
        logger.error(f"Error reading file: {e}")
        return pd.DataFrame([])

def process_excel_file(file): # Sync
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="Invalid file format. Please upload an .xlsx file")
    df = get_raw_file(file)
    if df.empty:
        raise HTTPException(status_code=500, detail="Failed to read Excel file or file is empty.")
    
    # Basic validation for required columns
    required_cols = ['FINAL_AREA', 'DL_CODE', 'LEADS_CHNAME', 'DL_ADDRESS', 'LEADS_NEW_OB'] # Added LEADS_NEW_OB as it's used
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        error_detail = f"Excel file must contain the following columns: {', '.join(missing_cols)}"
        logger.error(error_detail)
        raise HTTPException(status_code=400, detail=error_detail)
        
    return {"data": df.to_dict(orient='records')}


def fill_template(doc, mapping, barcode_buffer=None): # Sync
    def replace_in_text(text):
        for k, v in mapping.items():
            if k == "«IMAGE_BARCODE»" and v: # Check if v (barcode_buffer) is not None or empty
                continue
            text = text.replace(k, str(v) if v is not None else "") # Replace None with empty string
        return text

    for para in doc.paragraphs:
        for run in para.runs:
            if run.text:
                run.text = replace_in_text(run.text)

    for section in doc.sections:
        for header_footer_item in [section.header, section.first_page_header, section.even_page_header, 
                                   section.footer, section.first_page_footer, section.even_page_footer]: # Combined loop
            if header_footer_item:
                for para in header_footer_item.paragraphs:
                    for run in para.runs:
                        if run.text:
                            run.text = replace_in_text(run.text)
                for table in header_footer_item.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                if "«IMAGE_BARCODE»" in para.text and barcode_buffer:
                                    para.clear()
                                    # Ensure cell.paragraphs[0] exists or handle appropriately
                                    if cell.paragraphs:
                                        target_para = cell.paragraphs[0]
                                        target_para.paragraph_format.left_indent = Pt(-20)
                                        run_pic = target_para.add_run()
                                        try:
                                            run_pic.add_picture(barcode_buffer, width=Inches(3.0), height=Inches(0.35))
                                        except Exception as e:
                                            logger.error(f"Error adding barcode picture: {e}")
                                    else: # Create a new paragraph if none exist
                                        new_para = cell.add_paragraph()
                                        new_para.paragraph_format.left_indent = Pt(-20)
                                        run_pic = new_para.add_run()
                                        try:
                                            run_pic.add_picture(barcode_buffer, width=Inches(3.0), height=Inches(0.35))
                                        except Exception as e:
                                            logger.error(f"Error adding barcode picture to new paragraph: {e}")
                                else:
                                    for run in para.runs:
                                        if run.text:
                                            run.text = replace_in_text(run.text)
    return doc

def clear_placeholders(inner_table): # Sync
    def replace_in_text(text):
        return re.sub(r'«[^»]+»', "", text)
    for row in inner_table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    if run.text:
                        run.text = replace_in_text(run.text)

def fill_inner_table(inner_table, mapping, qrcode_buffer=None): # Sync
    def replace_in_text(text):
        for k, v in mapping.items():
            if k == "«IMAGE_QRCODE»":
                continue
            text = text.replace(k, str(v) if v is not None else "")
        return text
    for row in inner_table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                if "«IMAGE_QRCODE»" in para.text and qrcode_buffer:
                    para.clear()
                    para.paragraph_format.left_indent = Pt(-5)
                    run = para.add_run()
                    try:
                        run.add_picture(qrcode_buffer, width=Inches(1), height=Inches(1))
                    except Exception as e:
                        logger.error(f"Error adding QR code picture: {e}")
                else:
                    for run in para.runs:
                        if run.text:
                            run.text = replace_in_text(run.text)

def fill_transmittal_template(template_path, group_df): # Sync
    try:
        temp_doc_orig = Document(template_path) # Load original once
        if not temp_doc_orig.tables:
            logger.error("No tables found in the transmittal template.")
            return None
        
        total_records_in_group = len(group_df)
        # Each page of transmittal has 4 slots.
        total_pages_needed = (total_records_in_group + 3) // 4 # Equivalent to ceil(total_records_in_group / 4)
        
        filled_transmittal_docs = []

        for page_num in range(total_pages_needed):
            # Create a new document for each page from the original template
            current_page_doc = Document(template_path) # Or deepcopy temp_doc_orig if python-docx supports it well
            current_page_table = current_page_doc.tables[0] # Assuming the structure is consistent

            # Get data for the current page (up to 4 records)
            start_idx = page_num * 4
            end_idx = start_idx + 4
            page_records_df = group_df.iloc[start_idx:end_idx]

            for table_row_idx, table_row_obj in enumerate(current_page_table.rows):
                if table_row_idx < len(page_records_df):
                    # This slot in the table corresponds to a record
                    record_data = page_records_df.iloc[table_row_idx]
                    mapping = {f"«{col.upper()}»": str(record_data[col]) for col in group_df.columns if pd.notnull(record_data[col])}
                    
                    qrcode_buffer = None
                    if dl_code := record_data.get('DL_CODE', ''):
                        qrcode_buffer = generate_qrcode(dl_code) # generate_qrcode is sync
                    
                    for cell in table_row_obj.cells:
                        if cell.tables: # Expecting one inner table per cell
                            inner_table = cell.tables[0]
                            fill_inner_table(inner_table, mapping, qrcode_buffer) # fill_inner_table is sync
                else:
                    # This slot in the table is empty, clear placeholders
                    for cell in table_row_obj.cells:
                        if cell.tables:
                            inner_table = cell.tables[0]
                            clear_placeholders(inner_table) # clear_placeholders is sync
            
            filled_transmittal_docs.append(current_page_doc)
            
        return filled_transmittal_docs
    except Exception as e:
        logger.error(f"Failed to fill transmittal template: {e}", exc_info=True)
        return None


def format_time_duration(seconds): # Sync
    """Convert seconds to human-readable format"""
    if seconds < 60:
        return f"{seconds:.1f} seconds"
    elif seconds < 3600:
        minutes = seconds / 60
        return f"{minutes:.1f} minutes"
    else:
        hours = seconds / 3600
        return f"{hours:.1f} hours"

async def generate_pdfs_stream(uploaded_file, dataframe: pd.DataFrame, user_info: dict): # Async generator
    user_email = user_info.get("email", "")
    # get_user_session_state is sync, run in thread
    session_state = await asyncio.to_thread(get_user_session_state, user_email)
    
    if not await asyncio.to_thread(session_state['processing_lock'].acquire, blocking=False):
        yield json.dumps({'error': 'Another processing task is already running for this user.'}) + '\n'
        return
    
    zipf = None
    try:
        if not session_state.get('selected_mode'):
            yield json.dumps({'error': 'No processing mode selected. Please select a mode first.'}) + '\n'
            return

        if session_state.get('excel_errors'): # Assuming excel_errors is populated elsewhere if needed
            error_msg = f"Excel validation errors found: {'; '.join(session_state['excel_errors'])}"
            yield json.dumps({'error': error_msg}) + '\n'
            return

        if session_state.get('signature_error'): # Assuming signature_error is populated elsewhere
            yield json.dumps({'error': session_state['signature_error']}) + '\n'
            return

        if not session_state.get('base_template') and session_state.get('selected_mode') in ["DL Only", "DL w/ Transmittal"]:
            yield json.dumps({'error': 'DL template (base_template) not loaded. Please select folder, DL type, and content file.'}) + '\n'
            return
        if not session_state.get('transmittal_template_path') and session_state.get('selected_mode') in ["DL w/ Transmittal", "Transmittal Only"]:
            yield json.dumps({'error': 'Transmittal template not loaded. Ensure mode selection was successful.'}) + '\n'
            return

        logger.debug(f"Excel file loaded with {len(dataframe)} rows for user {user_email}")
        from datetime import datetime
        today_date = datetime.now().strftime("%B %d, %Y")
        
        # Data transformations (ensure columns exist before applying)
        if 'DL_ADDRESS' in dataframe.columns:
            dataframe['DL_ADDRESS'] = dataframe['DL_ADDRESS'].astype(str).str.upper()
        if 'LEADS_NEW_OB' in dataframe.columns:
            dataframe['LEADS_NEW_OB'] = dataframe['LEADS_NEW_OB'].apply(lambda x: f"{float(x):,.2f}" if pd.notnull(x) and str(x).replace('.', '', 1).isdigit() else str(x))
        else: # If LEADS_NEW_OB is critical and missing
            yield json.dumps({'error': 'Missing LEADS_NEW_OB column in Excel file.'}) + '\n'
            return


        valid_rows = dataframe[dataframe['LEADS_CHNAME'].notna()] if 'LEADS_CHNAME' in dataframe.columns else pd.DataFrame()
        if 'LEADS_CHNAME' not in dataframe.columns:
            yield json.dumps({'error': 'Missing LEADS_CHNAME column in Excel file.'}) + '\n'
            return

        total_records = len(valid_rows)
        if total_records == 0:
            yield json.dumps({'error': 'No valid rows found (LEADS_CHNAME missing or empty file).'}) + '\n'
            return
        
        user_name = user_info.get("name", user_email)
        selected_folder_audit = session_state.get('selected_folder', 'Unknown')
        selected_mode_audit = session_state.get('selected_mode', 'Unknown')
        selected_dl_type_audit = session_state.get('selected_dl_type', 'Unknown')

        # Add audit entry (now async)
        audit_id = await add_audit_entry(
            selected_folder_audit, user_name, total_records, 
            selected_mode_audit, selected_folder_audit, selected_dl_type_audit
        )
        
        output_format = session_state.get('output_format', 'zip')
        user_output_dir = Path(session_state['user_output_dir']) # Ensure this is a Path object

        # Use a single TemporaryDirectory for all temp files related to this run
        with TemporaryDirectory() as global_temp_dir_str:
            global_temp_dir = Path(global_temp_dir_str)

            if output_format == "zip":
                # Short unique ID (timestamp + 4-char UUID)
                timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
                short_uid = uuid.uuid4().hex[:4]

                # Optional: only take part before @ in email
                email_prefix = user_email.split('@')[0].replace('.', '_').lower()

                # Build compact filename
                filename = f"docs_{email_prefix}_{timestamp}_{short_uid}.zip"

                # Create zip file directly in the final location to avoid moving
                final_zip_path_on_server = user_output_dir / filename
                zipf = zipfile.ZipFile(final_zip_path_on_server, 'w', zipfile.ZIP_DEFLATED)

            # Save base and transmittal templates to the global temp dir to avoid issues with original paths
            temp_base_path = None
            if session_state.get('selected_mode') in ["DL Only", "DL w/ Transmittal"] and session_state.get('base_template'):
                temp_base_path = global_temp_dir / "base_template.docx"
                await asyncio.to_thread(session_state['base_template'].save, temp_base_path)
            
            temp_transmittal_path = None
            if session_state.get('selected_mode') in ["DL w/ Transmittal", "Transmittal Only"] and session_state.get('transmittal_template_path'):
                temp_transmittal_path = global_temp_dir / "transmittal_template.docx"
                # Document constructor and save are sync
                original_transmittal_doc = await asyncio.to_thread(Document, session_state['transmittal_template_path'])
                await asyncio.to_thread(original_transmittal_doc.save, temp_transmittal_path)


            processed_records_count = 0
            start_time_processing = time.time()
            account_batch_to_db = []
            db_batch_size = 100
            area_docx_files_for_print = {}

            # Progress callback for batch_convert_libreoffice
            async def conversion_progress_callback(progress_data):
                # This callback might be called from a thread if batch_convert_libreoffice uses to_thread
                # Ensure this yield is safe (FastAPI StreamingResponse handles it)
                if progress_data['type'] == 'conversion_start':
                    yield json.dumps({
                        'progress': 85, # Assuming 85% is start of conversion phase
                        'message': f"Starting PDF conversion: {progress_data['total_files']} files in {progress_data['total_batches']} batches",
                        'conversion_details': progress_data
                    }) + '\n'
                elif progress_data['type'] == 'batch_start':
                     # Calculate progress within the 85-95% range for conversion
                    batch_progress_percent = (progress_data['batch_id'] / progress_data.get('total_batches', 1)) * 10 
                    yield json.dumps({
                        'progress': 85 + batch_progress_percent,
                        'message': f"Converting batch {progress_data['batch_id']}/{progress_data.get('total_batches',1)}: {progress_data['batch_size']} files (Attempt {progress_data['attempt']})",
                        'conversion_details': progress_data
                    }) + '\n'
                elif progress_data['type'] == 'batch_complete':
                    batch_progress_percent = (progress_data['batch_id'] / progress_data.get('total_batches', 1)) * 10
                    yield json.dumps({
                        'progress': 85 + batch_progress_percent, # Update progress based on completed batches
                        'message': f"Batch {progress_data['batch_id']} complete: {progress_data['successful']}/{progress_data['successful'] + progress_data['failed']} files ({progress_data['success_rate']:.1f}% success)",
                        'conversion_details': progress_data
                    }) + '\n'
                elif progress_data['type'] == 'conversion_summary':
                    yield json.dumps({
                        'progress': 95, # Conversion phase finished
                        'message': f"PDF conversion complete: {progress_data['successful']}/{progress_data['total_files']} files in {progress_data['total_time_formatted']} ({progress_data['conversion_rate']:.1f} files/sec)",
                        'conversion_summary': progress_data
                    }) + '\n'


            for final_area, group_df in valid_rows.groupby('FINAL_AREA'):
                logger.debug(f"Processing FINAL_AREA: {final_area} ({len(group_df)} records) for user {user_email}")
                area_output_dir = user_output_dir / final_area
                await asyncio.to_thread(os.makedirs, area_output_dir, exist_ok=True) # os.makedirs is sync
                
                current_area_docx_files = [] # DOCX files generated for this area

                if session_state.get('selected_mode') in ["DL Only", "DL w/ Transmittal"]:
                    if not temp_base_path: # Check if base template was loaded
                        yield json.dumps({'error': f"Base template not available for FINAL_AREA: {final_area}. Processing skipped for this area."}) + '\n'
                        continue

                    for _, row in group_df.iterrows():
                        # doc_code = generate_doc_code()
                        account_batch_to_db.append((
                            audit_id, row.get('DL_CODE', ''), row.get('LEADS_CHNAME', ''),
                            row.get('DL_ADDRESS', ''), row.get('FINAL_AREA', '')
                        ))
                        if len(account_batch_to_db) >= db_batch_size:
                            await add_processed_accounts(audit_id, account_batch_to_db) # add_processed_accounts is async
                            account_batch_to_db = []

                        barcode_val = row.get('DL_CODE', '')
                        barcode_buffer = await asyncio.to_thread(generate_barcode, barcode_val) if barcode_val else None # generate_barcode is sync
                        
                        amount_val = row.get('LEADS_NEW_OB', '0.00') # Use LEADS_NEW_OB for amount
                        amount_words = await asyncio.to_thread(amount_to_words, amount_val) # amount_to_words is sync

                        mapping = {f"«{col.upper()}»": str(row[col]) if pd.notnull(row[col]) else "" for col in dataframe.columns}
                        mapping.update({
                            "«IMAGE_BARCODE»": barcode_buffer or "", # Placeholder for fill_template logic
                            "«DL_DATE»": today_date,
                            "«AMOUNT_ABBR»": amount_words, # This should be «AMOUNT_WORDS» or similar based on template
                            # Signature is handled by combine_templates if path is provided
                        })
                        
                        # fill_template and Document constructor are sync
                        filled_doc_obj = await asyncio.to_thread(Document, temp_base_path)
                        filled_doc_obj = await asyncio.to_thread(fill_template, filled_doc_obj, mapping, barcode_buffer)
                        
                        if filled_doc_obj:
                            unique_name = f"dl_{final_area}_{uuid.uuid4().hex[:8]}.docx"
                            docx_output_path = area_output_dir / unique_name
                            await asyncio.to_thread(filled_doc_obj.save, docx_output_path)
                            current_area_docx_files.append(str(docx_output_path))
                        
                        processed_records_count += 1
                        progress_percent = (processed_records_count / total_records) * 80 # Up to 80% for DOCX generation
                        yield json.dumps({
                            'progress': progress_percent,
                            'message': f"Generating DOCX {processed_records_count}/{total_records} (Area: {final_area})"
                        }) + '\n'
                        await asyncio.sleep(0) # Yield control

                if session_state.get('selected_mode') in ["DL w/ Transmittal", "Transmittal Only"]:
                    if not temp_transmittal_path: # Check if transmittal template was loaded
                        yield json.dumps({'error': f"Transmittal template not available for FINAL_AREA: {final_area}. Transmittal skipped."}) + '\n'
                    else:
                        # fill_transmittal_template is sync
                        transmittal_docs_list = await asyncio.to_thread(fill_transmittal_template, temp_transmittal_path, group_df)
                        if transmittal_docs_list:
                            for doc_idx, transmittal_doc_obj in enumerate(transmittal_docs_list):
                                unique_name = f"transmittal_{final_area}_{doc_idx}_{uuid.uuid4().hex[:8]}.docx"
                                transmittal_docx_output_path = area_output_dir / unique_name
                                await asyncio.to_thread(transmittal_doc_obj.save, transmittal_docx_output_path)
                                current_area_docx_files.append(str(transmittal_docx_output_path))
                        
                        # If mode is Transmittal Only, add accounts to DB here
                        if session_state.get('selected_mode') == "Transmittal Only":
                            for _, row in group_df.iterrows():
                                # doc_code = generate_doc_code()
                                account_batch_to_db.append((
                                    audit_id,row.get('DL_CODE', ''), row.get('LEADS_CHNAME', ''),
                                    row.get('DL_ADDRESS', ''), row.get('FINAL_AREA', '')
                                ))
                                if len(account_batch_to_db) >= db_batch_size:
                                    await add_processed_accounts(audit_id, account_batch_to_db)
                                    account_batch_to_db = []
                            # Update progress for Transmittal Only mode (generation part)
                            processed_records_count += len(group_df) # Count all records in group as processed for this step
                            progress_percent = (processed_records_count / total_records) * 80
                            yield json.dumps({
                                'progress': progress_percent,
                                'message': f"Generated Transmittal DOCX for {final_area}"
                            }) + '\n'
                            await asyncio.sleep(0)


                if output_format == "zip" and current_area_docx_files:
                    # batch_convert_libreoffice is now async
                    pdf_files_for_area = await batch_convert_libreoffice(current_area_docx_files, area_output_dir, progress_callback=conversion_progress_callback)
                    
                    dl_merger = PdfMerger()
                    transmittal_merger = PdfMerger()
                    has_dl_pdfs = False
                    has_transmittal_pdfs = False

                    for pdf_file_path_str in pdf_files_for_area:
                        pdf_file_path = Path(pdf_file_path_str)
                        if "transmittal" in pdf_file_path.name:
                            await asyncio.to_thread(transmittal_merger.append, str(pdf_file_path))
                            has_transmittal_pdfs = True
                        elif "dl" in pdf_file_path.name: # Assuming DL files are identified by "dl"
                            await asyncio.to_thread(dl_merger.append, str(pdf_file_path))
                            has_dl_pdfs = True
                    
                    if has_dl_pdfs and session_state.get('selected_mode') in ["DL Only", "DL w/ Transmittal"]:
                        dl_merged_path = area_output_dir / f"{final_area}_DL_MERGED.pdf"
                        with open(dl_merged_path, 'wb') as output_f: # sync file open/write
                            await asyncio.to_thread(dl_merger.write, output_f)
                        await asyncio.to_thread(zipf.write, dl_merged_path, f"{final_area}/{final_area}_DL_MERGED.pdf")
                    if has_transmittal_pdfs and session_state.get('selected_mode') in ["DL w/ Transmittal", "Transmittal Only"]:
                        transmittal_merged_path = area_output_dir / f"{final_area}_TRANSMITTAL_MERGED.pdf"
                        with open(transmittal_merged_path, 'wb') as output_f: # sync file open/write
                            await asyncio.to_thread(transmittal_merger.write, output_f)
                        await asyncio.to_thread(zipf.write, transmittal_merged_path, f"{final_area}/{final_area}_TRANSMITTAL_MERGED.pdf")
                    
                    # Close mergers
                    await asyncio.to_thread(dl_merger.close)
                    await asyncio.to_thread(transmittal_merger.close)

                    # Cleanup individual DOCX and PDF files for this area after zipping
                    for f_path_str in current_area_docx_files + pdf_files_for_area:
                        try:
                            await asyncio.to_thread(os.remove, f_path_str)
                        except OSError as e:
                            logger.warning(f"Could not delete temp file {f_path_str}: {e}")
                
                elif output_format == "print" and current_area_docx_files:
                    area_docx_files_for_print[final_area] = current_area_docx_files
                    # Progress for print format (DOCX generation is the main work)
                    yield json.dumps({
                        'progress': (processed_records_count / total_records) * 100 if total_records > 0 else 100,
                        'message': f"Prepared {len(current_area_docx_files)} DOCX files for printing in {final_area}"
                    }) + '\n'
                    await asyncio.sleep(0)


            if account_batch_to_db: # Process any remaining accounts
                await add_processed_accounts(audit_id, account_batch_to_db)

            total_time_val = time.time() - start_time_processing
            total_time_formatted_val = format_time_duration(total_time_val) # sync helper

            if output_format == "zip":
                if zipf:
                    # Properly close the zip file
                    try:
                        await asyncio.to_thread(zipf.close)
                        zipf = None  # Clear reference
                        logger.debug(f"Zip file closed successfully: {final_zip_path_on_server}")
                    except Exception as close_error:
                        logger.error(f"Error closing zip file: {close_error}")
                        zipf = None
                
                session_state['zip_path'] = str(final_zip_path_on_server)

                yield json.dumps({
                    'progress': 100,
                    'message': f"Processing complete! ZIP file generated in {total_time_formatted_val}.",
                    'download_ready': True,
                    'processing_summary': { 'total_records': total_records, 'total_time_formatted': total_time_formatted_val }
                }) + '\n'
            else: # print format
                yield json.dumps({
                    'progress': 100,
                    'message': f"Processing complete! DOCX files ready for printing in {total_time_formatted_val}.",
                    'print_ready': True,
                    'areas': list(area_docx_files_for_print.keys()),
                    # 'docx_files': area_docx_files_for_print, # Sending all file paths might be too large
                    'processing_summary': { 'total_records': total_records, 'total_time_formatted': total_time_formatted_val }
                }) + '\n'
                
    except Exception as e: # Catch errors in initial setup of the generator
        logger.error(f"Outer error in generate_pdfs_stream for user {user_email}: {e}", exc_info=True)
        yield json.dumps({'error': f'Processing setup failed: {str(e)}'}) + '\n'
    finally:
        # Ensure zip file is closed and lock is released
        if zipf:
            try:
                await asyncio.to_thread(zipf.close)
            except Exception as close_error:
                logger.warning(f"Error closing zip file: {close_error}")
        
        # Ensure lock is released
        if 'session_state' in locals() and session_state['processing_lock'].locked():
            await asyncio.to_thread(session_state['processing_lock'].release)


def get_zip_for_download(user_email: str): # Sync
    """Get ZIP file for download"""
    session_state = get_user_session_state(user_email) # Sync
    zip_path_str = session_state.get('zip_path')
    
    if not zip_path_str:
        logger.error(f"Zip path not found in session for user {user_email}")
        raise HTTPException(status_code=404, detail="ZIP file path not found in session. Please generate files again.")
        
    zip_path = Path(zip_path_str)
    if not zip_path.exists():
        logger.error(f"ZIP file not found at path: {zip_path} for user {user_email}")
        raise HTTPException(status_code=404, detail="ZIP file not found on server. It might have been cleaned up or not generated.")
    
    # The cleanup of the zip file itself is handled by reset_user_session_state or a general cleanup task.
    # Here, we just serve the file.
    
    return FileResponse(str(zip_path), filename=zip_path.name, media_type="application/zip")
