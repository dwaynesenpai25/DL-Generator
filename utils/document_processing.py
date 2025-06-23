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
from utils.config import logger
from utils.session_management import get_user_session_state
from utils.document_utils import generate_barcode, generate_qrcode, amount_to_words, combine_templates, extract_placeholders
from utils.pdf_conversion import batch_convert_libreoffice
from utils.database import generate_doc_code, add_audit_entry, add_processed_accounts
import os
import re
from PyPDF2 import PdfMerger 

def get_raw_file(file):
  try:
      contents = file.file.read()
      df = pd.read_excel(BytesIO(contents), dtype=str)
      return df
  except Exception as e:
      logger.error(f"Error reading file: {e}")
      return pd.DataFrame([])

def process_excel_file(file):
  if not file.filename.endswith('.xlsx'):
      raise HTTPException(status_code=400, detail="Invalid file format. Please upload an .xlsx file")
  df = get_raw_file(file)
  if df.empty:
      raise HTTPException(status_code=500, detail="Failed to read Excel file or file is empty.")
  
  required_cols = ['FINAL_AREA', 'DL_CODE', 'LEADS_CHNAME', 'DL_ADDRESS', 'LEADS_NEW_OB']
  missing_cols = [col for col in required_cols if col not in df.columns]
  if missing_cols:
      error_detail = f"Excel file must contain the following columns: {', '.join(missing_cols)}"
      logger.error(error_detail)
      raise HTTPException(status_code=400, detail=error_detail)
      
  return {"data": df.to_dict(orient='records')}

def fill_template(doc, mapping, barcode_buffer=None):
  def replace_in_text(text):
      for k, v in mapping.items():
          if k == "«IMAGE_BARCODE»" and v:
              continue
          text = text.replace(k, str(v) if v is not None else "")
      return text

  for para in doc.paragraphs:
      for run in para.runs:
          if run.text:
              run.text = replace_in_text(run.text)

  for section in doc.sections:
      for header_footer_item in [section.header, section.first_page_header, section.even_page_header, 
                                 section.footer, section.first_page_footer, section.even_page_footer]:
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
                                  if cell.paragraphs:
                                      target_para = cell.paragraphs[0]
                                      target_para.paragraph_format.left_indent = Pt(-20)
                                      run_pic = target_para.add_run()
                                      try:
                                          run_pic.add_picture(barcode_buffer, width=Inches(3.0), height=Inches(0.35))
                                      except Exception as e:
                                          logger.error(f"Error adding barcode picture: {e}")
                                  else:
                                      new_para = cell.add_paragraph()
                                      new_para.paragraph_format.left_indent = Pt(-20)
                                      run_pic = new_para.add_run()
                                      try:
                                          run_pic.add_picture(barcode_buffer, width=Inches(3.0), height=Inches(0.35))
                                      except Exception as e:
                                          logger.error(f"Error adding barcode picture: {e}")
                              else:
                                  for run in para.runs:
                                      if run.text:
                                          run.text = replace_in_text(run.text)
  return doc

def clear_placeholders(inner_table):
  def replace_in_text(text):
      return re.sub(r'«[^»]+»', "", text)
  for row in inner_table.rows:
      for cell in row.cells:
          for para in cell.paragraphs:
              for run in para.runs:
                  if run.text:
                      run.text = replace_in_text(run.text)

def fill_inner_table(inner_table, mapping, qrcode_buffer=None):
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

def fill_transmittal_template(template_path, group_df):
  try:
      temp_doc_orig = Document(template_path)
      if not temp_doc_orig.tables:
          logger.error("No tables found in the transmittal template.")
          return None
      
      total_records_in_group = len(group_df)
      total_pages_needed = (total_records_in_group + 3) // 4
      
      filled_transmittal_docs = []

      for page_num in range(total_pages_needed):
          current_page_doc = Document(template_path)
          current_page_table = current_page_doc.tables[0]

          start_idx = page_num * 4
          end_idx = start_idx + 4
          page_records_df = group_df.iloc[start_idx:end_idx]

          for table_row_idx, table_row_obj in enumerate(current_page_table.rows):
              if table_row_idx < len(page_records_df):
                  record_data = page_records_df.iloc[table_row_idx]
                  mapping = {f"«{col.upper()}»": str(record_data[col]) for col in group_df.columns if pd.notnull(record_data[col])}
                  
                  qrcode_buffer = None
                  if dl_code := record_data.get('DL_CODE', ''):
                      qrcode_buffer = generate_qrcode(dl_code)
                  
                  for cell in table_row_obj.cells:
                      if cell.tables:
                          inner_table = cell.tables[0]
                          fill_inner_table(inner_table, mapping, qrcode_buffer)
              else:
                  for cell in table_row_obj.cells:
                      if cell.tables:
                          inner_table = cell.tables[0]
                          clear_placeholders(inner_table)
          
          filled_transmittal_docs.append(current_page_doc)
          
      return filled_transmittal_docs
  except Exception as e:
      logger.error(f"Failed to fill transmittal template: {e}", exc_info=True)
      return None

def format_time_duration(seconds):
  """Convert seconds to human-readable format"""
  if seconds < 60:
      return f"{seconds:.1f} seconds"
  elif seconds < 3600:
      minutes = seconds / 60
      return f"{minutes:.1f} minutes"
  else:
      hours = seconds / 3600
      return f"{hours:.1f} hours"

async def generate_pdfs_stream(uploaded_file, dataframe: pd.DataFrame, user_info: dict):
  user_email = user_info.get("email", "")
  session_state = await asyncio.to_thread(get_user_session_state, user_email)
  
  if not await asyncio.to_thread(session_state['processing_lock'].acquire, blocking=False):
      yield json.dumps({'error': 'Another processing task is already running for this user.'}) + '\n'
      return
  
  zipf = None
  try:
      if not session_state.get('selected_mode'):
          yield json.dumps({'error': 'No processing mode selected. Please select a mode first.'}) + '\n'
          return

      if session_state.get('excel_errors'):
          error_msg = f"Excel validation errors found: {'; '.join(session_state['excel_errors'])}"
          yield json.dumps({'error': error_msg}) + '\n'
          return

      if session_state.get('signature_error'):
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
      
      # Data transformations
      if 'DL_ADDRESS' in dataframe.columns:
          dataframe['DL_ADDRESS'] = dataframe['DL_ADDRESS'].astype(str).str.upper()
      if 'LEADS_NEW_OB' in dataframe.columns:
          dataframe['LEADS_NEW_OB'] = dataframe['LEADS_NEW_OB'].apply(lambda x: f"{float(x):,.2f}" if pd.notnull(x) and str(x).replace('.', '', 1).isdigit() else str(x))
      else:
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

      # Add audit entry
      audit_id = await add_audit_entry(
          selected_folder_audit, user_name, total_records, 
          selected_mode_audit, selected_folder_audit, selected_dl_type_audit
      )
      
      output_format = session_state.get('output_format', 'zip')
      user_output_dir = Path(session_state['user_output_dir'])

      # Enhanced progress tracking variables
      total_areas = len(valid_rows.groupby('FINAL_AREA'))
      current_area_index = 0
      total_docx_files = 0
      processed_docx_files = 0

      with TemporaryDirectory() as global_temp_dir_str:
          global_temp_dir = Path(global_temp_dir_str)

          if output_format == "zip":
              timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
              short_uid = uuid.uuid4().hex[:4]
              email_prefix = user_email.split('@')[0].replace('.', '_').lower()
              filename = f"{email_prefix}_{timestamp}_{short_uid}.zip"
              final_zip_path_on_server = user_output_dir / filename
              zipf = zipfile.ZipFile(final_zip_path_on_server, 'w', zipfile.ZIP_DEFLATED)

          # Save templates to temp directory
          temp_base_path = None
          if session_state.get('selected_mode') in ["DL Only", "DL w/ Transmittal"] and session_state.get('base_template'):
              temp_base_path = global_temp_dir / "base_template.docx"
              await asyncio.to_thread(session_state['base_template'].save, temp_base_path)
          
          temp_transmittal_path = None
          if session_state.get('selected_mode') in ["DL w/ Transmittal", "Transmittal Only"] and session_state.get('transmittal_template_path'):
              temp_transmittal_path = global_temp_dir / "transmittal_template.docx"
              original_transmittal_doc = await asyncio.to_thread(Document, session_state['transmittal_template_path'])
              await asyncio.to_thread(original_transmittal_doc.save, temp_transmittal_path)

          processed_records_count = 0
          start_time_processing = time.time()
          account_batch_to_db = []
          db_batch_size = 300
          area_docx_files_for_print = {}

          # Process each area
          for final_area, group_df in valid_rows.groupby('FINAL_AREA'):
              current_area_index += 1
              area_records = len(group_df)
              
              # Enhanced area progress reporting
              area_progress = (current_area_index / total_areas) * 70  # Up to 70% for area processing
              yield json.dumps({
                  'progress': area_progress,
                  'message': f"🏢 Processing area {current_area_index}/{total_areas}: {final_area} ({area_records} records)",
                  'stage': 'area_processing',
                  'area_details': {
                      'current_area': final_area,
                      'area_index': current_area_index,
                      'total_areas': total_areas,
                      'records_in_area': area_records
                  }
              }) + '\n'
              
              logger.debug(f"Processing FINAL_AREA: {final_area} ({len(group_df)} records) for user {user_email}")
              area_output_dir = user_output_dir / final_area
              await asyncio.to_thread(os.makedirs, area_output_dir, exist_ok=True)
              
              current_area_docx_files = []

              if session_state.get('selected_mode') in ["DL Only", "DL w/ Transmittal"]:
                  if not temp_base_path:
                      yield json.dumps({'error': f"Base template not available for FINAL_AREA: {final_area}. Processing skipped for this area."}) + '\n'
                      continue

                  # Process individual records with enhanced progress
                  for record_index, (_, row) in enumerate(group_df.iterrows()):
                      record_progress = area_progress + (record_index / area_records) * (70 / total_areas)
                      
                      yield json.dumps({
                          'progress': record_progress,
                          'message': f"📝 Generating document {record_index + 1}/{area_records} for {row.get('LEADS_CHNAME', 'Unknown')} in {final_area}",
                          'stage': 'document_generation',
                          'document_details': {
                              'area': final_area,
                              'record_index': record_index + 1,
                              'total_records_in_area': area_records,
                              'client_name': row.get('LEADS_CHNAME', 'Unknown'),
                              'dl_code': row.get('DL_CODE', '')
                          }
                      }) + '\n'

                      account_batch_to_db.append((
                          audit_id, row.get('DL_CODE', ''), row.get('LEADS_CHNAME', ''),
                          row.get('DL_ADDRESS', ''), row.get('FINAL_AREA', '')
                      ))
                      if len(account_batch_to_db) >= db_batch_size:
                          await add_processed_accounts(audit_id, account_batch_to_db)
                          account_batch_to_db = []

                      barcode_val = row.get('DL_CODE', '')
                      barcode_buffer = await asyncio.to_thread(generate_barcode, barcode_val) if barcode_val else None
                      
                      amount_val = row.get('LEADS_NEW_OB', '0.00')
                      amount_words = await asyncio.to_thread(amount_to_words, amount_val)

                      mapping = {f"«{col.upper()}»": str(row[col]) if pd.notnull(row[col]) else "" for col in dataframe.columns}
                      mapping.update({
                          "«IMAGE_BARCODE»": barcode_buffer or "",
                          "«DL_DATE»": today_date,
                          "«AMOUNT_ABBR»": amount_words,
                      })
                      
                      filled_doc_obj = await asyncio.to_thread(Document, temp_base_path)
                      filled_doc_obj = await asyncio.to_thread(fill_template, filled_doc_obj, mapping, barcode_buffer)
                      
                      if filled_doc_obj:
                          unique_name = f"dl_{final_area}_{uuid.uuid4().hex[:8]}.docx"
                          docx_output_path = area_output_dir / unique_name
                          await asyncio.to_thread(filled_doc_obj.save, docx_output_path)
                          current_area_docx_files.append(str(docx_output_path))
                          total_docx_files += 1
                      
                      processed_records_count += 1
                      await asyncio.sleep(0)  # Yield control

              if session_state.get('selected_mode') in ["DL w/ Transmittal", "Transmittal Only"]:
                  if not temp_transmittal_path:
                      yield json.dumps({'error': f"Transmittal template not available for FINAL_AREA: {final_area}. Transmittal skipped."}) + '\n'
                  else:
                      # Calculate transmittal pages needed
                      total_pages_needed = (area_records + 3) // 4
                      
                      yield json.dumps({
                          'progress': area_progress + 5,
                          'message': f"📋 Generating transmittal documents for {final_area} ({total_pages_needed} pages)",
                          'stage': 'transmittal_generation',
                          'transmittal_details': {
                              'area': final_area,
                              'current_page': 0,
                              'total_pages': total_pages_needed,
                              'records_per_page': 0
                          }
                      }) + '\n'

                      transmittal_docs_list = await asyncio.to_thread(fill_transmittal_template, temp_transmittal_path, group_df)
                      if transmittal_docs_list:
                          for doc_idx, transmittal_doc_obj in enumerate(transmittal_docs_list):
                              # Calculate records per page for this specific page
                              records_on_this_page = min(4, area_records - (doc_idx * 4))
                              
                              # Progress for each transmittal page
                              transmittal_progress = area_progress + 5 + (doc_idx / len(transmittal_docs_list)) * 5
                              yield json.dumps({
                                  'progress': transmittal_progress,
                                  'message': f"📋 Creating transmittal page {doc_idx + 1}/{len(transmittal_docs_list)} for {final_area} ({records_on_this_page} records)",
                                  'stage': 'transmittal_generation',
                                  'transmittal_details': {
                                      'area': final_area,
                                      'current_page': doc_idx + 1,
                                      'total_pages': len(transmittal_docs_list),
                                      'records_per_page': records_on_this_page
                                  }
                              }) + '\n'
                              
                              unique_name = f"transmittal_{final_area}_{doc_idx}_{uuid.uuid4().hex[:8]}.docx"
                              transmittal_docx_output_path = area_output_dir / unique_name
                              await asyncio.to_thread(transmittal_doc_obj.save, transmittal_docx_output_path)
                              current_area_docx_files.append(str(transmittal_docx_output_path))
                      
                      if session_state.get('selected_mode') == "Transmittal Only":
                          for _, row in group_df.iterrows():
                              account_batch_to_db.append((
                                  audit_id, row.get('DL_CODE', ''), row.get('LEADS_CHNAME', ''),
                                  row.get('DL_ADDRESS', ''), row.get('FINAL_AREA', '')
                              ))
                              if len(account_batch_to_db) >= db_batch_size:
                                  await add_processed_accounts(audit_id, account_batch_to_db)
                                  account_batch_to_db = []
                          processed_records_count += len(group_df)

              # PDF conversion and merging with enhanced progress
              if output_format == "zip" and current_area_docx_files:
                  yield json.dumps({
                      'progress': 75 + (current_area_index / total_areas) * 5,
                      'message': f"🔄 Preparing PDF conversion for {len(current_area_docx_files)} files in {final_area}",
                      'stage': 'pdf_conversion_start',
                      'conversion_details': {
                          'area': final_area,
                          'files_to_convert': len(current_area_docx_files),
                          'estimated_batches': (len(current_area_docx_files) + 9) // 10
                      }
                  }) + '\n'

                  # Convert the PDF files with a custom progress handler
                  pdf_files_for_area = []
                  
                  # Calculate batches
                  batch_size = 250
                  batches = [current_area_docx_files[i:i + batch_size] for i in range(0, len(current_area_docx_files), batch_size)]
                  total_batches = len(batches)
                  
                  # Send conversion start
                  yield json.dumps({
                      'progress': 80,
                      'message': f"🔄 Starting PDF conversion for {len(current_area_docx_files)} files in {total_batches} batches",
                      'stage': 'conversion',
                      'conversion_details': {
                          'total_files': len(current_area_docx_files),
                          'total_batches': total_batches,
                          'batch_size': batch_size,
                          'current_batch': 0,
                          'batches_completed': 0
                      }
                  }) + '\n'
                  
                  # Process each batch with progress updates
                  for batch_id, batch_files in enumerate(batches, 1):
                      batch_progress = 80 + (batch_id / total_batches) * 15
                      
                      # Send batch start
                      yield json.dumps({
                          'progress': batch_progress,
                          'message': f"📄 Processing batch {batch_id}/{total_batches} ({len(batch_files)} files)",
                          'stage': 'conversion',
                          'conversion_details': {
                              'current_batch': batch_id,
                              'total_batches': total_batches,
                              'batch_size': len(batch_files),
                              'attempt': 1,
                              'batches_completed': batch_id - 1,
                              'files_in_current_batch': len(batch_files)
                          }
                      }) + '\n'
                      
                      # Convert this batch
                      batch_pdfs = await batch_convert_libreoffice(
                          batch_files, 
                          area_output_dir, 
                          batch_size=len(batch_files)  # Process all files in this batch at once
                      )
                      pdf_files_for_area.extend(batch_pdfs)
                      
                      # Send batch complete
                      success_rate = (len(batch_pdfs) / len(batch_files)) * 100 if batch_files else 0
                      status_icon = "✅" if success_rate > 90 else "⚠️" if success_rate > 70 else "❌"
                      
                      yield json.dumps({
                          'progress': batch_progress,
                          'message': f"{status_icon} Batch {batch_id}/{total_batches} complete: {len(batch_pdfs)}/{len(batch_files)} files ({success_rate:.1f}% success)",
                          'stage': 'conversion',
                          'conversion_details': {
                              'batch_id': batch_id,
                              'total_batches': total_batches,
                              'successful': len(batch_pdfs),
                              'failed': len(batch_files) - len(batch_pdfs),
                              'success_rate': success_rate,
                              'batch_time': 0,  # We don't have timing here, but that's ok
                              'batches_completed': batch_id,
                              'batches_remaining': total_batches - batch_id
                          }
                      }) + '\n'
                  
                  # Send conversion summary
                  total_success_rate = (len(pdf_files_for_area) / len(current_area_docx_files)) * 100 if current_area_docx_files else 0
                  status_icon = "🎉" if total_success_rate > 95 else "✅" if total_success_rate > 80 else "⚠️"
                  
                  yield json.dumps({
                      'progress': 95,
                      'message': f"{status_icon} PDF conversion complete: {len(pdf_files_for_area)}/{len(current_area_docx_files)} files ({total_success_rate:.1f}% success)",
                      'stage': 'conversion_complete',
                      'conversion_summary': {
                          'total_files': len(current_area_docx_files),
                          'total_batches': total_batches,
                          'successful': len(pdf_files_for_area),
                          'failed': len(current_area_docx_files) - len(pdf_files_for_area),
                          'success_rate': total_success_rate,
                          'total_time': "N/A",
                          'conversion_rate': 0,
                          'batches_processed': total_batches
                      }
                  }) + '\n'
              
              elif output_format == "print" and current_area_docx_files:
                  area_docx_files_for_print[final_area] = current_area_docx_files
                  yield json.dumps({
                      'progress': (current_area_index / total_areas) * 100,
                      'message': f"📄 Prepared {len(current_area_docx_files)} DOCX files for printing in {final_area}",
                      'stage': 'print_preparation',
                      'print_details': {
                          'area': final_area,
                          'files_prepared': len(current_area_docx_files)
                      }
                  }) + '\n'

          # Final processing
          if account_batch_to_db:
              await add_processed_accounts(audit_id, account_batch_to_db)

          total_time_val = time.time() - start_time_processing
          total_time_formatted_val = format_time_duration(total_time_val)

          if output_format == "zip":
              if zipf:
                  try:
                      await asyncio.to_thread(zipf.close)
                      zipf = None
                      logger.debug(f"Zip file closed successfully: {final_zip_path_on_server}")
                  except Exception as close_error:
                      logger.error(f"Error closing zip file: {close_error}")
                      zipf = None
              
              session_state['zip_path'] = str(final_zip_path_on_server)

              yield json.dumps({
                  'progress': 100,
                  'message': f"🎉 Processing complete! ZIP file generated successfully in {total_time_formatted_val}",
                  'stage': 'complete',
                  'download_ready': True,
                  'processing_summary': {
                      'total_records': total_records,
                      'total_areas': total_areas,
                      'total_time_formatted': total_time_formatted_val,
                      'files_generated': total_docx_files
                  }
              }) + '\n'
          else:
              yield json.dumps({
                  'progress': 100,
                  'message': f"🎉 Processing complete! DOCX files ready for printing in {total_time_formatted_val}",
                  'stage': 'complete',
                  'print_ready': True,
                  'areas': list(area_docx_files_for_print.keys()),
                  'processing_summary': {
                      'total_records': total_records,
                      'total_areas': total_areas,
                      'total_time_formatted': total_time_formatted_val,
                      'files_generated': total_docx_files
                  }
              }) + '\n'
              
  except Exception as e:
      logger.error(f"Outer error in generate_pdfs_stream for user {user_email}: {e}", exc_info=True)
      yield json.dumps({
          'error': f'Processing failed: {str(e)}',
          'stage': 'error'
      }) + '\n'
  finally:
      if zipf:
          try:
              await asyncio.to_thread(zipf.close)
          except Exception as close_error:
              logger.warning(f"Error closing zip file: {close_error}")
      
      if 'session_state' in locals() and session_state['processing_lock'].locked():
          await asyncio.to_thread(session_state['processing_lock'].release)

def get_zip_for_download(user_email: str):
  """Get ZIP file for download"""
  session_state = get_user_session_state(user_email)
  zip_path_str = session_state.get('zip_path')
  
  if not zip_path_str:
      logger.error(f"Zip path not found in session for user {user_email}")
      raise HTTPException(status_code=404, detail="ZIP file path not found in session. Please generate files again.")
      
  zip_path = Path(zip_path_str)
  if not zip_path.exists():
      logger.error(f"ZIP file not found at path: {zip_path} for user {user_email}")
      raise HTTPException(status_code=404, detail="ZIP file not found on server. It might have been cleaned up or not generated.")
  
  return FileResponse(str(zip_path), filename=zip_path.name, media_type="application/zip")
