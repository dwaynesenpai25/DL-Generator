import subprocess
import time
import psutil
import shutil
from pathlib import Path
from tempfile import TemporaryDirectory
from utils.config import logger
import os

def kill_libreoffice_processes():
    killed_count = 0
    try:
        for proc in psutil.process_iter(['pid', 'name']):
            if 'soffice' in proc.info['name'].lower():
                try:
                    proc.kill()
                    killed_count += 1
                except psutil.NoSuchProcess:
                    continue
        if killed_count > 0:
            logger.info(f"Killed {killed_count} existing LibreOffice processes")
        time.sleep(2)
    except Exception as e:
        logger.error(f"Error killing LibreOffice processes: {e}")

def convert_batch_with_retry(batch_files, output_dir, batch_id, timeout=180, progress_callback=None):
    max_retries = 3
    output_dir = Path(output_dir)
    
    for attempt in range(max_retries):
        try:
            with TemporaryDirectory() as temp_batch_dir:
                temp_output = Path(temp_batch_dir)
                logger.debug(f"Batch {batch_id} (Attempt {attempt + 1}): Converting {len(batch_files)} files...")
                
                # Report batch start
                if progress_callback:
                    progress_callback({
                        'type': 'batch_start',
                        'batch_id': batch_id,
                        'batch_size': len(batch_files),
                        'attempt': attempt + 1
                    })
                
                cmd = [
                    r"C:\Program Files\LibreOffice\program\soffice.exe",
                    "--headless",
                    "--invisible",
                    "--nodefault",
                    "--nolockcheck",
                    "--nologo",
                    "--norestore",
                    "--convert-to", "pdf",
                    "--outdir", str(temp_output)
                ] + [str(Path(f)) for f in batch_files]
                
                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0,
                    text=True
                )
                
                try:
                    stdout, stderr = process.communicate(timeout=timeout)
                    if process.returncode != 0:
                        logger.error(f"Batch {batch_id} LibreOffice error (code {process.returncode}): {stderr}")
                        if attempt < max_retries - 1:
                            logger.info(f"Retrying batch {batch_id}...")
                            kill_libreoffice_processes()
                            time.sleep(2)
                            continue
                        return [], batch_files
                except subprocess.TimeoutExpired:
                    process.kill()
                    logger.error(f"Batch {batch_id} timeout after {timeout} seconds (attempt {attempt + 1})")
                    if attempt < max_retries - 1:
                        logger.info(f"Retrying batch {batch_id}...")
                        kill_libreoffice_processes()
                        time.sleep(2)
                        continue
                    return [], batch_files
                
                batch_pdfs = []
                failed_files = []
                
                for docx_path in batch_files:
                    docx_name = Path(docx_path).stem
                    temp_pdf = temp_output / f"{docx_name}.pdf"
                    final_pdf = output_dir / f"{docx_name}.pdf"
                    if temp_pdf.exists():
                        shutil.move(str(temp_pdf), str(final_pdf))
                        batch_pdfs.append(str(final_pdf))
                    else:
                        failed_files.append(docx_path)
                        logger.warning(f"Batch {batch_id}: Failed to convert {Path(docx_path).name}")
                
                success_rate = len(batch_pdfs) / len(batch_files) * 100 if batch_files else 0
                logger.info(f"Batch {batch_id} result: {len(batch_pdfs)}/{len(batch_files)} successful ({success_rate:.1f}%)")
                
                # Report batch completion
                if progress_callback:
                    progress_callback({
                        'type': 'batch_complete',
                        'batch_id': batch_id,
                        'successful': len(batch_pdfs),
                        'failed': len(failed_files),
                        'success_rate': success_rate
                    })
                
                return batch_pdfs, failed_files
                
        except Exception as e:
            logger.error(f"Batch {batch_id} conversion error (attempt {attempt + 1}): {e}")
            if attempt < max_retries - 1:
                kill_libreoffice_processes()
                time.sleep(2)
                continue
            return [], batch_files
    
    return [], batch_files

def batch_convert_libreoffice(docx_files, output_dir, batch_size=350, progress_callback=None):
    if not docx_files:
        return []
    
    pdf_files = []
    total_failed = []
    output_dir = Path(output_dir)
    kill_libreoffice_processes()
    
    batches = [docx_files[i:i + batch_size] for i in range(0, len(docx_files), batch_size)]
    logger.info(f"Converting {len(docx_files)} DOCX files in {len(batches)} batches (size: {batch_size})...")
    
    start_time = time.time()
    
    # Report conversion start
    if progress_callback:
        progress_callback({
            'type': 'conversion_start',
            'total_files': len(docx_files),
            'total_batches': len(batches),
            'batch_size': batch_size
        })
    
    for batch_id, batch in enumerate(batches, 1):
        batch_pdfs, batch_failed = convert_batch_with_retry(
            batch, output_dir, batch_id, progress_callback=progress_callback
        )
        pdf_files.extend(batch_pdfs)
        total_failed.extend(batch_failed)
    
    kill_libreoffice_processes()
    
    total_time = time.time() - start_time
    success_rate = len(pdf_files) / len(docx_files) * 100 if docx_files else 0
    
    # Enhanced conversion summary
    conversion_summary = {
        'type': 'conversion_summary',
        'total_files': len(docx_files),
        'successful': len(pdf_files),
        'failed': len(total_failed),
        'success_rate': success_rate,
        'total_time': total_time,
        'total_time_formatted': format_time_duration(total_time),
        'conversion_rate': len(pdf_files) / total_time if total_time > 0 else 0,
        'total_batches': len(batches)
    }
    
    logger.info(f"\n=== CONVERSION SUMMARY ===")
    logger.info(f"Total files: {len(docx_files)}")
    logger.info(f"Successful: {len(pdf_files)} ({success_rate:.1f}%)")
    logger.info(f"Failed: {len(total_failed)} ({100-success_rate:.1f}%)")
    logger.info(f"Time: {format_time_duration(total_time)} | Rate: {len(pdf_files)/total_time:.1f} PDFs/sec" if total_time > 0 else "Time: 0s")
    
    # Report final summary
    if progress_callback:
        progress_callback(conversion_summary)
    
    return pdf_files

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
