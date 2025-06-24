import asyncio
import time
import psutil
import shutil
import aiofiles.os
from pathlib import Path
from utils.config import logger
import os
import subprocess
import traceback
import platform
import datetime
from tempfile import TemporaryDirectory

# Set the event loop policy globally for Windows at module level
if platform.system() == "Windows":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

async def kill_libreoffice_processes_async():
    killed_count = 0  # Initialize as integer
    try:
        procs = await asyncio.to_thread(list, psutil.process_iter(['pid', 'name']))
        for proc_info in procs:
            name = await asyncio.to_thread(lambda: proc_info.info['name'])
            if 'soffice' in name.lower():
                try:
                    await asyncio.to_thread(proc_info.kill)
                    killed_count = killed_count + 1  # Increment integer
                except psutil.NoSuchProcess:
                    continue
        if killed_count > 0:
            logger.info(f"Killed {killed_count} existing LibreOffice processes (async)")
        await asyncio.sleep(2)
    except Exception as e:
        logger.error(f"Error killing LibreOffice processes (async): {e}\n{traceback.format_exc()}")

async def validate_files_async(files):
    """Validate that input files exist and are accessible."""
    valid_files = []
    invalid_files = []
    for file in files:
        try:
            file_path = Path(file)
            logger.debug(f"Validating: {file_path} | Length: {len(str(file_path))}")
            exists = await aiofiles.os.path.exists(str(file_path))
            if exists and file_path.is_file():
                valid_files.append(str(file_path))
            else:
                invalid_files.append(str(file_path))
                logger.warning(f"File does not exist or is not a file: {file_path}")
        except Exception as e:
            invalid_files.append(str(file_path))
            logger.error(f"Error validating file {file}: {e}\n{traceback.format_exc()}")
    return valid_files, invalid_files

async def convert_batch_with_retry_async(batch_files, output_dir, batch_id, timeout=180, progress_callback=None):
    max_retries = 3
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    valid_files, invalid_files = await validate_files_async(batch_files)
    if not valid_files:
        logger.error(f"Batch {batch_id}: No valid files to process")
        return [], batch_files

    # Log the current event loop type for debugging
    loop = asyncio.get_running_loop()
    logger.debug(f"Event loop in use for batch {batch_id}: {type(loop)}")

    # Ensure C:\Temp exists
    temp_base_dir = "C:\\Temp"
    try:
        os.makedirs(temp_base_dir, exist_ok=True)
        logger.debug(f"Ensured {temp_base_dir} exists")
    except Exception as e:
        logger.error(f"Failed to create {temp_base_dir}: {e}\n{traceback.format_exc()}")
        return [], valid_files + invalid_files

    for attempt in range(max_retries):
        temp_batch_dir = None  # Initialize to avoid undefined variable
        try:
            # Create temporary directory in C:\Temp for short paths
            with TemporaryDirectory(prefix=f"batch_{batch_id}_", dir=temp_base_dir) as temp_batch_dir:
                temp_input = Path(temp_batch_dir)
                temp_output = temp_input / "output"
                temp_output.mkdir()
                logger.debug(f"Batch {batch_id} (Attempt {attempt + 1}): Using temp dir {temp_batch_dir} with {len(valid_files)} files")

                # Copy files to temporary directory
                temp_files = []
                for file in valid_files:
                    src_path = Path(file)
                    dst_path = temp_input / src_path.name
                    await asyncio.to_thread(shutil.copy2, str(src_path), str(dst_path))
                    temp_files.append(str(dst_path))
                    logger.debug(f"Copied {src_path} to {dst_path} | Length: {len(str(dst_path))}")

                if progress_callback:
                    try:
                        update = progress_callback({
                            'type': 'batch_start',
                            'batch_id': batch_id,
                            'batch_size': len(valid_files),
                            'attempt': attempt + 1
                        })
                        if update:
                            pass
                    except Exception as e:
                        logger.error(f"Error in progress callback for batch_start: {e}\n{traceback.format_exc()}")

                libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
                if not await aiofiles.os.path.exists(libreoffice_path):
                    logger.error(f"LibreOffice executable not found at {libreoffice_path}")
                    return [], valid_files

                cmd = [
                    libreoffice_path,
                    "--headless", "--invisible", "--nodefault", "--nolockcheck",
                    "--nologo", "--norestore", "--convert-to", "pdf",
                    "--outdir", str(temp_output)
                ] + temp_files

                # Log command length for debugging
                logger.debug(f"Command length: {len(' '.join(cmd))}")

                # Check if we're on Windows and using SelectorEventLoop
                is_windows_selector = (
                    platform.system() == "Windows" and 
                    not isinstance(loop, asyncio.windows_events.ProactorEventLoop)
                )

                batch_start_time = time.time()

                try:
                    if is_windows_selector:
                        logger.debug(f"Batch {batch_id}: Using thread-based subprocess execution")
                        result = await asyncio.to_thread(
                            run_subprocess_sync, 
                            cmd, 
                            timeout
                        )
                        returncode, stdout, stderr = result
                    else:
                        process = await asyncio.create_subprocess_exec(
                            *cmd,
                            stdout=asyncio.subprocess.PIPE,
                            stderr=asyncio.subprocess.PIPE,
                            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                        )
                        
                        stdout_bytes, stderr_bytes = await asyncio.wait_for(
                            process.communicate(), 
                            timeout=timeout
                        )
                        stdout = stdout_bytes.decode(errors='ignore')
                        stderr = stderr_bytes.decode(errors='ignore')
                        returncode = process.returncode

                    # Log subprocess output for debugging
                    logger.debug(f"Batch {batch_id}: LibreOffice return code: {returncode}")
                    if stdout:
                        logger.debug(f"Batch {batch_id}: LibreOffice stdout: {stdout}")
                    if stderr:
                        logger.debug(f"Batch {batch_id}: LibreOffice stderr: {stderr}")
                        
                except asyncio.TimeoutError:
                    if not is_windows_selector:
                        process.kill()
                        await process.wait()
                    logger.error(f"Batch {batch_id} timeout after {timeout} seconds (attempt {attempt + 1}) (async)")
                    if attempt < max_retries - 1:
                        logger.info(f"Retrying batch {batch_id} (async)...")
                        await kill_libreoffice_processes_async()
                        await asyncio.sleep(2)
                        continue
                    return [], valid_files + invalid_files
                except Exception as e:
                    logger.error(f"Batch {batch_id}: Subprocess execution failed: {e}\n{traceback.format_exc()}")
                    if attempt < max_retries - 1:
                        logger.info(f"Retrying batch {batch_id} (async)...")
                        await kill_libreoffice_processes_async()
                        await asyncio.sleep(2)
                        continue
                    return [], valid_files + invalid_files

                if returncode != 0:
                    logger.error(f"Batch {batch_id} LibreOffice error (code {returncode}): {stderr}\nStdout: {stdout}")
                    if attempt < max_retries - 1:
                        logger.info(f"Retrying batch {batch_id} (async)...")
                        await kill_libreoffice_processes_async()
                        await asyncio.sleep(2)
                        continue
                    return [], valid_files + invalid_files

                batch_pdfs = []
                failed_files = []

                for docx_path_str in valid_files:
                    docx_path = Path(docx_path_str)
                    docx_name = docx_path.stem
                    temp_pdf = temp_output / f"{docx_name}.pdf"
                    final_pdf = output_dir / f"{docx_name}.pdf"

                    temp_pdf_exists = await aiofiles.os.path.exists(str(temp_pdf))
                    if temp_pdf_exists:
                        try:
                            await asyncio.to_thread(shutil.move, str(temp_pdf), str(final_pdf))
                            batch_pdfs.append(str(final_pdf))
                            logger.debug(f"Moved {temp_pdf} to {final_pdf}")
                        except Exception as e:
                            failed_files.append(str(docx_path))
                            logger.error(f"Batch {batch_id}: Failed to move {temp_pdf} to {final_pdf}: {e}\n{traceback.format_exc()}")
                    else:
                        failed_files.append(str(docx_path))
                        logger.warning(f"Batch {batch_id}: Failed to convert {docx_path.name} (async)")

                batch_time = time.time() - batch_start_time
                success_rate = len(batch_pdfs) / len(valid_files) * 100 if valid_files else 0
                logger.info(f"Batch {batch_id} result: {len(batch_pdfs)}/{len(valid_files)} successful ({success_rate:.1f}%) (async)")

                if progress_callback:
                    try:
                        update = progress_callback({
                            'type': 'batch_complete',
                            'batch_id': batch_id,
                            'successful': len(batch_pdfs),
                            'failed': len(failed_files) + len(invalid_files),
                            'success_rate': success_rate,
                            'batch_time': batch_time
                        })
                        if update:
                            pass
                    except Exception as e:
                        logger.error(f"Error in progress callback for batch_complete: {e}\n{traceback.format_exc()}")

                logger.debug(f"Batch {batch_id}: Temporary directory {temp_batch_dir} will be cleaned up automatically")
                return batch_pdfs, failed_files + invalid_files

        except Exception as e:
            logger.error(f"Batch {batch_id} conversion error (attempt {attempt + 1}) (async): {e}\n{traceback.format_exc()}")
            if attempt < max_retries - 1:
                await kill_libreoffice_processes_async()
                await asyncio.sleep(2)
                continue
            return [], valid_files + invalid_files
        finally:
            # No manual cleanup needed; context manager handles it
            if 'temp_batch_dir' in locals() and isinstance(temp_batch_dir, TemporaryDirectory):
                logger.debug(f"Batch {batch_id}: Temporary directory {temp_batch_dir.name} cleanup ensured")

    return [], valid_files + invalid_files

def run_subprocess_sync(cmd, timeout):
    """Synchronous subprocess execution for Windows SelectorEventLoop compatibility"""
    import subprocess
    import os
    
    try:
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        )
        return result.returncode, result.stdout.decode(errors='ignore'), result.stderr.decode(errors='ignore')
    except subprocess.TimeoutExpired:
        raise asyncio.TimeoutError(f"Process timed out after {timeout} seconds")
    except Exception as e:
        raise RuntimeError(f"Subprocess execution failed: {e}")

async def batch_convert_libreoffice(docx_files, output_dir, batch_size=300, progress_callback=None):
    # Log the event loop type for debugging
    loop = asyncio.get_running_loop()
    logger.debug(f"Event loop for batch_convert_libreoffice: {type(loop)}")

    if not docx_files:
        logger.warning("No DOCX files provided for conversion")
        return []

    pdf_files = []
    total_failed = []
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    await kill_libreoffice_processes_async()

    valid_docx_files, invalid_files = await validate_files_async(docx_files)
    total_failed.extend(invalid_files)
    if not valid_docx_files:
        logger.error("No valid DOCX files to convert")
        return []

    batches = [valid_docx_files[i:i + batch_size] for i in range(0, len(valid_docx_files), batch_size)]
    logger.info(f"Converting {len(valid_docx_files)} DOCX files in {len(batches)} batches (size: {batch_size}) (async)...")

    start_time = time.time()

    # Store total batches for access in batch processing
    total_batches = len(batches)

    if progress_callback:
        # Send conversion start update
        try:
            update = progress_callback({
                'type': 'conversion_start',
                'total_files': len(valid_docx_files),
                'total_batches': total_batches,
                'batch_size': batch_size
            })
            if update:
                pass
        except Exception as e:
            logger.error(f"Error in progress callback for conversion_start: {e}\n{traceback.format_exc()}")

    for batch_id, batch_list_files in enumerate(batches, 1):
        # Create a simple callback for this batch
        def batch_callback(progress_data):
            if progress_callback:
                progress_data['total_batches'] = total_batches
                try:
                    return progress_callback(progress_data)
                except Exception as e:
                    logger.error(f"Error in batch callback: {e}\n{traceback.format_exc()}")
            return None

        batch_pdfs, batch_failed_list = await convert_batch_with_retry_async(
            batch_list_files, output_dir, batch_id, progress_callback=batch_callback
        )
        pdf_files.extend(batch_pdfs)
        total_failed.extend(batch_failed_list)
        logger.debug(f"Batch {batch_id}: Returning {len(batch_pdfs)} PDFs: {batch_pdfs}")

    await kill_libreoffice_processes_async()

    total_time = time.time() - start_time
    success_rate = len(pdf_files) / len(valid_docx_files) * 100 if valid_docx_files else 0

    conversion_summary = {
        'type': 'conversion_summary',
        'total_files': len(valid_docx_files),
        'successful': len(pdf_files),
        'failed': len(total_failed),
        'success_rate': success_rate,
        'total_time': total_time,
        'total_time_formatted': format_time_duration(total_time),
        'conversion_rate': len(pdf_files) / total_time if total_time > 0 else 0,
        'total_batches': total_batches
    }

    logger.info(f"\n=== ASYNC CONVERSION SUMMARY ===")
    logger.info(f"Total files: {len(valid_docx_files)}")
    logger.info(f"Successful: {len(pdf_files)} ({success_rate:.1f}%)")
    logger.info(f"Failed: {len(total_failed)} ({100-success_rate:.1f}%)")
    logger.info(f"Time: {format_time_duration(total_time)} | Rate: {len(pdf_files)/total_time:.1f} PDFs/sec" if total_time > 0 else "Time: 0s")
    logger.debug(f"Returning {len(pdf_files)} PDFs: {pdf_files}")

    if progress_callback:
        # Send conversion summary
        try:
            update = progress_callback(conversion_summary)
            if update:
                pass
        except Exception as e:
            logger.error(f"Error in progress callback for conversion_summary: {e}\n{traceback.format_exc()}")

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