from fastapi import FastAPI, HTTPException, UploadFile, File, Depends, Response, Request
from fastapi.responses import FileResponse, StreamingResponse, RedirectResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pathlib import Path
import os
import uuid
import asyncio
from contextlib import asynccontextmanager

# Import our modules
from utils.models import *
from utils.database import init_database, get_user_clients_and_access, add_audit_entry
from utils.auth import get_current_user, get_app_access_token, get_user_access_token, refresh_user_access_token, get_user_info, SESSION_STORE, init_session_db, load_sessions, cleanup_expired_sessions, delete_session, save_session_to_db
from utils.document_processing import generate_pdfs_stream
from utils.session_management import get_user_session_state, reset_user_session_state
from utils.template_operations import get_transmittal_placeholders, get_ftp_folders, get_ftp_templates, get_placeholders_for_template, get_transmittal_placeholders
from utils.printing import get_available_printers, print_files_for_area
from utils.config import *

# Background task for session cleanup
async def session_cleanup_task():
    while True:
        try:
            cleanup_expired_sessions()
            await asyncio.sleep(3600)  # Run every hour
        except Exception as e:
            logger.error(f"Error in session cleanup task: {e}")
            await asyncio.sleep(300)  # Wait 5 minutes before retrying

@asynccontextmanager
async def lifespan(app: FastAPI):
    # Startup
    init_database()
    init_session_db()  # Initialize SQLite session database
    load_sessions()
    
    # Start background task for session cleanup
    cleanup_task = asyncio.create_task(session_cleanup_task())
    
    yield
    
    # Shutdown
    cleanup_task.cancel()
    try:
        await cleanup_task
    except asyncio.CancelledError:
        pass

# Initialize FastAPI app with lifespan
app = FastAPI(title="DL Generator API", lifespan=lifespan)
app.mount("/static", StaticFiles(directory="static"), name="static")

origins = ["http://localhost:8000"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Authentication endpoints
@app.get("/api/login")
async def login():
    return RedirectResponse(url=AUTH_URL)

@app.get("/api/lark_callback")
async def lark_callback(code: str, response: Response):
    logger.info("Processing Lark callback")
    try:
        tenant_access_token = get_app_access_token()
        user_access_token, refresh_token, expires_at = get_user_access_token(code, tenant_access_token)
        user_info = get_user_info(user_access_token)
        
        if not user_info:
            raise HTTPException(status_code=401, detail="Failed to retrieve user info")
        
        # Generate session ID
        session_id = str(uuid.uuid4())
        
        # Create session data
        session_data = {
            "user_access_token": user_access_token,
            "refresh_token": refresh_token,
            "expires_at": expires_at.isoformat(),
            "user_info": user_info
        }
        
        # Store session in memory cache
        from utils.auth import SESSION_LOCK
        with SESSION_LOCK:
            SESSION_STORE[session_id] = session_data
        
        # Save to database
        save_session_to_db(session_id, session_data)
        
        logger.info(f"Created session: {session_id}")
        logger.debug(f"Session data keys: {list(session_data.keys())}")
        
        # Get user access level from database
        user_email = user_info.get("email", "")
        user_clients, user_access = get_user_clients_and_access(user_email)
        
        if user_access == "admin":
            role = "Admin"
        elif user_access == "user":
            role = "User"
        else:
            # Default logic for users not in database
            role = "Admin" if user_email.endswith("@spmadridlaw.com") else "User"
        
        # Set session cookie with proper security settings
        response.set_cookie(
            key="session_id", 
            value=session_id, 
            httponly=True, 
            secure=True, 
            samesite="lax",
            max_age=86400  # 24 hours
        )
        
        return {
            "success": True,
            "role": role,
            "username": user_info.get("name", user_info.get("email", "Unknown"))
        }
        
    except Exception as e:
        logger.error(f"Error in Lark callback: {e}")
        raise HTTPException(status_code=500, detail="Authentication failed")

@app.get("/api/logout")
async def logout(request: Request, response: Response, user_info: dict = Depends(get_current_user)):
    session_id = request.cookies.get("session_id")
    
    if session_id:
        # Delete session from both memory and database
        delete_session(session_id)
    
    response.delete_cookie("session_id")
    
    # Reset user-specific session state on logout
    user_email = user_info.get("email", "")
    reset_user_session_state(user_email)
    
    return {"success": True, "message": "Logged out successfully"}

@app.get("/api/check_sessions")
async def check_session(user_info: dict = Depends(get_current_user)):
    logger.info("Checking session")
    user_email = user_info.get("email", "")
    user_clients, user_access = get_user_clients_and_access(user_email)

    # Determine access level
    if user_access == "admin":
        access_level = "admin"
    elif user_access == "user":
        access_level = "user"
    else:
        # Default logic for users not in database
        access_level = "admin" if user_email.endswith("@spmadridlaw.com") else "user"
    
    return {
        "success": True,
        "username": user_info.get("name", user_info.get("email", "Unknown")),
        "role": access_level.title(),
        "access": access_level,
        "clients": user_clients,
        "avatar": user_info
    }

# Debug endpoint to check session status
@app.get("/api/debug/sessions")
async def debug_sessions():
    from utils.auth import get_session_info
    return get_session_info()

# Template and folder endpoints
@app.get("/api/folders")
async def get_folders(user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    user_clients, user_access = get_user_clients_and_access(user_email)
    return get_ftp_folders(user_clients, user_access)

@app.get("/api/all_folders")
async def get_all_folders(user_info: dict = Depends(get_current_user)):
    """Get all available folders for admin users (used in user management modal)"""
    user_email = user_info.get("email", "")
    user_clients, user_access = get_user_clients_and_access(user_email)
    
    if user_access != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    return get_ftp_folders([], "modal")

@app.post("/api/dl_types")
async def get_dl_types(request: FolderRequest, user_info: dict = Depends(get_current_user)):
    from utils.template_operations import get_dl_types_for_folder
    user_email = user_info.get("email", "")
    user_clients, user_access = get_user_clients_and_access(user_email)
    return get_dl_types_for_folder(request.folder, user_clients, user_access)

@app.post("/api/templates")
async def get_templates(request: FolderRequest, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    user_clients, user_access = get_user_clients_and_access(user_email)
    return get_ftp_templates(request.folder, user_clients, user_access)

@app.post("/api/placeholders")
async def get_placeholders(request: TemplateRequest, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    user_clients, user_access = get_user_clients_and_access(user_email)
    return get_placeholders_for_template(request, user_email, user_clients, user_access)

@app.get("/api/transmittal_placeholders")
async def get_transmittal_placeholders_endpoint(folder: str = None, user_info: dict = Depends(get_current_user)):
    """Get placeholders for transmittal template"""
    user_email = user_info.get("email", "")
    user_clients, user_access = get_user_clients_and_access(user_email)
    return get_transmittal_placeholders(user_email, user_clients, user_access, folder)

# Processing endpoints
@app.post("/api/upload_excel")
async def upload_excel(file: UploadFile = File(...)):
    from utils.document_processing import process_excel_file
    return process_excel_file(file)

@app.post("/api/set_mode")
async def set_mode(request: ModeRequest, user_info: dict = Depends(get_current_user)):
    from utils.session_management import set_processing_mode
    user_email = user_info.get("email", "")
    return set_processing_mode(request.mode, user_email)

@app.post("/api/set_output_format")
async def set_output_format(request: OutputFormatRequest, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    session_state = get_user_session_state(user_email)
    session_state['output_format'] = request.format
    
    return {
        "success": True,
        "format": request.format
    }

@app.post("/api/generate_pdfs")
async def generate_pdfs(file: UploadFile = File(...), user_info: dict = Depends(get_current_user)):
    from utils.document_processing import get_raw_file
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="Invalid file format. Please upload an .xlsx file")
    df = get_raw_file(file)
    if df.empty:
        raise HTTPException(status_code=500, detail="Failed to read Excel file")
    return StreamingResponse(
        generate_pdfs_stream(file, df, user_info),
        media_type="application/json"
    )

@app.get("/api/download_zip")
async def download_zip(user_info: dict = Depends(get_current_user)):
    from utils.document_processing import get_zip_for_download
    user_email = user_info.get("email", "")
    return get_zip_for_download(user_email)

# Printing endpoints
@app.get("/api/printers")
async def get_printers(user_info: dict = Depends(get_current_user)):
    """Get list of available printers"""
    printers = get_available_printers()
    return {"printers": printers}

@app.get("/api/print_files/{area}")
async def print_files(area: str, printer: str = None, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    return print_files_for_area(area, printer, user_email)

# User Management endpoints
@app.get("/api/users")
async def get_users(user_info: dict = Depends(get_current_user)):
    from utils.database import get_all_users
    user_email = user_info.get("email", "")
    user_clients, user_access = get_user_clients_and_access(user_email)
    
    if user_access != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    return get_all_users()

@app.post("/api/users")
async def create_user(user: UserCreate, user_info: dict = Depends(get_current_user)):
    from utils.database import create_new_user
    user_email = user_info.get("email", "")
    user_clients, user_access = get_user_clients_and_access(user_email)
    
    if user_access != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    return create_new_user(user)

@app.delete("/api/users/{email}")
async def delete_user(email: str, user_info: dict = Depends(get_current_user)):
    from utils.database import delete_user_by_email
    user_email = user_info.get("email", "")
    user_clients, user_access = get_user_clients_and_access(user_email)
    
    if user_access != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    return delete_user_by_email(email)

@app.put("/api/users/{email}")
async def update_user(email: str, user: UserCreate, user_info: dict = Depends(get_current_user)):
    from utils.database import update_user_by_email
    user_email = user_info.get("email", "")
    user_clients, user_access = get_user_clients_and_access(user_email)
    
    if user_access != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    return update_user_by_email(email, user)

# Audit Trail endpoints
@app.get("/api/audit_trail")
async def get_audit_trail(page: int = 1, limit: int = 10, user_info: dict = Depends(get_current_user)):
    from utils.database import get_audit_trail_paginated
    return get_audit_trail_paginated(page, limit)

@app.get("/api/audit_details/{audit_id}")
async def get_audit_details(audit_id: int, page: int = 1, limit: int = 50, user_info: dict = Depends(get_current_user)):
    from utils.database import get_audit_details_paginated
    user_email = user_info.get("email", "")
    user_clients, user_access = get_user_clients_and_access(user_email)
    
    # Only admin can access audit details
    if user_access != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    return get_audit_details_paginated(audit_id, page, limit)

# Cleanup endpoint
@app.post("/api/cleanup")
async def cleanup(user_info: dict = Depends(get_current_user)):
    try:
        user_email = user_info.get("email", "")
        reset_user_session_state(user_email)
        return {"success": True, "message": "Files cleaned up and session reset successfully"}
    except Exception as e:
        logger.error(f"Cleanup failed for user {user_email}: {e}")
        return {"success": False, "detail": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=5000)
