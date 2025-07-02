from fastapi import FastAPI, HTTPException, UploadFile, File, Depends, Response, Request
from fastapi.responses import FileResponse, StreamingResponse, RedirectResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pathlib import Path
import os
import uuid
import asyncio
from contextlib import asynccontextmanager
from datetime import datetime

# Import our modules
from utils.models import *
from utils.database import (
    init_database as init_pg_db, 
    get_user_clients_and_access, 
    add_audit_entry,
    get_all_users as db_get_all_users,
    create_new_user as db_create_new_user,
    delete_user_by_email as db_delete_user_by_email,
    update_user_by_email as db_update_user_by_email,
    get_audit_trail_paginated as db_get_audit_trail_paginated,
    get_audit_details_paginated as db_get_audit_details_paginated,
    close_db_pool as close_pg_pool,
    get_db_pool
)
from utils.auth import (
    get_current_user, 
    get_app_access_token, 
    get_user_access_token as auth_get_user_access_token,
    refresh_user_access_token, 
    get_user_info as auth_get_user_info,
    SESSION_STORE, 
    load_sessions, 
    cleanup_expired_sessions_wrapper as cleanup_expired_sessions, 
    delete_session, 
    save_session_to_db_wrapper as save_session_to_db,
    run_sync_request
)
from utils.document_processing import generate_pdfs_stream, process_excel_file, get_raw_file
from utils.session_management import get_user_session_state, reset_user_session_state, set_processing_mode
from utils.template_operations import (
    get_ftp_folders as template_get_ftp_folders, 
    get_ftp_templates as template_get_ftp_templates, 
    get_placeholders_for_template as template_get_placeholders,
    get_transmittal_placeholders as template_get_transmittal_placeholders,
    get_dl_types_for_folder as template_get_dl_types
)
from utils.printing import get_available_printers, print_files_for_area
from utils.config import *

# Background task for session cleanup
async def session_cleanup_task():
    while True:
        try:
            await cleanup_expired_sessions()
            await asyncio.sleep(3600)  # Run every hour
        except Exception as e:
            logger.error(f"Error in session cleanup task: {e}")
            await asyncio.sleep(300)  # Wait 5 minutes before retrying

@asynccontextmanager
async def lifespan(app: FastAPI):
    # Startup
    await get_db_pool()  # Initialize PostgreSQL pool
    await init_pg_db()
    await load_sessions()  # Load sessions from PostgreSQL
    
    cleanup_task = asyncio.create_task(session_cleanup_task())
    
    yield
    
    # Shutdown
    cleanup_task.cancel()
    try:
        await cleanup_task
    except asyncio.CancelledError:
        logger.info("Session cleanup task cancelled.")
    await close_pg_pool()  # Close PostgreSQL pool
    logger.info("Application shutdown complete.")

app = FastAPI(title="DL Generator API", lifespan=lifespan)
# app.mount("/static", StaticFiles(directory="static"), name="static")

origins = [
    # "http://localhost:8000",
    # "http://127.0.0.1:8000"
    "http://172.20.0.86:8000"
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Enhanced user auto-creation function
async def ensure_user_exists(user_email: str):
    """
    Ensures a user exists in the database. If not, creates them with default settings.
    Returns tuple of (user_clients, user_access, was_created)
    """
    try:
        user_clients, user_access = await get_user_clients_and_access(user_email)
        
        # If user doesn't exist (both clients and access are None/empty), create them
        if not user_clients and user_access is None:
            from utils.models import UserCreate
            new_user = UserCreate(
                email=user_email,
                clients=[],  # Empty clients list - will trigger no-clients modal
                access="user"  # Default to regular user
            )
            
            try:
                await db_create_new_user(new_user)
                logger.info(f"Auto-created new user: {user_email} with default settings")
                # Return the default values we just created
                return [], "user", True
            except Exception as create_error:
                logger.error(f"Failed to auto-create user {user_email}: {create_error}")
                # Return defaults even if creation fails
                return [], "user", False
        
        return user_clients, user_access, False
        
    except Exception as e:
        logger.error(f"Error in ensure_user_exists for {user_email}: {e}")
        # Return defaults if there's any error
        return [], "user", False

# Authentication endpoints
@app.get("/api/login")
async def login():
    return RedirectResponse(url=AUTH_URL)

@app.get("/api/lark_callback")
async def lark_callback(code: str, response: Response):
    logger.info("Processing Lark callback")
    try:
        tenant_access_token = await get_app_access_token()
        user_access_token, refresh_token, expires_at = await auth_get_user_access_token(code, tenant_access_token)
        user_info_data = await auth_get_user_info(user_access_token)
        
        if not user_info_data:
            raise HTTPException(status_code=401, detail="Failed to retrieve user info")
        
        user_email = user_info_data.get("email", "")

        # Ensure user exists in database (auto-create if needed)
        user_clients, user_access, was_created = await ensure_user_exists(user_email)
        
        if was_created:
            logger.info(f"New user {user_email} was auto-created during login")
        
        # Check if user already has an active session
        from utils.database import check_existing_user_session, delete_user_sessions
        existing_session = await check_existing_user_session(user_email)
        
        if existing_session["has_active_session"]:
            logger.warning(f"User {user_email} attempted to login but already has an active session")
            raise HTTPException(
                status_code=409, 
                detail="You already have an active session in another browser. Please logout first or contact an administrator."
            )
        
        session_id = str(uuid.uuid4())
        session_data = {
            "user_access_token": user_access_token,
            "refresh_token": refresh_token,
            "expires_at": expires_at.isoformat(),
            "user_info": user_info_data
        }
        
        from utils.auth import SESSION_LOCK
        async with SESSION_LOCK:
            SESSION_STORE[session_id] = session_data
        
        await save_session_to_db(session_id, session_data)
        
        logger.info(f"Created session: {session_id} for user: {user_email}")
        
        role = "User"
        if user_access == "admin":
            role = "Admin"

        response.set_cookie(
            key="session_id", value=session_id, httponly=True, 
            secure=False, samesite="Lax", max_age=86400
        )
        
        return {
            "success": True,
            "role": role,
            "username": user_info_data.get("name", user_info_data.get("email", "Unknown"))
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error in Lark callback: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Authentication failed: {str(e)}")

@app.get("/api/logout")
async def logout(request: Request, response: Response, user_info: dict = Depends(get_current_user)):
    session_id = request.cookies.get("session_id")
    user_email = user_info.get("email", "")
    
    if session_id:
        await delete_session(session_id)
        
        # Also clean up from memory store
        from utils.auth import SESSION_LOCK
        async with SESSION_LOCK:
            # Remove any sessions for this user from memory
            sessions_to_remove = []
            for sid, session_data in SESSION_STORE.items():
                if session_data.get("user_info", {}).get("email") == user_email:
                    sessions_to_remove.append(sid)
            
            for sid in sessions_to_remove:
                if sid in SESSION_STORE:
                    del SESSION_STORE[sid]
    
    response.delete_cookie("session_id", samesite="Lax", secure=False)
    
    await asyncio.to_thread(reset_user_session_state, user_email)
    
    return {"success": True, "message": "Logged out successfully"}

@app.get("/api/check_sessions")
async def check_session(request: Request):
    """Check session and redirect if already logged in"""
    session_id = request.cookies.get("session_id")
    
    if not session_id:
        logger.warning("No session_id cookie found")
        raise HTTPException(status_code=401, detail="No active session")
    
    try:
        # Try to get current user to validate session
        from utils.auth import SESSION_LOCK, load_session_from_db
        
        async with SESSION_LOCK:
            session = SESSION_STORE.get(session_id)
        
        if not session:
            session = await load_session_from_db(session_id)
            if session:
                async with SESSION_LOCK:
                    SESSION_STORE[session_id] = session
            else:
                logger.warning(f"Session ID {session_id} not found in database")
                raise HTTPException(status_code=401, detail="Invalid session")
        
        # Check if session is expired
        from datetime import datetime
        expires_at = datetime.fromisoformat(session["expires_at"])
        if datetime.now() > expires_at:
            # Try to refresh the session
            try:
                tenant_access_token = await get_app_access_token()
                user_access_token, refresh_token, new_expires_at = await refresh_user_access_token(
                    session["refresh_token"], tenant_access_token
                )
                
                session["user_access_token"] = user_access_token
                session["refresh_token"] = refresh_token
                session["expires_at"] = new_expires_at.isoformat()
                
                async with SESSION_LOCK:
                    SESSION_STORE[session_id] = session
                
                await save_session_to_db(session_id, session)
                logger.info(f"Session {session_id} refreshed successfully")
            except Exception as e:
                logger.error(f"Failed to refresh session {session_id}: {e}")
                await delete_session(session_id)
                raise HTTPException(status_code=401, detail="Session expired")
        
        user_info = session["user_info"]
        user_email = user_info.get("email", "")
        
        # Ensure user exists in database (auto-create if needed)
        user_clients, user_access, was_created = await ensure_user_exists(user_email)
        
        if was_created:
            logger.info(f"User {user_email} was auto-created during session check")

        access_level = "user"
        if user_access == "admin":
            access_level = "admin"
        
        return {
            "success": True,
            "username": user_info.get("name", user_info.get("email", "Unknown")),
            "role": access_level.title(),
            "access": access_level,
            "clients": user_clients or [],
            "avatar": user_info
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error checking session: {e}")
        raise HTTPException(status_code=401, detail="Session validation failed")

@app.get("/api/debug/sessions")
async def debug_sessions():
    from utils.auth import get_session_info_wrapper
    return await get_session_info_wrapper()

# Admin session management endpoints
@app.post("/api/admin/force_logout")
async def force_logout_user(request_data: dict, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    _, user_access = await get_user_clients_and_access(user_email)
    
    if user_access != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    target_email = request_data.get("email")
    if not target_email:
        raise HTTPException(status_code=400, detail="Email is required")
    
    from utils.database import delete_user_sessions
    deleted_count = await delete_user_sessions(target_email)
    
    # Also clean up from memory store
    from utils.auth import SESSION_LOCK
    async with SESSION_LOCK:
        sessions_to_remove = []
        for sid, session_data in SESSION_STORE.items():
            if session_data.get("user_info", {}).get("email") == target_email:
                sessions_to_remove.append(sid)
        
        for sid in sessions_to_remove:
            if sid in SESSION_STORE:
                del SESSION_STORE[sid]
    
    return {
        "success": True, 
        "message": f"Forced logout for {target_email}. Deleted {deleted_count} sessions."
    }

@app.get("/api/admin/active_sessions")
async def get_active_sessions(user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    _, user_access = await get_user_clients_and_access(user_email)
    
    if user_access != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            now = datetime.now()
            rows = await conn.fetch('''
                SELECT session_id, user_info->>'email' as email, 
                       user_info->>'name' as name, created_at, expires_at
                FROM sessions
                WHERE expires_at > $1
                ORDER BY created_at DESC
            ''', now)
            
            sessions = []
            for row in rows:
                sessions.append({
                    "session_id": row["session_id"],
                    "email": row["email"],
                    "name": row["name"],
                    "created_at": row["created_at"].isoformat(),
                    "expires_at": row["expires_at"].isoformat()
                })
            
            return {"sessions": sessions}
        except Exception as e:
            logger.error(f"Failed to get active sessions: {e}")
            raise HTTPException(status_code=500, detail="Failed to retrieve active sessions")

# Template and folder endpoints
@app.get("/api/folders")
async def get_folders(user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    user_clients, user_access = await get_user_clients_and_access(user_email)
    return await asyncio.to_thread(template_get_ftp_folders, user_clients, user_access)

@app.get("/api/all_folders")
async def get_all_folders(user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    _, user_access = await get_user_clients_and_access(user_email)
    
    return await asyncio.to_thread(template_get_ftp_folders, [], "modal")

@app.post("/api/dl_types")
async def get_dl_types(request_data: FolderRequest, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    user_clients, user_access = await get_user_clients_and_access(user_email)
    return await asyncio.to_thread(template_get_dl_types, request_data.folder, user_clients, user_access)

@app.post("/api/templates")
async def get_templates(request_data: FolderRequest, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    user_clients, user_access = await get_user_clients_and_access(user_email)
    return await asyncio.to_thread(template_get_ftp_templates, request_data.folder, user_clients, user_access)

@app.post("/api/placeholders")
async def get_placeholders(request_data: TemplateRequest, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    user_clients, user_access = await get_user_clients_and_access(user_email)
    return await asyncio.to_thread(template_get_placeholders, request_data, user_email, user_clients, user_access)

@app.get("/api/transmittal_placeholders")
async def get_transmittal_placeholders_endpoint(folder: str = None, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    user_clients, user_access = await get_user_clients_and_access(user_email)
    return await asyncio.to_thread(template_get_transmittal_placeholders, user_email, user_clients, user_access, folder)

# Processing endpoints
@app.post("/api/upload_excel")
async def upload_excel(file: UploadFile = File(...)):
    return await asyncio.to_thread(process_excel_file, file)

@app.post("/api/set_mode")
async def set_mode(request_data: ModeRequest, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    return await asyncio.to_thread(set_processing_mode, request_data.mode, user_email)

@app.post("/api/set_output_format")
async def set_output_format(request_data: OutputFormatRequest, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    session_state = await asyncio.to_thread(get_user_session_state, user_email)
    session_state['output_format'] = request_data.format
    return {"success": True, "format": request_data.format}

@app.post("/api/generate_pdfs")
async def generate_pdfs(file: UploadFile = File(...), user_info: dict = Depends(get_current_user)):
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="Invalid file format. Please upload an .xlsx file")
    
    df = await asyncio.to_thread(get_raw_file, file)
    if df.empty:
        raise HTTPException(status_code=500, detail="Failed to read Excel file or file is empty.")
    
    return StreamingResponse(
        generate_pdfs_stream(file, df, user_info),
        media_type="application/json"
    )

@app.get("/api/download_zip")
async def download_zip(user_info: dict = Depends(get_current_user)):
    from utils.document_processing import get_zip_for_download
    user_email = user_info.get("email", "")
    return await asyncio.to_thread(get_zip_for_download, user_email)

# Printing endpoints
@app.get("/api/printers")
async def get_printers(user_info: dict = Depends(get_current_user)):
    printers = await asyncio.to_thread(get_available_printers)
    return {"printers": printers}

@app.get("/api/print_files/{area}")
async def print_files(area: str, printer: str = None, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    return await asyncio.to_thread(print_files_for_area, area, printer, user_email)

# User Management endpoints
@app.get("/api/users")
async def get_users(user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    _, user_access = await get_user_clients_and_access(user_email)
    if user_access != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    return await db_get_all_users()

@app.post("/api/users")
async def create_user(user: UserCreate, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    _, user_access = await get_user_clients_and_access(user_email)
    if user_access != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    return await db_create_new_user(user)

@app.delete("/api/users/{email_param}")
async def delete_user(email_param: str, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    _, user_access = await get_user_clients_and_access(user_email)
    if user_access != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    return await db_delete_user_by_email(email_param)

@app.put("/api/users/{email_param}")
async def update_user(email_param: str, user: UserCreate, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    _, user_access = await get_user_clients_and_access(user_email)
    if user_access != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    return await db_update_user_by_email(email_param, user)

# Audit Trail endpoints
@app.get("/api/audit_trail")
async def get_audit_trail(page: int = 1, limit: int = 10, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    _, user_access = await get_user_clients_and_access(user_email)
    if user_access != "admin":
         raise HTTPException(status_code=403, detail="Admin access required for audit trail")
    return await db_get_audit_trail_paginated(page, limit)

@app.get("/api/audit_details/{audit_id}")
async def get_audit_details(audit_id: int, page: int = 1, limit: int = 50, user_info: dict = Depends(get_current_user)):
    user_email = user_info.get("email", "")
    _, user_access = await get_user_clients_and_access(user_email)
    if user_access != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    return await db_get_audit_details_paginated(audit_id, page, limit)

# Cleanup endpoint
@app.post("/api/cleanup")
async def cleanup(user_info: dict = Depends(get_current_user)):
    try:
        user_email = user_info.get("email", "")
        await asyncio.to_thread(reset_user_session_state, user_email)
        return {"success": True, "message": "Files cleaned up and session reset successfully"}
    except Exception as e:
        logger.error(f"Cleanup failed for user {user_email}: {e}", exc_info=True)
        error_detail = str(e) if e else "Unknown cleanup error"
        raise HTTPException(status_code=500, detail=f"Cleanup operation failed: {error_detail}")

# Root redirect endpoint
@app.get("/")
async def serve_index(request: Request):
    """Serve index.html and handle session redirects"""
    session_id = request.cookies.get("session_id")
    
    if session_id:
        try:
            # Check if user has valid session
            from utils.auth import SESSION_LOCK, load_session_from_db
            from datetime import datetime
            
            async with SESSION_LOCK:
                session = SESSION_STORE.get(session_id)
            
            if not session:
                session = await load_session_from_db(session_id)
                if session:
                    async with SESSION_LOCK:
                        SESSION_STORE[session_id] = session
            
            if session:
                # Check if session is still valid
                expires_at = datetime.fromisoformat(session["expires_at"])
                if datetime.now() <= expires_at:
                    # User has valid session, they'll be automatically logged in by frontend
                    logger.info(f"User with valid session {session_id} accessing root")
                else:
                    # Session expired, clean it up
                    await delete_session(session_id)
                    logger.info(f"Expired session {session_id} cleaned up")
        except Exception as e:
            logger.error(f"Error checking session on root access: {e}")
            # Continue to serve index.html even if session check fails
    
    return FileResponse("index.html")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=5000, reload=True)
