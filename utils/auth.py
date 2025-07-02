import requests
import json
import uuid
from datetime import datetime, timedelta
from fastapi import HTTPException, Request
from utils.config import APP_ID, APP_SECRET, TOKEN_URL, USER_ACCESS_TOKEN_URL, USER_INFO_URL, REFRESH_TOKEN_URL, logger
import asyncio
from utils.database import (
    save_session_to_db, 
    load_session_from_db, 
    delete_session_from_db, 
    cleanup_expired_sessions,
    load_active_sessions,
    get_session_info
)

# Global session store for authentication with asyncio Lock
SESSION_STORE = {}
SESSION_LOCK = asyncio.Lock()

# Synchronous HTTP requests wrapped for asyncio
async def run_sync_request(func, *args, **kwargs):
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, lambda: func(*args, **kwargs))

async def get_app_access_token():
    headers = {"Content-Type": "application/json; charset=utf-8"}
    data = {"app_id": APP_ID, "app_secret": APP_SECRET}
    try:
        response = await run_sync_request(requests.post, TOKEN_URL, json=data, headers=headers)
        response.raise_for_status()
        return response.json().get("tenant_access_token")
    except Exception as e:
        logger.error(f"Failed to get app access token: {e}")
        raise HTTPException(status_code=500, detail="Failed to get app access token")

async def get_user_access_token(auth_code: str, tenant_access_token: str):
    headers = {
        "Content-Type": "application/json; charset=utf-8",
        "Authorization": f"Bearer {tenant_access_token}"
    }
    data = {"grant_type": "authorization_code", "code": auth_code}
    try:
        response = await run_sync_request(requests.post, USER_ACCESS_TOKEN_URL, json=data, headers=headers)
        response_data = response.json()
        print(response_data)
        if "data" not in response_data or "access_token" not in response_data["data"]:
            logger.error(f"Failed to get user access token: {response_data}")
            raise HTTPException(status_code=401, detail="Failed to get user access token")
        expires_in = response_data["data"]["expires_in"]
        expires_at = datetime.now() + timedelta(seconds=expires_in)
        return (
            response_data["data"]["access_token"],
            response_data["data"]["refresh_token"],
            expires_at
        )
    except Exception as e:
        logger.error(f"Error getting user access token: {e}")
        raise HTTPException(status_code=500, detail="Error getting user access token")

async def refresh_user_access_token(refresh_token: str, tenant_access_token: str):
    headers = {
        "Content-Type": "application/json; charset=utf-8",
        "Authorization": f"Bearer {tenant_access_token}"
    }
    data = {"grant_type": "refresh_token", "refresh_token": refresh_token}
    try:
        response = await run_sync_request(requests.post, REFRESH_TOKEN_URL, json=data, headers=headers)
        response_data = response.json()
        if response_data.get("code") != 0:
            logger.error(f"Failed to refresh access token: {response_data.get('message', 'Unknown error')}")
            raise HTTPException(status_code=401, detail="Failed to refresh access token")
        expires_in = response_data["data"]["expires_in"]
        expires_at = datetime.now() + timedelta(seconds=expires_in)
        return (
            response_data["data"]["access_token"],
            response_data["data"]["refresh_token"],
            expires_at
        )
    except Exception as e:
        logger.error(f"Error refreshing access token: {e}")
        raise HTTPException(status_code=500, detail="Error refreshing access token")

async def get_user_info(user_access_token: str):
    headers = {"Authorization": f"Bearer {user_access_token}"}
    try:
        response = await run_sync_request(requests.get, USER_INFO_URL, headers=headers)
        response.raise_for_status()
        response_data = response.json()
        if "data" not in response_data:
            logger.error(f"Failed to get user information: {response_data}")
            raise HTTPException(status_code=401, detail="Failed to get user information")
        return response_data["data"]
    except Exception as e:
        logger.error(f"Error getting user info: {e}")
        raise HTTPException(status_code=500, detail="Error getting user info")

async def get_current_user(request: Request):
    session_id = request.cookies.get("session_id")
    
    if not session_id:
        logger.warning("No session_id cookie found")
        raise HTTPException(status_code=401, detail="Invalid or missing session")
    
    async with SESSION_LOCK:
        session = SESSION_STORE.get(session_id)
    
    if not session:
        try:
            session = await load_session_from_db(session_id)
            if session:
                async with SESSION_LOCK:
                    SESSION_STORE[session_id] = session
            else:
                logger.warning(f"Session ID {session_id} not found in database")
                raise HTTPException(status_code=401, detail="Invalid or missing session")
        except Exception as e:
            logger.error(f"Error loading session {session_id} from database: {e}")
            raise HTTPException(status_code=401, detail="Session error")
    
    try:
        expires_at = datetime.fromisoformat(session["expires_at"])
        if datetime.now() > expires_at:
            logger.info(f"Session {session_id} expired, attempting refresh")
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
        logger.error(f"Error handling session {session_id}: {e}")
        await delete_session(session_id)
        raise HTTPException(status_code=401, detail="Session expired or invalid")
    
    return session["user_info"]

async def save_session_to_db_wrapper(session_id, session_data):
    """Wrapper function for backward compatibility"""
    return await save_session_to_db(session_id, session_data)

async def load_session_from_db_wrapper(session_id):
    """Wrapper function for backward compatibility"""
    return await load_session_from_db(session_id)

async def load_sessions():
    """Load all active sessions from database into memory cache asynchronously"""
    global SESSION_STORE
    
    try:
        sessions = await load_active_sessions()
        async with SESSION_LOCK:
            SESSION_STORE.clear()
            SESSION_STORE.update(sessions)
        
        logger.info(f"Loaded {len(SESSION_STORE)} active sessions into memory (async)")
    except Exception as e:
        logger.error(f"Failed to load sessions from database (async): {e}")
        async with SESSION_LOCK:
            SESSION_STORE.clear()

async def delete_session(session_id):
    """Delete session from both memory cache and database asynchronously"""
    try:
        async with SESSION_LOCK:
            if session_id in SESSION_STORE:
                del SESSION_STORE[session_id]
        
        await delete_session_from_db(session_id)
        logger.info(f"Session {session_id} deleted (async)")
        return True
    except Exception as e:
        logger.error(f"Failed to delete session {session_id} (async): {e}")
        return False

async def cleanup_expired_sessions_wrapper():
    """Remove expired sessions from memory and database asynchronously"""
    try:
        # Clean up from database
        deleted_count = await cleanup_expired_sessions()
        
        # Clean up from memory
        expired_sessions_mem = []
        async with SESSION_LOCK:
            for session_id, session_data in list(SESSION_STORE.items()):
                try:
                    expires_at = datetime.fromisoformat(session_data["expires_at"])
                    if datetime.now() > expires_at:
                        expired_sessions_mem.append(session_id)
                except Exception:
                    expired_sessions_mem.append(session_id)
            
            for session_id in expired_sessions_mem:
                if session_id in SESSION_STORE:
                    del SESSION_STORE[session_id]
        
        if deleted_count > 0 or expired_sessions_mem:
            logger.info(f"Cleaned up {deleted_count} expired sessions from database and {len(expired_sessions_mem)} from memory (async)")
        
        return deleted_count
    except Exception as e:
        logger.error(f"Failed to cleanup expired sessions (async): {e}")
        return 0

async def get_session_info_wrapper():
    """Debug function to get current session information asynchronously"""
    try:
        db_info = await get_session_info()
        
        async with SESSION_LOCK:
            memory_count = len(SESSION_STORE)
            memory_ids = list(SESSION_STORE.keys())
        
        return {
            **db_info,
            "memory_sessions": memory_count,
            "memory_session_ids": memory_ids
        }
    except Exception as e:
        logger.error(f"Error getting session info (async): {e}")
        return {
            "error": str(e),
            "database_type": "PostgreSQL"
        }
