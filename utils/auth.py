import requests
import json
import uuid
import sqlite3
from datetime import datetime, timedelta
from fastapi import HTTPException, Request
from utils.config import APP_ID, APP_SECRET, TOKEN_URL, USER_ACCESS_TOKEN_URL, USER_INFO_URL, REFRESH_TOKEN_URL, logger
import os
import threading
import time
from pathlib import Path

# Global session store for authentication with thread lock
SESSION_STORE = {}
SESSION_LOCK = threading.RLock()

# SQLite database path
DB_PATH = Path("sessions.db")

def init_session_db():
    """Initialize the SQLite database for session storage"""
    try:
        conn = sqlite3.connect(str(DB_PATH))
        cursor = conn.cursor()
        
        # Create sessions table if it doesn't exist
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS sessions (
            session_id TEXT PRIMARY KEY,
            user_access_token TEXT NOT NULL,
            refresh_token TEXT NOT NULL,
            expires_at TEXT NOT NULL,
            user_info TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        ''')
        
        conn.commit()
        conn.close()
        logger.info("Session database initialized successfully")
    except Exception as e:
        logger.error(f"Failed to initialize session database: {e}")
        raise

def get_app_access_token():
    headers = {"Content-Type": "application/json; charset=utf-8"}
    data = {"app_id": APP_ID, "app_secret": APP_SECRET}
    try:
        response = requests.post(TOKEN_URL, json=data, headers=headers)
        response.raise_for_status()
        return response.json().get("tenant_access_token")
    except Exception as e:
        logger.error(f"Failed to get app access token: {e}")
        raise HTTPException(status_code=500, detail="Failed to get app access token")

def get_user_access_token(auth_code: str, tenant_access_token: str):
    headers = {
        "Content-Type": "application/json; charset=utf-8",
        "Authorization": f"Bearer {tenant_access_token}"
    }
    data = {"grant_type": "authorization_code", "code": auth_code}
    try:
        response = requests.post(USER_ACCESS_TOKEN_URL, json=data, headers=headers)
        response_data = response.json()
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

def refresh_user_access_token(refresh_token: str, tenant_access_token: str):
    headers = {
        "Content-Type": "application/json; charset=utf-8",
        "Authorization": f"Bearer {tenant_access_token}"
    }
    data = {"grant_type": "refresh_token", "refresh_token": refresh_token}
    try:
        response = requests.post(REFRESH_TOKEN_URL, json=data, headers=headers)
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

def get_user_info(user_access_token: str):
    headers = {"Authorization": f"Bearer {user_access_token}"}
    try:
        response = requests.get(USER_INFO_URL, headers=headers)
        response.raise_for_status()
        response_data = response.json()
        if "data" not in response_data:
            logger.error(f"Failed to get user information: {response_data}")
            raise HTTPException(status_code=401, detail="Failed to get user information")
        return response_data["data"]
    except Exception as e:
        logger.error(f"Error getting user info: {e}")
        raise HTTPException(status_code=500, detail="Error getting user info")

# Dependency to get current user
async def get_current_user(request: Request):
    session_id = request.cookies.get("session_id")
    
    if not session_id:
        logger.warning("No session_id cookie found")
        raise HTTPException(status_code=401, detail="Invalid or missing session")
    
    # First check in-memory cache
    with SESSION_LOCK:
        session = SESSION_STORE.get(session_id)
    
    # If not in memory, try to load from database
    if not session:
        try:
            session = load_session_from_db(session_id)
            if session:
                # Cache in memory for faster access
                with SESSION_LOCK:
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
            tenant_access_token = get_app_access_token()
            user_access_token, refresh_token, new_expires_at = refresh_user_access_token(
                session["refresh_token"], tenant_access_token
            )
            
            # Update session with new tokens
            session["user_access_token"] = user_access_token
            session["refresh_token"] = refresh_token
            session["expires_at"] = new_expires_at.isoformat()
            
            # Update in memory and database
            with SESSION_LOCK:
                SESSION_STORE[session_id] = session
            
            save_session_to_db(session_id, session)
            logger.info(f"Session {session_id} refreshed successfully")
    except Exception as e:
        logger.error(f"Error handling session {session_id}: {e}")
        # Clean up invalid session
        delete_session(session_id)
        raise HTTPException(status_code=401, detail="Session expired")
    
    return session["user_info"]

def save_session_to_db(session_id, session_data):
    """Save session to SQLite database"""
    try:
        conn = sqlite3.connect(str(DB_PATH))
        cursor = conn.cursor()
        
        now = datetime.now().isoformat()
        
        # Convert user_info to JSON string
        user_info_json = json.dumps(session_data["user_info"])
        
        # Check if session exists
        cursor.execute("SELECT 1 FROM sessions WHERE session_id = ?", (session_id,))
        exists = cursor.fetchone() is not None
        
        if exists:
            # Update existing session
            cursor.execute('''
            UPDATE sessions 
            SET user_access_token = ?, refresh_token = ?, expires_at = ?, 
                user_info = ?, updated_at = ?
            WHERE session_id = ?
            ''', (
                session_data["user_access_token"],
                session_data["refresh_token"],
                session_data["expires_at"],
                user_info_json,
                now,
                session_id
            ))
        else:
            # Insert new session
            cursor.execute('''
            INSERT INTO sessions 
            (session_id, user_access_token, refresh_token, expires_at, user_info, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                session_id,
                session_data["user_access_token"],
                session_data["refresh_token"],
                session_data["expires_at"],
                user_info_json,
                now,
                now
            ))
        
        conn.commit()
        conn.close()
        logger.debug(f"Session {session_id} saved to database")
        return True
    except Exception as e:
        logger.error(f"Failed to save session {session_id} to database: {e}")
        return False

def load_session_from_db(session_id):
    """Load session from SQLite database"""
    try:
        conn = sqlite3.connect(str(DB_PATH))
        conn.row_factory = sqlite3.Row  # This enables column access by name
        cursor = conn.cursor()
        
        cursor.execute('''
        SELECT session_id, user_access_token, refresh_token, expires_at, user_info
        FROM sessions
        WHERE session_id = ?
        ''', (session_id,))
        
        row = cursor.fetchone()
        conn.close()
        
        if row:
            # Convert row to dict and parse user_info from JSON
            session = dict(row)
            session["user_info"] = json.loads(session["user_info"])
            return session
        
        return None
    except Exception as e:
        logger.error(f"Failed to load session {session_id} from database: {e}")
        return None

def load_sessions():
    """Load all active sessions from database into memory cache"""
    global SESSION_STORE
    
    try:
        conn = sqlite3.connect(str(DB_PATH))
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # Get only non-expired sessions
        now = datetime.now().isoformat()
        cursor.execute('''
        SELECT session_id, user_access_token, refresh_token, expires_at, user_info
        FROM sessions
        WHERE expires_at > ?
        ''', (now,))
        
        rows = cursor.fetchall()
        conn.close()
        
        with SESSION_LOCK:
            SESSION_STORE.clear()
            for row in rows:
                session_id = row["session_id"]
                session = dict(row)
                session["user_info"] = json.loads(session["user_info"])
                SESSION_STORE[session_id] = session
        
        logger.info(f"Loaded {len(SESSION_STORE)} active sessions into memory")
    except Exception as e:
        logger.error(f"Failed to load sessions from database: {e}")
        with SESSION_LOCK:
            SESSION_STORE.clear()

def delete_session(session_id):
    """Delete session from both memory cache and database"""
    try:
        # Remove from memory
        with SESSION_LOCK:
            if session_id in SESSION_STORE:
                del SESSION_STORE[session_id]
        
        # Remove from database
        conn = sqlite3.connect(str(DB_PATH))
        cursor = conn.cursor()
        cursor.execute("DELETE FROM sessions WHERE session_id = ?", (session_id,))
        conn.commit()
        conn.close()
        
        logger.info(f"Session {session_id} deleted")
        return True
    except Exception as e:
        logger.error(f"Failed to delete session {session_id}: {e}")
        return False

def cleanup_expired_sessions():
    """Remove expired sessions from memory and database"""
    try:
        now = datetime.now().isoformat()
        
        # Clean database
        conn = sqlite3.connect(str(DB_PATH))
        cursor = conn.cursor()
        cursor.execute("DELETE FROM sessions WHERE expires_at < ?", (now,))
        deleted_count = cursor.rowcount
        conn.commit()
        conn.close()
        
        # Clean memory cache
        expired_sessions = []
        with SESSION_LOCK:
            for session_id, session_data in list(SESSION_STORE.items()):
                try:
                    expires_at = datetime.fromisoformat(session_data["expires_at"])
                    if datetime.now() > expires_at:
                        expired_sessions.append(session_id)
                except Exception:
                    expired_sessions.append(session_id)
            
            for session_id in expired_sessions:
                if session_id in SESSION_STORE:
                    del SESSION_STORE[session_id]
        
        if deleted_count > 0 or expired_sessions:
            logger.info(f"Cleaned up {deleted_count} expired sessions from database and {len(expired_sessions)} from memory")
        
        return deleted_count
    except Exception as e:
        logger.error(f"Failed to cleanup expired sessions: {e}")
        return 0

def get_session_info():
    """Debug function to get current session information"""
    try:
        conn = sqlite3.connect(str(DB_PATH))
        cursor = conn.cursor()
        
        # Count total sessions
        cursor.execute("SELECT COUNT(*) FROM sessions")
        db_count = cursor.fetchone()[0]
        
        # Count active sessions
        now = datetime.now().isoformat()
        cursor.execute("SELECT COUNT(*) FROM sessions WHERE expires_at > ?", (now,))
        active_count = cursor.fetchone()[0]
        
        conn.close()
        
        with SESSION_LOCK:
            memory_count = len(SESSION_STORE)
            memory_ids = list(SESSION_STORE.keys())
        
        return {
            "total_db_sessions": db_count,
            "active_db_sessions": active_count,
            "memory_sessions": memory_count,
            "memory_session_ids": memory_ids,
            "db_exists": DB_PATH.exists()
        }
    except Exception as e:
        logger.error(f"Error getting session info: {e}")
        return {
            "error": str(e),
            "db_exists": DB_PATH.exists() if DB_PATH else False
        }
