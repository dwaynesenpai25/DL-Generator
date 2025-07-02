import asyncpg
from fastapi import HTTPException
from utils.config import DATABASE_CONFIG, logger
from utils.models import UserCreate
from typing import List, Tuple, Optional
import asyncio
import json
from datetime import datetime
import uuid
import time
# Create a connection pool for asyncpg
DB_POOL = None

async def get_db_pool():
    global DB_POOL
    if DB_POOL is None:
        try:
            DB_POOL = await asyncpg.create_pool(
                user=DATABASE_CONFIG['user'],
                password=DATABASE_CONFIG['password'],
                database=DATABASE_CONFIG['database'],
                host=DATABASE_CONFIG['host'],
                port=DATABASE_CONFIG.get('port', 5432)
            )
            logger.info("Database connection pool created successfully.")
        except Exception as e:
            logger.error(f"Failed to create database connection pool: {e}")
            raise HTTPException(status_code=500, detail="Database connection pool failed")
    return DB_POOL

async def init_database():
    """Initialize PostgreSQL database with required tables asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        async with conn.transaction():
            try:
                # Existing tables
                await conn.execute('''
                    CREATE TABLE IF NOT EXISTS users (
                        id SERIAL PRIMARY KEY,
                        email VARCHAR(255) UNIQUE NOT NULL,
                        clients TEXT NOT NULL,
                        access VARCHAR(50) NOT NULL DEFAULT 'user',
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
                
                await conn.execute('''
                    CREATE TABLE IF NOT EXISTS audit_trail (
                        id SERIAL PRIMARY KEY,
                        client VARCHAR(255) NOT NULL,
                        processed_by VARCHAR(255) NOT NULL,
                        processed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        total_accounts INTEGER NOT NULL,
                        mode VARCHAR(100) NOT NULL,
                        template_folder VARCHAR(255),
                        dl_type VARCHAR(255)
                    )
                ''')
                
                await conn.execute('''
                    CREATE TABLE IF NOT EXISTS processed_accounts (
                        id SERIAL PRIMARY KEY,
                        audit_id INTEGER NOT NULL,
                        doc_code VARCHAR(255) UNIQUE NOT NULL,
                        dl_code VARCHAR(255),
                        leads_chname TEXT,
                        dl_address TEXT,
                        final_area VARCHAR(255),
                        processed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY (audit_id) REFERENCES audit_trail (id) ON DELETE CASCADE
                    )
                ''')
                
                # NEW: Sessions table for PostgreSQL
                await conn.execute('''
                    CREATE TABLE IF NOT EXISTS sessions (
                        session_id VARCHAR(255) PRIMARY KEY,
                        user_access_token TEXT NOT NULL,
                        refresh_token TEXT NOT NULL,
                        expires_at TIMESTAMP NOT NULL,
                        user_info JSONB NOT NULL,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
                
                await conn.execute('''
                    CREATE INDEX IF NOT EXISTS idx_processed_accounts_doc_code 
                    ON processed_accounts(doc_code)
                ''')
                
                await conn.execute('''
                    CREATE INDEX IF NOT EXISTS idx_sessions_expires_at 
                    ON sessions(expires_at)
                ''')
                
                logger.info("Database initialized successfully (async)")
            except Exception as e:
                logger.error(f"Failed to initialize database (async): {e}")
                raise

# Session management functions
async def save_session_to_db(session_id: str, session_data: dict):
    """Save session to PostgreSQL database asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            now = datetime.now()
            user_info_json = json.dumps(session_data["user_info"])
            expires_at = datetime.fromisoformat(session_data["expires_at"])
            
            # Check if session exists
            existing = await conn.fetchrow("SELECT 1 FROM sessions WHERE session_id = $1", session_id)
            
            if existing:
                await conn.execute('''
                    UPDATE sessions 
                    SET user_access_token = $1, refresh_token = $2, expires_at = $3, 
                        user_info = $4, updated_at = $5
                    WHERE session_id = $6
                ''', session_data["user_access_token"],
                     session_data["refresh_token"],
                     expires_at,
                     user_info_json,
                     now,
                     session_id)
            else:
                await conn.execute('''
                    INSERT INTO sessions 
                    (session_id, user_access_token, refresh_token, expires_at, user_info, created_at, updated_at)
                    VALUES ($1, $2, $3, $4, $5, $6, $7)
                ''', session_id,
                     session_data["user_access_token"],
                     session_data["refresh_token"],
                     expires_at,
                     user_info_json,
                     now,
                     now)
            logger.debug(f"Session {session_id} saved to database (async)")
            return True
        except Exception as e:
            logger.error(f"Failed to save session {session_id} to database (async): {e}")
            return False

async def load_session_from_db(session_id: str):
    """Load session from PostgreSQL database asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            row = await conn.fetchrow('''
                SELECT session_id, user_access_token, refresh_token, expires_at, user_info
                FROM sessions
                WHERE session_id = $1
            ''', session_id)
            
            if row:
                session = dict(row)
                session["user_info"] = json.loads(session["user_info"])
                session["expires_at"] = session["expires_at"].isoformat()
                return session
            
            return None
        except Exception as e:
            logger.error(f"Failed to load session {session_id} from database (async): {e}")
            return None

async def delete_session_from_db(session_id: str):
    """Delete session from PostgreSQL database asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            await conn.execute("DELETE FROM sessions WHERE session_id = $1", session_id)
            logger.info(f"Session {session_id} deleted from database (async)")
            return True
        except Exception as e:
            logger.error(f"Failed to delete session {session_id} from database (async): {e}")
            return False

async def cleanup_expired_sessions():
    """Remove expired sessions from PostgreSQL database asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            now = datetime.now()
            result = await conn.execute("DELETE FROM sessions WHERE expires_at < $1", now)
            deleted_count = int(result.split()[-1]) if result.startswith('DELETE') else 0
            
            if deleted_count > 0:
                logger.info(f"Cleaned up {deleted_count} expired sessions from database (async)")
            
            return deleted_count
        except Exception as e:
            logger.error(f"Failed to cleanup expired sessions (async): {e}")
            return 0

async def load_active_sessions():
    """Load all active sessions from PostgreSQL database asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            now = datetime.now()
            rows = await conn.fetch('''
                SELECT session_id, user_access_token, refresh_token, expires_at, user_info
                FROM sessions
                WHERE expires_at > $1
            ''', now)
            
            sessions = {}
            for row in rows:
                session_id = row["session_id"]
                session = dict(row)
                session["user_info"] = json.loads(session["user_info"])
                session["expires_at"] = session["expires_at"].isoformat()
                sessions[session_id] = session
            
            logger.info(f"Loaded {len(sessions)} active sessions from database (async)")
            return sessions
        except Exception as e:
            logger.error(f"Failed to load active sessions from database (async): {e}")
            return {}

async def get_session_info():
    """Debug function to get current session information asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            total_count = await conn.fetchval("SELECT COUNT(*) FROM sessions")
            now = datetime.now()
            active_count = await conn.fetchval("SELECT COUNT(*) FROM sessions WHERE expires_at > $1", now)
            
            return {
                "total_db_sessions": total_count,
                "active_db_sessions": active_count,
                "database_type": "PostgreSQL"
            }
        except Exception as e:
            logger.error(f"Error getting session info (async): {e}")
            return {
                "error": str(e),
                "database_type": "PostgreSQL"
            }

async def check_existing_user_session(email: str):
    """Check if user already has an active session"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            now = datetime.now()
            row = await conn.fetchrow('''
                SELECT session_id, expires_at
                FROM sessions
                WHERE user_info->>'email' = $1 AND expires_at > $2
                ORDER BY created_at DESC
                LIMIT 1
            ''', email, now)
            
            if row:
                return {
                    "has_active_session": True,
                    "session_id": row["session_id"],
                    "expires_at": row["expires_at"].isoformat()
                }
            
            return {"has_active_session": False}
        except Exception as e:
            logger.error(f"Failed to check existing session for {email}: {e}")
            return {"has_active_session": False}

async def delete_user_sessions(email: str):
    """Delete all sessions for a specific user"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            result = await conn.execute('''
                DELETE FROM sessions 
                WHERE user_info->>'email' = $1
            ''', email)
            
            deleted_count = int(result.split()[-1]) if result.startswith('DELETE') else 0
            logger.info(f"Deleted {deleted_count} sessions for user {email}")
            return deleted_count
        except Exception as e:
            logger.error(f"Failed to delete sessions for user {email}: {e}")
            return 0

async def get_user_clients_and_access(email: str) -> Tuple[List[str], Optional[str]]:
    """Get user's assigned clients and access level from database asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            row = await conn.fetchrow("SELECT clients, access FROM users WHERE email = $1", email)
            if row:
                clients_str, access = row['clients'], row['access']
                clients = clients_str.split(',') if clients_str else []
                return clients, access
            return [], None
        except Exception as e:
            logger.error(f"Failed to get user clients and access (async): {e}")
            return [], None

async def add_audit_entry(client: str, processed_by: str, total_accounts: int, mode: str, template_folder: str = None, dl_type: str = None) -> int:
    """Add entry to audit trail and return the audit ID asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            audit_id = await conn.fetchval('''
                INSERT INTO audit_trail (client, processed_by, total_accounts, mode, template_folder, dl_type)
                VALUES ($1, $2, $3, $4, $5, $6) RETURNING id
            ''', client, processed_by, total_accounts, mode, template_folder, dl_type)
            return audit_id
        except Exception as e:
            logger.error(f"Failed to add audit entry (async): {e}")
            raise

async def add_processed_accounts(audit_id: int, raw_accounts: List[Tuple]):
    """
    Add processed accounts to database asynchronously with UUID-based unique doc_code generation.
    raw_accounts: List of tuples like (audit_id, dl_code, leads_chname, dl_address, final_area)
    """
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            accounts = []
            for acc in raw_accounts:
                audit_id_val, dl_code, leads_chname, dl_address, final_area = acc

                # Inline unique doc_code generation
                date_str = datetime.now().strftime('%Y%m%d')
                timestamp_suffix = str(int(time.time() * 1000000))[-5:]
                uuid_suffix = uuid.uuid4().hex[:4].upper()
                doc_code = f"DOC-{date_str}-{timestamp_suffix}{uuid_suffix}"

                accounts.append((audit_id_val, doc_code, dl_code, leads_chname, dl_address, final_area))

            # Insert all accounts in a single transaction
            async with conn.transaction():
                successful_inserts = 0
                for account in accounts:
                    try:
                        await conn.execute(
                            '''
                            INSERT INTO processed_accounts (
                                audit_id, doc_code, dl_code, leads_chname, dl_address, final_area
                            ) VALUES ($1, $2, $3, $4, $5, $6)
                            ''',
                            *account
                        )
                        successful_inserts += 1
                    except asyncpg.UniqueViolationError:
                        logger.warning(f"Duplicate doc_code detected for {account[1]} (UUID collision - rare)")
                        
                        # Retry with a new doc_code
                        date_str = datetime.now().strftime('%Y%m%d')
                        timestamp_suffix = str(int(time.time() * 1000000))[-5:]
                        uuid_suffix = uuid.uuid4().hex[:4].upper()
                        new_doc_code = f"DOC-{date_str}-{timestamp_suffix}{uuid_suffix}"

                        new_account = (account[0], new_doc_code, account[2], account[3], account[4], account[5])
                        try:
                            await conn.execute(
                                '''
                                INSERT INTO processed_accounts (
                                    audit_id, doc_code, dl_code, leads_chname, dl_address, final_area
                                ) VALUES ($1, $2, $3, $4, $5, $6)
                                ''',
                                *new_account
                            )
                            successful_inserts += 1
                        except Exception as retry_error:
                            logger.error(f"Failed to insert after regenerating doc_code: {retry_error}")
                            continue  # Continue with the rest of the batch

                logger.debug(f"Inserted {successful_inserts}/{len(accounts)} processed accounts")

        except Exception as e:
            logger.error(f"Failed to add processed accounts (async): {e}")
            raise

async def generate_doc_code_individual(conn):
    """Generate a unique DOC code using date and 5-character UUID"""
    try:
        import uuid
        date_str = datetime.now().strftime('%Y%m%d')
        
        # Try up to 5 times to find a unique code (very unlikely to need more than 1)
        for attempt in range(5):
            # Generate 5-character UUID (uppercase for consistency)
            uuid_suffix = uuid.uuid4().hex[:5].upper()
            new_code = f"DOC-{date_str}-{uuid_suffix}"
            
            # Check if this code already exists
            existing = await conn.fetchval(
                'SELECT 1 FROM processed_accounts WHERE doc_code = $1', 
                new_code
            )
            
            if not existing:
                return new_code
        
        # If somehow all 5 attempts failed (extremely unlikely), use timestamp
        import time
        timestamp_suffix = str(int(time.time() * 1000))[-5:]
        fallback_code = f"DOC-{date_str}-{timestamp_suffix}"
        return fallback_code
        
    except Exception as e:
        logger.error(f"Failed to generate individual doc_code: {e}")
        raise HTTPException(status_code=500, detail="Failed to generate document code")

async def get_all_users():
    """Get all users from database asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            rows = await conn.fetch("SELECT email, clients, access FROM users ORDER BY email")
            users = []
            for row in rows:
                clients_str, access = row['clients'], row['access']
                clients = clients_str.split(',') if clients_str else []
                users.append({"email": row['email'], "clients": clients, "access": access})
            return users
        except Exception as e:
            logger.error(f"Failed to get users (async): {e}")
            raise HTTPException(status_code=500, detail="Failed to retrieve users")

async def create_new_user(user: UserCreate):
    """Create a new user asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            clients_str = ','.join(user.clients)
            await conn.execute("INSERT INTO users (email, clients, access) VALUES ($1, $2, $3)", 
                               user.email, clients_str, user.access)
            return {"success": True, "message": "User created successfully"}
        except asyncpg.UniqueViolationError:
            raise HTTPException(status_code=400, detail="User with this email already exists")
        except Exception as e:
            logger.error(f"Failed to create user (async): {e}")
            raise HTTPException(status_code=500, detail="Failed to create user")

async def delete_user_by_email(email: str):
    """Delete user by email asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            result = await conn.execute("DELETE FROM users WHERE email = $1", email)
            if result == 'DELETE 0':
                raise HTTPException(status_code=404, detail="User not found")
            return {"success": True, "message": "User deleted successfully"}
        except Exception as e:
            logger.error(f"Failed to delete user (async): {e}")
            raise HTTPException(status_code=500, detail="Failed to delete user")

async def update_user_by_email(email: str, user: UserCreate):
    """Update user by email asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            clients_str = ','.join(user.clients)
            result = await conn.execute("UPDATE users SET email = $1, clients = $2, access = $3 WHERE email = $4", 
                                        user.email, clients_str, user.access, email)
            if result == 'UPDATE 0':
                raise HTTPException(status_code=404, detail="User not found")
            return {"success": True, "message": "User updated successfully"}
        except asyncpg.UniqueViolationError:
            raise HTTPException(status_code=400, detail="User with this email already exists")
        except Exception as e:
            logger.error(f"Failed to update user (async): {e}")
            raise HTTPException(status_code=500, detail="Failed to update user")

async def get_audit_trail_paginated(page: int = 1, limit: int = 10):
    """Get paginated audit trail asynchronously"""
    offset = (page - 1) * limit
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            total_count_row = await conn.fetchrow('SELECT COUNT(*) as total FROM audit_trail')
            total_count = total_count_row["total"]
            
            rows = await conn.fetch('''
                SELECT id, client, processed_by, processed_at, total_accounts, mode 
                FROM audit_trail 
                ORDER BY processed_at DESC 
                LIMIT $1 OFFSET $2
            ''', limit, offset)
            
            audit_entries = [dict(row) for row in rows]
            total_pages = (total_count + limit - 1) // limit
            
            return {
                "entries": audit_entries,
                "pagination": {
                    "current_page": page,
                    "total_pages": total_pages,
                    "total_count": total_count,
                    "limit": limit,
                    "has_next": page < total_pages,
                    "has_prev": page > 1
                }
            }
        except Exception as e:
            logger.error(f"Failed to get audit trail (async): {e}")
            raise HTTPException(status_code=500, detail="Failed to retrieve audit trail")

async def get_audit_details_paginated(audit_id: int, page: int = 1, limit: int = 50):
    """Get detailed information about processed accounts for a specific audit entry asynchronously"""
    pool = await get_db_pool()
    async with pool.acquire() as conn:
        try:
            audit_entry_row = await conn.fetchrow("SELECT id, client, processed_by, processed_at, total_accounts, mode FROM audit_trail WHERE id = $1", audit_id)
            if not audit_entry_row:
                raise HTTPException(status_code=404, detail="Audit entry not found")
            audit_entry = dict(audit_entry_row)
            
            total_count_row = await conn.fetchrow("SELECT COUNT(*) as total FROM processed_accounts WHERE audit_id = $1", audit_id)
            total_count = total_count_row["total"]
            
            offset = (page - 1) * limit
            account_rows = await conn.fetch("""
                SELECT dl_code, doc_code, leads_chname, dl_address, final_area
                FROM processed_accounts
                WHERE audit_id = $1
                ORDER BY id
                LIMIT $2 OFFSET $3
            """, audit_id, limit, offset)
            
            accounts = []
            for row in account_rows:
                accounts.append({
                    "dl_code": row["dl_code"],
                    "doc_code": row["doc_code"],
                    "name": row["leads_chname"],
                    "address": row["dl_address"],
                    "area": row["final_area"]
                })
            
            total_pages = (total_count + limit - 1) // limit
            
            return {
                "audit_id": audit_id,
                "client": audit_entry["client"],
                "processed_by": audit_entry["processed_by"],
                "processed_at": audit_entry["processed_at"],
                "total_accounts": audit_entry["total_accounts"],
                "mode": audit_entry["mode"],
                "accounts": accounts,
                "pagination": {
                    "current_page": page,
                    "total_pages": total_pages,
                    "total_count": total_count,
                    "limit": limit,
                    "has_next": page < total_pages,
                    "has_prev": page > 1
                }
            }
        except Exception as e:
            logger.error(f"Failed to get audit details (async): {e}")
            raise HTTPException(status_code=500, detail="Failed to retrieve audit details")

async def close_db_pool():
    global DB_POOL
    if DB_POOL:
        await DB_POOL.close()
        DB_POOL = None
        logger.info("Database connection pool closed.")
