import psycopg2
from psycopg2.extras import RealDictCursor
from psycopg2 import sql
from fastapi import HTTPException
from utils.config import DATABASE_CONFIG, logger
from utils.models import UserCreate
from typing import List, Tuple, Optional
import uuid

def get_db_connection():
    """Get database connection"""
    try:
        conn = psycopg2.connect(**DATABASE_CONFIG)
        return conn
    except Exception as e:
        logger.error(f"Failed to connect to database: {e}")
        raise HTTPException(status_code=500, detail="Database connection failed")

def init_database():
    """Initialize PostgreSQL database with required tables"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        # Create users table with multiple clients support
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                email VARCHAR(255) UNIQUE NOT NULL,
                clients TEXT NOT NULL,
                access VARCHAR(50) NOT NULL DEFAULT 'user',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Create audit_trail table
        cursor.execute('''
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
        
        # Create processed_accounts table to store actual account data
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS processed_accounts (
                id SERIAL PRIMARY KEY,
                audit_id INTEGER NOT NULL,
                dl_code VARCHAR(255),
                leads_chname TEXT,
                dl_address TEXT,
                final_area VARCHAR(255),
                processed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (audit_id) REFERENCES audit_trail (id) ON DELETE CASCADE
            )
        ''')
        
        conn.commit()
        logger.info("Database initialized successfully")
    except Exception as e:
        logger.error(f"Failed to initialize database: {e}")
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()

def get_user_clients_and_access(email: str) -> Tuple[List[str], Optional[str]]:
    """Get user's assigned clients and access level from database"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("SELECT clients, access FROM users WHERE email = %s", (email,))
        result = cursor.fetchone()
        if result:
            clients_str, access = result
            clients = clients_str.split(',') if clients_str else []
            return clients, access
        return [], None
    except Exception as e:
        logger.error(f"Failed to get user clients and access: {e}")
        return [], None
    finally:
        cursor.close()
        conn.close()

def add_audit_entry(client: str, processed_by: str, total_accounts: int, mode: str, template_folder: str = None, dl_type: str = None) -> int:
    """Add entry to audit trail and return the audit ID"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute('''
            INSERT INTO audit_trail (client, processed_by, total_accounts, mode, template_folder, dl_type)
            VALUES (%s, %s, %s, %s, %s, %s) RETURNING id
        ''', (client, processed_by, total_accounts, mode, template_folder, dl_type))
        audit_id = cursor.fetchone()[0]
        conn.commit()
        return audit_id
    except Exception as e:
        logger.error(f"Failed to add audit entry: {e}")
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()

def add_processed_accounts(audit_id: int, accounts: List[Tuple]):
    """Add processed accounts to database"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        cursor.executemany(
            "INSERT INTO processed_accounts (audit_id, dl_code, leads_chname, dl_address, final_area) VALUES (%s, %s, %s, %s, %s)",
            accounts
        )
        conn.commit()
    except Exception as e:
        logger.error(f"Failed to add processed accounts: {e}")
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()

def get_all_users():
    """Get all users from database"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("SELECT email, clients, access FROM users ORDER BY email")
        users = []
        for row in cursor.fetchall():
            email, clients_str, access = row
            clients = clients_str.split(',') if clients_str else []
            users.append({"email": email, "clients": clients, "access": access})
        return users
    except Exception as e:
        logger.error(f"Failed to get users: {e}")
        raise HTTPException(status_code=500, detail="Failed to retrieve users")
    finally:
        cursor.close()
        conn.close()

def create_new_user(user: UserCreate):
    """Create a new user"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        clients_str = ','.join(user.clients)
        cursor.execute("INSERT INTO users (email, clients, access) VALUES (%s, %s, %s)", 
                      (user.email, clients_str, user.access))
        conn.commit()
        return {"success": True, "message": "User created successfully"}
    except psycopg2.IntegrityError:
        conn.rollback()
        raise HTTPException(status_code=400, detail="User with this email already exists")
    except Exception as e:
        logger.error(f"Failed to create user: {e}")
        conn.rollback()
        raise HTTPException(status_code=500, detail="Failed to create user")
    finally:
        cursor.close()
        conn.close()

def delete_user_by_email(email: str):
    """Delete user by email"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("DELETE FROM users WHERE email = %s", (email,))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="User not found")
        conn.commit()
        return {"success": True, "message": "User deleted successfully"}
    except Exception as e:
        logger.error(f"Failed to delete user: {e}")
        conn.rollback()
        raise HTTPException(status_code=500, detail="Failed to delete user")
    finally:
        cursor.close()
        conn.close()

def update_user_by_email(email: str, user: UserCreate):
    """Update user by email"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        clients_str = ','.join(user.clients)
        cursor.execute("UPDATE users SET email = %s, clients = %s, access = %s WHERE email = %s", 
                      (user.email, clients_str, user.access, email))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="User not found")
        conn.commit()
        return {"success": True, "message": "User updated successfully"}
    except psycopg2.IntegrityError:
        conn.rollback()
        raise HTTPException(status_code=400, detail="User with this email already exists")
    except Exception as e:
        logger.error(f"Failed to update user: {e}")
        conn.rollback()
        raise HTTPException(status_code=500, detail="Failed to update user")
    finally:
        cursor.close()
        conn.close()

def get_audit_trail_paginated(page: int = 1, limit: int = 10):
    """Get paginated audit trail"""
    offset = (page - 1) * limit
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    
    try:
        # Get total count
        cursor.execute('SELECT COUNT(*) as total FROM audit_trail')
        total_count = cursor.fetchone()["total"]
        
        # Get paginated results
        cursor.execute('''
            SELECT id, client, processed_by, processed_at, total_accounts, mode 
            FROM audit_trail 
            ORDER BY processed_at DESC 
            LIMIT %s OFFSET %s
        ''', (limit, offset))
        
        audit_entries = []
        for row in cursor.fetchall():
            audit_entries.append({
                "id": row["id"],
                "client": row["client"],
                "processed_by": row["processed_by"],
                "processed_at": row["processed_at"],
                "total_accounts": row["total_accounts"],
                "mode": row["mode"]
            })
        
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
        logger.error(f"Failed to get audit trail: {e}")
        raise HTTPException(status_code=500, detail="Failed to retrieve audit trail")
    finally:
        cursor.close()
        conn.close()

def get_audit_details_paginated(audit_id: int, page: int = 1, limit: int = 50):
    """Get detailed information about processed accounts for a specific audit entry"""
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    
    try:
        # First, verify the audit entry exists and get its details
        cursor.execute("SELECT id, client, processed_by, processed_at, total_accounts, mode FROM audit_trail WHERE id = %s", (audit_id,))
        audit_entry = cursor.fetchone()
        
        if not audit_entry:
            raise HTTPException(status_code=404, detail="Audit entry not found")
        
        # Get total count of accounts
        cursor.execute("SELECT COUNT(*) as total FROM processed_accounts WHERE audit_id = %s", (audit_id,))
        total_count = cursor.fetchone()["total"]
        
        # Get paginated accounts
        offset = (page - 1) * limit
        cursor.execute("""
            SELECT dl_code, leads_chname, dl_address, final_area
            FROM processed_accounts
            WHERE audit_id = %s
            ORDER BY id
            LIMIT %s OFFSET %s
        """, (audit_id, limit, offset))
        
        accounts = []
        for row in cursor.fetchall():
            accounts.append({
                "dl_code": row["dl_code"],
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
        logger.error(f"Failed to get audit details: {e}")
        raise HTTPException(status_code=500, detail="Failed to retrieve audit details")
    finally:
        cursor.close()
        conn.close()
