import sqlite3
import os
import datetime

DB_PATH = "sessions.db"

def init_db():
    """Initialize the SQLite database and create tables if they don't exist."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            phone TEXT,
            group_code TEXT,
            target_date TEXT,
            day_of_week TEXT,
            time TEXT,
            status TEXT DEFAULT 'Not Sent',
            error_detail TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS coordinator_states (
            target_date TEXT PRIMARY KEY,
            state_json TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()

def insert_session(name, phone, group_code, target_date, day_of_week, time, status="Not Sent", error_detail=""):
    """Insert a single session into the database."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO sessions (name, phone, group_code, target_date, day_of_week, time, status, error_detail)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (name, phone, group_code, target_date, day_of_week, time, status, error_detail))
    session_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return session_id

def insert_sessions_bulk(sessions_list):
    """
    Insert multiple sessions at once.
    sessions_list should be a list of dicts with keys: 
    name, phone, group_code, target_date, day_of_week, time, status, error_detail
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    rows = []
    for s in sessions_list:
        rows.append((
            s.get("name", ""),
            s.get("phone", ""),
            s.get("group_code", ""),
            s.get("target_date", ""),
            s.get("day", ""), # day_of_week
            s.get("time", ""),
            s.get("status", "Not Sent"),
            s.get("error_detail", "")
        ))
        
    cursor.executemany("""
        INSERT INTO sessions (name, phone, group_code, target_date, day_of_week, time, status, error_detail)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, rows)
    conn.commit()
    conn.close()

def get_sessions(target_date=None):
    """
    Retrieve sessions. 
    If target_date is provided, returns sessions for that date.
    If target_date is 'All Saved', returns all sessions.
    """
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    if target_date and target_date != "All Saved":
        cursor.execute("SELECT * FROM sessions WHERE target_date = ?", (target_date,))
    else:
        cursor.execute("SELECT * FROM sessions")
        
    rows = cursor.fetchall()
    conn.close()
    
    # Convert to list of dicts to match app format
    results = []
    for r in rows:
        results.append({
            "db_id": r["id"],
            "name": r["name"],
            "phone": r["phone"],
            "group_code": r["group_code"],
            "target_date": r["target_date"],
            "day": r["day_of_week"],
            "time": r["time"],
            "status": r["status"],
            "error_detail": r["error_detail"],
            "selected": False,
            "visible": True
        })
    return results

def update_session_status(db_id, status, error_detail=""):
    """Update the status of a session."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE sessions 
        SET status = ?, error_detail = ?
        WHERE id = ?
    """, (status, error_detail, db_id))
    conn.commit()
    conn.close()

def update_session_details(db_id, name, phone, group_code, day_of_week, time):
    """Update details of a session."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE sessions 
        SET name = ?, phone = ?, group_code = ?, day_of_week = ?, time = ?
        WHERE id = ?
    """, (name, phone, group_code, day_of_week, time, db_id))
    conn.commit()
    conn.close()

def delete_session(db_id):
    """Delete a session."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM sessions WHERE id = ?", (db_id,))
    conn.commit()
    conn.close()

def get_all_dates():
    """Get a list of all unique target_dates in the database, sorted descending."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT target_date FROM sessions WHERE target_date IS NOT NULL AND target_date != '' ORDER BY target_date DESC")
    rows = cursor.fetchall()
    conn.close()
    return [r[0] for r in rows]

def clear_old_sessions(before_date):
    """Delete all sessions before a certain date."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM sessions WHERE target_date < ? AND target_date != ''", (before_date,))
    deleted = cursor.rowcount
    conn.commit()
    conn.close()
    return deleted

def delete_date(target_date):
    """Delete all sessions and coordinator states for a specific date."""
    if not target_date or target_date in ("All Saved", "Temporary"):
        return 0
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM sessions WHERE target_date = ?", (target_date,))
    deleted = cursor.rowcount
    cursor.execute("DELETE FROM coordinator_states WHERE target_date = ?", (target_date,))
    deleted += cursor.rowcount
    conn.commit()
    conn.close()
    return deleted

def clear_temporary_sessions():
    """Delete all sessions marked as Temporary."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM sessions WHERE target_date = 'Temporary'")
    deleted = cursor.rowcount
    conn.commit()
    conn.close()
    return deleted


import sqlite3
import json
from db_manager import DB_PATH

def save_coordinator_state(target_date, state_dict):
    """Save the full state dictionary to the coordinator_states table as JSON."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    state_json = json.dumps(state_dict, ensure_ascii=False)
    cursor.execute("""
        INSERT INTO coordinator_states (target_date, state_json)
        VALUES (?, ?)
        ON CONFLICT(target_date) DO UPDATE SET
            state_json=excluded.state_json,
            created_at=CURRENT_TIMESTAMP
    """, (target_date, state_json))
    conn.commit()
    conn.close()

def get_coordinator_state(target_date):
    """Retrieve and parse the state dictionary for a specific date."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT state_json FROM coordinator_states WHERE target_date=?", (target_date,))
    row = cursor.fetchone()
    conn.close()
    if row and row[0]:
        return json.loads(row[0])
    return None

def get_all_coordinator_dates():
    """Return a list of all distinct target dates sorted in descending order."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT target_date FROM coordinator_states WHERE target_date != 'Temporary' ORDER BY target_date DESC")
    rows = cursor.fetchall()
    conn.close()
    return [r[0] for r in rows if r[0]]

def clear_temporary_coordinator_states():
    """Delete all states marked as Temporary."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM coordinator_states WHERE target_date = 'Temporary'")
    deleted = cursor.rowcount
    conn.commit()
    conn.close()
    return deleted
