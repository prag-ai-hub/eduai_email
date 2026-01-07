import sqlite3
import threading
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent.parent / 'mailer.db'
_lock = threading.Lock()

def init_db():
    with _lock:
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT,
                category TEXT,
                subject TEXT,
                status TEXT,
                ts TEXT
            )
        ''')
        conn.commit()
        conn.close()

def log_entry(email, category, subject, status):
    ts = datetime.utcnow().isoformat()
    with _lock:
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute('INSERT INTO logs (email, category, subject, status, ts) VALUES (?,?,?,?,?)', (email, category, subject, status, ts))
        conn.commit()
        conn.close()

def get_logs(limit=200):
    with _lock:
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute('SELECT id, email, category, subject, status, ts FROM logs ORDER BY id DESC LIMIT ?', (limit,))
        rows = cur.fetchall()
        conn.close()
    return rows
