import sqlite3
from datetime import datetime, timedelta
import secrets
import string
from passlib.hash import sha256_crypt

DB_PATH = "database.db"

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            hashed_password TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS licenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            license_key TEXT UNIQUE NOT NULL,
            user_email TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            expires_at TIMESTAMP NOT NULL,
            is_active BOOLEAN DEFAULT 1,
            activated_at TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()
    print("✅ База данных инициализирована")

def generate_license_key():
    alphabet = string.ascii_uppercase + string.digits
    key = '-'.join(''.join(secrets.choice(alphabet) for _ in range(4)) for _ in range(4))
    return key

def create_license(email=None, days_valid=30):
    conn = get_db()
    license_key = generate_license_key()
    expires_at = datetime.now() + timedelta(days=days_valid)
    try:
        conn.execute(
            "INSERT INTO licenses (license_key, user_email, expires_at, is_active) VALUES (?, ?, ?, ?)",
            (license_key, email, expires_at, 1)
        )
        conn.commit()
        conn.close()
        return license_key
    except sqlite3.IntegrityError:
        conn.close()
        return None

def verify_and_activate_license(key):
    conn = get_db()
    license = conn.execute(
        "SELECT * FROM licenses WHERE license_key = ?", (key,)
    ).fetchone()
    if not license:
        conn.close()
        return {"valid": False, "reason": "not_found"}
    if not license["is_active"]:
        conn.close()
        return {"valid": False, "reason": "already_used"}
    if datetime.now() > datetime.fromisoformat(license["expires_at"]):
        conn.close()
        return {"valid": False, "reason": "expired"}
    conn.execute(
        "UPDATE licenses SET is_active = 0, activated_at = ? WHERE license_key = ?",
        (datetime.now(), key)
    )
    conn.commit()
    conn.close()
    return {"valid": True, "expires_at": license["expires_at"]}

def verify_license(key):
    """Проверяет только срок действия, не трогает is_active"""
    conn = get_db()
    license = conn.execute(
        "SELECT * FROM licenses WHERE license_key = ?", (key,)
    ).fetchone()
    conn.close()
    if not license:
        return None
    if datetime.now() > datetime.fromisoformat(license["expires_at"]):
        return {"valid": False, "reason": "expired"}
    return {"valid": True, "expires_at": license["expires_at"], "email": license["user_email"]}

def deactivate_license(key):
    conn = get_db()
    conn.execute("UPDATE licenses SET is_active = 0 WHERE license_key = ?", (key,))
    conn.commit()
    conn.close()
