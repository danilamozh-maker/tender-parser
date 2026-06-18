import sqlite3
from datetime import datetime
from passlib.context import CryptContext

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

def get_db():
    conn = sqlite3.connect("users.db")
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
    conn.commit()
    conn.close()

def create_user(email: str, password: str):
    # Обрезаем пароль до 72 символов для bcrypt
    password = password[:72]
    
    conn = get_db()
    hashed = pwd_context.hash(password)
    try:
        conn.execute(
            "INSERT INTO users (email, hashed_password) VALUES (?, ?)",
            (email, hashed)
        )
        conn.commit()
        conn.close()
        return True
    except sqlite3.IntegrityError:
        conn.close()
        return False

def get_user(email: str):
    conn = get_db()
    user = conn.execute(
        "SELECT * FROM users WHERE email = ?", (email,)
    ).fetchone()
    conn.close()
    return user

def verify_password(plain_password: str, hashed_password: str):
    # Обрезаем пароль до 72 символов при проверке
    plain_password = plain_password[:72]
    return pwd_context.verify(plain_password, hashed_password)
