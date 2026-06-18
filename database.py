import sqlite3
from passlib.hash import sha256_crypt

DB_PATH = "users.db"

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
    conn.commit()
    conn.close()
    print("✅ База данных инициализирована")

def create_user(email: str, password: str):
    # Обрезаем пароль до 72 символов (для безопасности)
    password = password[:72]
    
    conn = get_db()
    hashed = sha256_crypt.hash(password)
    try:
        conn.execute(
            "INSERT INTO users (email, hashed_password) VALUES (?, ?)",
            (email, hashed)
        )
        conn.commit()
        conn.close()
        print(f"✅ Пользователь {email} создан")
        return True
    except sqlite3.IntegrityError:
        conn.close()
        print(f"❌ Пользователь {email} уже существует")
        return False
    except Exception as e:
        conn.close()
        print(f"❌ Ошибка при создании пользователя: {e}")
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
    return sha256_crypt.verify(plain_password, hashed_password)
