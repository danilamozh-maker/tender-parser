import asyncpg
import os
import json
from datetime import datetime, timedelta
import secrets
import string
from passlib.hash import sha256_crypt

DB_HOST = os.getenv("DB_HOST", "localhost")
DB_PORT = os.getenv("DB_PORT", "5432")
DB_NAME = os.getenv("DB_NAME", "default_db")
DB_USER = os.getenv("DB_USER", "gen_user")
DB_PASSWORD = os.getenv("DB_PASSWORD", "")

_pool = None

async def get_pool():
    global _pool
    if _pool is None:
        _pool = await asyncpg.create_pool(
            host=DB_HOST,
            port=DB_PORT,
            database=DB_NAME,
            user=DB_USER,
            password=DB_PASSWORD,
            min_size=1,
            max_size=10
        )
    return _pool

async def init_db():
    pool = await get_pool()
    async with pool.acquire() as conn:
        await conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                email TEXT UNIQUE NOT NULL,
                hashed_password TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        await conn.execute("""
            CREATE TABLE IF NOT EXISTS licenses (
                id SERIAL PRIMARY KEY,
                license_key TEXT UNIQUE NOT NULL,
                user_email TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                expires_at TIMESTAMP NOT NULL,
                is_active BOOLEAN DEFAULT TRUE,
                activated_at TIMESTAMP
            )
        """)
        await conn.execute("""
            CREATE TABLE IF NOT EXISTS analysis_cache (
                reg_number TEXT PRIMARY KEY,
                result_json TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        await conn.execute("""
            CREATE TABLE IF NOT EXISTS user_usage (
                id SERIAL PRIMARY KEY,
                user_email TEXT NOT NULL,
                month_year TEXT NOT NULL,
                analyses_used INTEGER DEFAULT 0,
                analyses_limit INTEGER DEFAULT 50,
                UNIQUE(user_email, month_year)
            )
        """)
    print("✅ База данных PostgreSQL инициализирована")

# ================= ПОЛЬЗОВАТЕЛИ =================
async def create_user(email: str, password: str):
    password = password[:72]
    hashed = sha256_crypt.hash(password)
    pool = await get_pool()
    async with pool.acquire() as conn:
        try:
            await conn.execute(
                "INSERT INTO users (email, hashed_password) VALUES ($1, $2)",
                email, hashed
            )
            return True
        except asyncpg.UniqueViolationError:
            return False

async def get_user(email: str):
    pool = await get_pool()
    async with pool.acquire() as conn:
        row = await conn.fetchrow("SELECT * FROM users WHERE email = $1", email)
        return dict(row) if row else None

def verify_password(plain_password: str, hashed_password: str):
    plain_password = plain_password[:72]
    return sha256_crypt.verify(plain_password, hashed_password)

# ================= ЛИЦЕНЗИИ =================
def generate_license_key():
    alphabet = string.ascii_uppercase + string.digits
    return '-'.join(''.join(secrets.choice(alphabet) for _ in range(4)) for _ in range(4))

async def create_license(email=None, days_valid=30):
    pool = await get_pool()
    license_key = generate_license_key()
    expires_at = datetime.now() + timedelta(days=days_valid)
    async with pool.acquire() as conn:
        try:
            await conn.execute(
                "INSERT INTO licenses (license_key, user_email, expires_at, is_active) VALUES ($1, $2, $3, $4)",
                license_key, email, expires_at, True
            )
            return license_key
        except asyncpg.UniqueViolationError:
            return None

async def verify_and_activate_license(key):
    pool = await get_pool()
    async with pool.acquire() as conn:
        row = await conn.fetchrow("SELECT * FROM licenses WHERE license_key = $1", key)
        if not row:
            return {"valid": False, "reason": "not_found"}
        lic = dict(row)
        if not lic["is_active"]:
            return {"valid": False, "reason": "already_used"}
        if datetime.now() > lic["expires_at"]:
            return {"valid": False, "reason": "expired"}
        await conn.execute(
            "UPDATE licenses SET is_active = FALSE, activated_at = $1 WHERE license_key = $2",
            datetime.now(), key
        )
        return {"valid": True, "expires_at": lic["expires_at"]}

async def verify_license(key):
    pool = await get_pool()
    async with pool.acquire() as conn:
        row = await conn.fetchrow("SELECT * FROM licenses WHERE license_key = $1", key)
        if not row:
            return None
        lic = dict(row)
        if datetime.now() > lic["expires_at"]:
            return {"valid": False, "reason": "expired"}
        return {"valid": True, "expires_at": lic["expires_at"], "email": lic["user_email"]}

async def deactivate_license(key):
    pool = await get_pool()
    async with pool.acquire() as conn:
        await conn.execute("UPDATE licenses SET is_active = FALSE WHERE license_key = $1", key)

# ================= КЭШ АНАЛИЗА =================
async def get_cached_analysis(reg_number):
    pool = await get_pool()
    async with pool.acquire() as conn:
        row = await conn.fetchrow(
            "SELECT result_json FROM analysis_cache WHERE reg_number = $1", reg_number
        )
        if row:
            return json.loads(row["result_json"])
        return None

async def save_analysis_cache(reg_number, result):
    pool = await get_pool()
    async with pool.acquire() as conn:
        await conn.execute(
            "INSERT INTO analysis_cache (reg_number, result_json) VALUES ($1, $2) ON CONFLICT (reg_number) DO UPDATE SET result_json = EXCLUDED.result_json",
            reg_number, json.dumps(result, ensure_ascii=False)
        )

async def clear_cache():
    pool = await get_pool()
    async with pool.acquire() as conn:
        await conn.execute("DELETE FROM analysis_cache")

# ================= ТАРИФИКАЦИЯ =================
async def get_user_plan_limit(email):
    return 50

async def check_and_increment_usage(email):
    month_year = datetime.now().strftime("%Y-%m")
    pool = await get_pool()
    async with pool.acquire() as conn:
        row = await conn.fetchrow(
            "SELECT analyses_used, analyses_limit FROM user_usage WHERE user_email = $1 AND month_year = $2",
            email, month_year
        )
        if not row:
            limit = await get_user_plan_limit(email)
            await conn.execute(
                "INSERT INTO user_usage (user_email, month_year, analyses_used, analyses_limit) VALUES ($1, $2, 1, $3)",
                email, month_year, limit
            )
            return True
        used = row["analyses_used"]
        limit = row["analyses_limit"]
        if used >= limit:
            return False
        await conn.execute(
            "UPDATE user_usage SET analyses_used = analyses_used + 1 WHERE user_email = $1 AND month_year = $2",
            email, month_year
        )
        return True
