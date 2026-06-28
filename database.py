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
            max_size=10,
            timeout=30
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
                activated_at TIMESTAMP,
                total_requests INTEGER DEFAULT 0,
                total_tokens INTEGER DEFAULT 0
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
        await conn.execute("""
            CREATE TABLE IF NOT EXISTS device_trials (
                device_id TEXT PRIMARY KEY,
                start_date TIMESTAMP NOT NULL,
                trial_days INTEGER DEFAULT 2,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                total_requests INTEGER DEFAULT 0,
                total_tokens INTEGER DEFAULT 0,
                ip_address TEXT,
                user_agent TEXT
            )
        """)

        await conn.execute("ALTER TABLE licenses ADD COLUMN IF NOT EXISTS total_requests INTEGER DEFAULT 0")
        await conn.execute("ALTER TABLE licenses ADD COLUMN IF NOT EXISTS total_tokens INTEGER DEFAULT 0")
        await conn.execute("ALTER TABLE device_trials ADD COLUMN IF NOT EXISTS total_requests INTEGER DEFAULT 0")
        await conn.execute("ALTER TABLE device_trials ADD COLUMN IF NOT EXISTS total_tokens INTEGER DEFAULT 0")
        await conn.execute("ALTER TABLE device_trials ADD COLUMN IF NOT EXISTS ip_address TEXT")
        await conn.execute("ALTER TABLE device_trials ADD COLUMN IF NOT EXISTS user_agent TEXT")

    print("✅ База данных PostgreSQL инициализирована (с защитой по IP+User-Agent)")

async def create_user(email: str, password: str):
    password = password[:72]
    hashed = sha256_crypt.hash(password)
    pool = await get_pool()
    async with pool.acquire() as conn:
        try:
            await conn.execute("INSERT INTO users (email, hashed_password) VALUES ($1, $2)", email, hashed)
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
                "INSERT INTO licenses (license_key, user_email, expires_at, is_active, total_requests, total_tokens) VALUES ($1, $2, $3, $4, 0, 0)",
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
        await conn.execute("UPDATE licenses SET is_active = FALSE, activated_at = $1 WHERE license_key = $2", datetime.now(), key)
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

async def get_cached_analysis(reg_number):
    pool = await get_pool()
    async with pool.acquire() as conn:
        row = await conn.fetchrow("SELECT result_json FROM analysis_cache WHERE reg_number = $1", reg_number)
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

# ================= ПРОБНЫЙ ПЕРИОД (ЗАЩИТА ПО IP+USER-AGENT) =================
async def start_trial(device_id: str, trial_days: int = 2, ip_address: str = None, user_agent: str = None):
    pool = await get_pool()
    async with pool.acquire() as conn:
        # 1. Проверяем, есть ли уже запись для этого device_id
        existing = await conn.fetchrow(
            "SELECT start_date FROM device_trials WHERE device_id = $1",
            device_id
        )
        if existing:
            return {"status": "ok", "trial_start": existing["start_date"]}

        # 2. Если device_id нет — проверяем IP+User-Agent (защита от переустановок)
        if ip_address and user_agent:
            row = await conn.fetchrow(
                "SELECT device_id FROM device_trials WHERE ip_address = $1 AND user_agent = $2",
                ip_address, user_agent
            )
            if row:
                return {"status": "already_used", "device_id": row["device_id"]}

        # 3. Создаём новую запись
        now = datetime.now()
        await conn.execute(
            "INSERT INTO device_trials (device_id, start_date, trial_days, ip_address, user_agent, total_requests, total_tokens) VALUES ($1, $2, $3, $4, $5, 0, 0)",
            device_id, now, trial_days, ip_address, user_agent
        )
        return {"status": "ok", "trial_start": now}

async def get_trial_status(device_id: str):
    pool = await get_pool()
    async with pool.acquire() as conn:
        row = await conn.fetchrow(
            "SELECT start_date, trial_days FROM device_trials WHERE device_id = $1",
            device_id
        )
        if not row:
            return {"status": "not_started"}
        start_date = row["start_date"]
        trial_days = row["trial_days"]
        days_passed = (datetime.now() - start_date).days
        if days_passed >= trial_days:
            return {"status": "expired", "days": days_passed}
        else:
            return {
                "status": "active",
                "days_left": trial_days - days_passed,
                "trial_end": (start_date + timedelta(days=trial_days)).isoformat()
            }

async def check_trial_by_device(device_id: str) -> bool:
    status = await get_trial_status(device_id)
    return status.get("status") == "active"

async def increment_usage(license_key: str = None, device_id: str = None, tokens_used: int = 0):
    if not license_key and not device_id:
        return
    pool = await get_pool()
    async with pool.acquire() as conn:
        if license_key:
            await conn.execute(
                "UPDATE licenses SET total_requests = total_requests + 1, total_tokens = total_tokens + $1 WHERE license_key = $2",
                tokens_used, license_key
            )
        elif device_id:
            await conn.execute(
                "UPDATE device_trials SET total_requests = total_requests + 1, total_tokens = total_tokens + $1 WHERE device_id = $2",
                tokens_used, device_id
            )
