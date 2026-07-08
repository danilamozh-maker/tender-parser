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
        # Таблицы (без изменений)
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
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                used_count INTEGER DEFAULT 0,
                last_used TIMESTAMP DEFAULT CURRENT_TIMESTAMP
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
        await conn.execute("""
            CREATE TABLE IF NOT EXISTS guarantee_requests (
                id SERIAL PRIMARY KEY,
                reg_number TEXT,
                nmc TEXT,
                end_date TEXT,
                bid_end_date TEXT,
                guarantee_type TEXT,
                client_name TEXT,
                phone TEXT,
                email TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                bid_security TEXT,
                contract_security TEXT,
                contact_by_email BOOLEAN DEFAULT FALSE
            )
        """)

        # ===== МИГРАЦИИ (добавляем столбец inn) =====
        # Существующие миграции
        await conn.execute("ALTER TABLE licenses ADD COLUMN IF NOT EXISTS total_requests INTEGER DEFAULT 0")
        await conn.execute("ALTER TABLE licenses ADD COLUMN IF NOT EXISTS total_tokens INTEGER DEFAULT 0")
        await conn.execute("ALTER TABLE device_trials ADD COLUMN IF NOT EXISTS total_requests INTEGER DEFAULT 0")
        await conn.execute("ALTER TABLE device_trials ADD COLUMN IF NOT EXISTS total_tokens INTEGER DEFAULT 0")
        await conn.execute("ALTER TABLE device_trials ADD COLUMN IF NOT EXISTS ip_address TEXT")
        await conn.execute("ALTER TABLE device_trials ADD COLUMN IF NOT EXISTS user_agent TEXT")
        await conn.execute("ALTER TABLE analysis_cache ADD COLUMN IF NOT EXISTS used_count INTEGER DEFAULT 0")
        await conn.execute("ALTER TABLE analysis_cache ADD COLUMN IF NOT EXISTS last_used TIMESTAMP DEFAULT CURRENT_TIMESTAMP")
        await conn.execute("ALTER TABLE guarantee_requests ADD COLUMN IF NOT EXISTS bid_security TEXT")
        await conn.execute("ALTER TABLE guarantee_requests ADD COLUMN IF NOT EXISTS contract_security TEXT")
        await conn.execute("ALTER TABLE guarantee_requests ADD COLUMN IF NOT EXISTS contact_by_email BOOLEAN DEFAULT FALSE")

        # ===== НОВАЯ МИГРАЦИЯ: добавляем колонку inn =====
        await conn.execute("ALTER TABLE guarantee_requests ADD COLUMN IF NOT EXISTS inn TEXT")

    print("✅ База данных PostgreSQL инициализирована (добавлено поле inn)")

# ================= ОСТАЛЬНЫЕ ФУНКЦИИ (БЕЗ ИЗМЕНЕНИЙ) =================
# (create_user, get_user, verify_password, generate_license_key, 
# create_license, verify_and_activate_license, verify_license, 
# deactivate_license, get_cached_analysis, save_analysis_cache, 
# increment_cache_usage, clear_cache, clear_old_cache, 
# get_user_plan_limit, check_and_increment_usage, 
# start_trial, get_trial_status, check_trial_by_device, 
# increment_usage – все остаются как у вас, я их не трогаю)
