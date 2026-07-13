import os
import shutil
import zipfile
import io
import json
import time
import asyncio
import hashlib
import logging
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from docx import Document
import requests
import httpx
from fastapi import FastAPI, UploadFile, File, HTTPException, Form, Request, Response, Depends
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.middleware.trustedhost import TrustedHostMiddleware
from urllib.parse import quote
import openpyxl
import xlrd
import uvicorn
import database
import parser
from bs4 import BeautifulSoup
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field, validator
from typing import List, Optional
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded

# ========== НОВАЯ БИБЛИОТЕКА ЮKASSA ==========
from yookassa import Configuration, Payment

# ================= ЗАГРУЗКА .ENV =================
load_dotenv()

# ================= ЛОГИРОВАНИЕ =================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ================= НАСТРОЙКИ =================
API_KEY = os.getenv("DEEPSEEK_API_KEY")
if not API_KEY:
    raise ValueError("DEEPSEEK_API_KEY не найден в .env")

OLLAMA_API_URL = "https://api.deepseek.com/v1/chat/completions"
MODEL_NAME = "deepseek-chat"
MAX_TENDERS = 15
ADMIN_TOKEN = os.getenv("ADMIN_TOKEN")
if not ADMIN_TOKEN:
    raise ValueError("ADMIN_TOKEN не найден в .env")

# ================= НАСТРОЙКИ ЮKASSA (вместо Робокассы) =================
YOOKASSA_SHOP_ID = os.getenv("YOOKASSA_SHOP_ID")
YOOKASSA_SECRET_KEY = os.getenv("YOOKASSA_SECRET_KEY")
if not all([YOOKASSA_SHOP_ID, YOOKASSA_SECRET_KEY]):
    raise ValueError("ЮKassa настройки не найдены в .env")

Configuration.account_id = YOOKASSA_SHOP_ID
Configuration.secret_key = YOOKASSA_SECRET_KEY

# ================= НАСТРОЙКИ MAX BOT =================
MAX_BOT_TOKEN = os.getenv("MAX_BOT_TOKEN")
MAX_CHAT_ID = os.getenv("MAX_CHAT_ID")
MAX_API_URL = "https://platform-api2.max.ru/messages"

if not all([MAX_BOT_TOKEN, MAX_CHAT_ID]):
    raise ValueError("MAX BOT настройки не найдены в .env")

# ================= FASTAPI + RATE LIMITING =================
limiter = Limiter(key_func=get_remote_address)
app = FastAPI()
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

# ================= CORS =================
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "chrome-extension://*",
        "https://csb24-tender.ru",
        "http://csb24-tender.ru"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ================= MIDDLEWARE ЛОГИРОВАНИЯ =================
@app.middleware("http")
async def log_requests(request: Request, call_next):
    start = time.time()
    response = await call_next(request)
    duration = time.time() - start
    license_header = request.headers.get("X-License-Key", "none")
    logger.info(
        f"{request.method} {request.url.path} - {response.status_code} - "
        f"{duration:.2f}s - {request.client.host} - "
        f"License: {license_header[:8]}..."
    )
    return response

# ================= PYDANTIC МОДЕЛИ =================
class TenderData(BaseModel):
    url: str = Field(..., max_length=500)
    regNumber: str = Field(..., max_length=50)
    text: str = Field(..., max_length=50000)

    @validator("url")
    def validate_url(cls, v):
        if not v.startswith(("http://", "https://")):
            raise ValueError("URL должен начинаться с http:// или https://")
        return v

class AnalyzeRequest(BaseModel):
    tenders: List[TenderData] = Field(..., max_items=50)
    fields: List[str] = Field(default_factory=list, max_items=20)

class TrialRequest(BaseModel):
    device_id: str = Field(..., min_length=10, max_length=100)

class LicenseActivateRequest(BaseModel):
    key: str = Field(..., min_length=10, max_length=100)

# ================= ИНИЦИАЛИЗАЦИЯ БД =================
@app.on_event("startup")
async def startup():
    try:
        await database.init_db()
        logger.info("База данных PostgreSQL инициализирована")
    except Exception as e:
        logger.error(f"Ошибка инициализации БД: {e}")
        raise

# ================= ПРОВЕРКА ДОСТУПА =================
async def check_access(request: Request):
    license_key = request.headers.get("X-License-Key")
    if license_key:
        result = await database.verify_license(license_key)
        if result and result.get("valid"):
            return {"type": "license", "id": license_key}
   
    device_id = request.headers.get("X-Device-ID")
    if device_id:
        is_active = await database.check_trial_by_device(device_id)
        if is_active:
            return {"type": "trial", "id": device_id}
   
    logger.warning(f"Неавторизованный доступ с {request.client.host}")
    raise HTTPException(401, detail="Требуется действующая лицензия или активный пробный период")

# ================= ФУНКЦИЯ ОТПРАВКИ УВЕДОМЛЕНИЯ В MAX =================
async def send_max_notification(...):
    # (без изменений, оставлен как есть)
    pass

# ================= ЭНДПОЙНТЫ HEALTH CHECK =================
@app.get("/health")
async def health():
    return {"status": "ok"}

# ================= СТРАНИЦЫ САЙТА =================
@app.get("/", response_class=HTMLResponse)
async def main_page():
    with open("templates/main.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.get("/contacts", response_class=HTMLResponse)
async def contacts_page():
    with open("templates/contacts.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.get("/offer", response_class=HTMLResponse)
async def offer_page():
    with open("templates/offer.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.get("/oferta")
async def download_oferta():
    file_path = os.path.join("static", "oferta.docx")
    if not os.path.exists(file_path):
        raise HTTPException(404, "Файл оферты не найден")
    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="oferta.docx"
    )

@app.get("/privacy-policy", response_class=HTMLResponse)
async def privacy_policy_page():
    with open("templates/privacy-policy.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.get("/download-file/{filename}")
async def download_file(filename: str):
    allowed_files = ["tender-parser-extension.crx", "tender-parser-extension.zip"]
    if filename not in allowed_files:
        raise HTTPException(404, "Файл не найден")
    file_path = os.path.join(os.path.dirname(__file__), filename)
    if not os.path.exists(file_path):
        raise HTTPException(404, "Файл не найден на сервере")
    media_type = "application/x-chrome-extension" if filename.endswith(".crx") else "application/zip"
    return FileResponse(file_path, media_type=media_type, filename=filename)

@app.get("/download")
async def download_redirect():
    return RedirectResponse(url="/", status_code=302)

@app.get("/updates.xml")
async def updates_xml():
    file_path = os.path.join(os.path.dirname(__file__), "updates.xml")
    if not os.path.exists(file_path):
        raise HTTPException(404, "Файл updates.xml не найден")
    return FileResponse(file_path, media_type="application/xml")

# ================= ЭНДПОЙНТЫ ЮKASSA =================
@app.post("/api/create-payment")
@limiter.limit("5/minute")
async def create_payment(request: Request, data: dict = None):
    """
    Создаёт платёж в ЮKassa.
    Ожидает сумму (опционально, по умолчанию 2500) и передаёт device_id в метаданные.
    """
    amount_value = data.get("amount", 2500) if data else 2500
    device_id = request.headers.get("X-Device-ID", "unknown")

    try:
        payment = Payment.create({
            "amount": {
                "value": str(amount_value),
                "currency": "RUB"
            },
            "confirmation": {
                "type": "redirect",
                "return_url": f"https://csb24-tender.ru/success?payment_id={{id}}"
            },
            "capture": True,
            "description": "Лицензия для Тендерного парсера",
            "metadata": {
                "device_id": device_id
            }
        })

        # Сохраняем связь payment_id → device_id
        await database.save_payment(payment.id, device_id)

        logger.info(f"Платёж создан: id={payment.id}, сумма={payment.amount.value} RUB")
        return {
            "id": payment.id,
            "confirmation_url": payment.confirmation.confirmation_url,
            "amount": payment.amount.value
        }
    except Exception as e:
        logger.error(f"Ошибка создания платежа: {e}")
        raise HTTPException(500, "Не удалось создать платёж")

@app.post("/yookassa-webhook")
@limiter.limit("10/minute")
async def yookassa_webhook(request: Request):
    """Обрабатывает уведомления от ЮKassa."""
    body = await request.json()
    event = body.get("event")
    if not event:
        raise HTTPException(400, "Неверный запрос")

    logger.info(f"Получен вебхук: {event}")

    if event == "payment.succeeded":
        payment_obj = body["object"]
        payment_id = payment_obj["id"]
        metadata = payment_obj.get("metadata", {})
        device_id = metadata.get("device_id", "unknown")

        # Создаём лицензию для device_id (в поле user_email временно храним device_id)
        license_key = await database.create_license(email=device_id, days_valid=30)
        if license_key:
            # Сохраняем ключ в таблицу payments
            await database.update_payment_license(payment_id, license_key)
            logger.info(f"Лицензия {license_key} создана для device_id {device_id}")
        else:
            logger.warning(f"Не удалось создать лицензию для {device_id}")

        return {"status": "ok"}

    # Можно обрабатывать другие события (payment.canceled и т.д.)
    return {"status": "ok"}

@app.get("/success", response_class=HTMLResponse)
async def success_page(request: Request):
    """Страница успеха после оплаты. Отображает лицензионный ключ."""
    payment_id = request.query_params.get("payment_id")
    if not payment_id:
        return HTMLResponse("""
        <html>
        <head><meta charset="UTF-8"><title>Ошибка</title></head>
        <body style="text-align:center;padding:40px;font-family:Arial;">
            <h1 style="color:#ef4444;">Ошибка</h1>
            <p>Не указан идентификатор платежа.</p>
        </body>
        </html>
        """)

    # Пытаемся получить ключ из БД
    license_key = await database.get_license_by_payment(payment_id)

    if not license_key:
        # Если ключа ещё нет (платёж не обработан), показываем сообщение о загрузке
        return HTMLResponse(f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Оплата успешна</title>
            <style>
                body {{ font-family: Arial; text-align: center; padding: 40px; }}
                .key {{ background: #f1f5f9; padding: 15px; border-radius: 8px; font-size: 20px; font-weight: bold; letter-spacing: 2px; }}
            </style>
        </head>
        <body>
            <h1 style="color: #10b981;">✅ Оплата прошла успешно!</h1>
            <p>Ваш лицензионный ключ:</p>
            <div class="key" id="licenseKey">Загрузка...</div>
            <p style="margin-top: 20px;">Скопируйте его и вставьте в расширение.</p>
            <p><a href="/">На главную</a></p>

            <script>
                (async function() {{
                    const paymentId = "{payment_id}";
                    try {{
                        const resp = await fetch(`/api/get-license?payment_id=${{paymentId}}`);
                        const data = await resp.json();
                        if (data.license_key) {{
                            document.getElementById('licenseKey').textContent = data.license_key;
                        }} else {{
                            document.getElementById('licenseKey').textContent = 'Ожидание обработки платежа... Обновите страницу через несколько минут.';
                        }}
                    }} catch (e) {{
                        document.getElementById('licenseKey').textContent = 'Ошибка получения ключа. Попробуйте обновить страницу.';
                    }}
                }})();
            </script>
        </body>
        </html>
        """)

    # Если ключ уже есть, показываем его сразу
    return HTMLResponse(f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Оплата успешна</title>
        <style>
            body {{ font-family: Arial; text-align: center; padding: 40px; }}
            .key {{ background: #f1f5f9; padding: 15px; border-radius: 8px; font-size: 20px; font-weight: bold; letter-spacing: 2px; }}
        </style>
    </head>
    <body>
        <h1 style="color: #10b981;">✅ Оплата прошла успешно!</h1>
        <p>Ваш лицензионный ключ:</p>
        <div class="key">{license_key}</div>
        <p style="margin-top: 20px;">Скопируйте его и вставьте в расширение.</p>
        <p><a href="/">На главную</a></p>
    </body>
    </html>
    """)

@app.get("/api/get-license")
@limiter.limit("10/minute")
async def get_license(payment_id: str):
    """Возвращает лицензионный ключ по ID платежа."""
    if not payment_id:
        raise HTTPException(400, "Не указан payment_id")
    license_key = await database.get_license_by_payment(payment_id)
    if not license_key:
        raise HTTPException(404, "Лицензия ещё не создана")
    return {"license_key": license_key}

# ================= УДАЛЁННЫЕ ЭНДПОЙНТЫ РОБОКАССЫ =================
# (все эндпоинты /robokassa/... удалены, их больше нет)

# ================= ОСТАЛЬНЫЕ ЭНДПОЙНТЫ =================
# (без изменений: /api/create-order, /analyze_texts, /package_files, /guarantee, /api/guarantee/request, /api/trial/start, /api/trial/status, /search_tenders, /suggest_keywords, /ask_ai, /api/verify-license, /api/activate-license)

# ... (здесь всё остальное, что не связано с оплатой, остаётся как было, я не буду дублировать весь код, чтобы не превышать лимит символов, но вы можете просто заменить верхнюю часть и добавить новые эндпоинты)
