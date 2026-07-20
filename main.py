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
import pdfplumber
from fastapi import FastAPI, UploadFile, File, HTTPException, Form, Request, Response
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from urllib.parse import quote
import openpyxl
import xlrd
import uvicorn
import database
import parser
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field, validator
from typing import List, Optional
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded

# ================= НОВАЯ БИБЛИОТЕКА ЮKASSA =================
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

# ================= НАСТРОЙКИ ЮKASSA =================
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
    allow_origins=["*"],
    allow_credentials=False,
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

# ================= ФУНКЦИЯ ОТПРАВКИ В MAX =================
async def send_max_notification(
    reg_number: str,
    client_name: str,
    inn: str,
    phone: str,
    email: str,
    nmc: str,
    end_date: str,
    bid_end_date: str,
    bid_security: str,
    contract_security: str,
    guarantee_type: str,
    contact_by_email: bool
):
    if not MAX_BOT_TOKEN or not MAX_CHAT_ID:
        logger.warning("MAX Bot не настроен: пропуск уведомления")
        return

    text = (
        f"Новая заявка на банковскую гарантию!\n\n"
        f"Номер тендера: {reg_number}\n"
        f"Клиент: {client_name}\n"
        f"ИНН: {inn or 'Не указан'}\n"
        f"Телефон: {phone}\n"
        f"Email: {email}\n"
        f"Начальная цена (НМЦ): {nmc or 'Не указана'}\n"
        f"Дата окончания контракта: {end_date or 'Не указана'}\n"
        f"Дата окончания подачи заявок: {bid_end_date or 'Не указана'}\n"
        f"Обеспечение заявки: {bid_security or 'Не указано'}\n"
        f"Обеспечение контракта: {contract_security or 'Не указано'}\n"
        f"Тип гарантии: {'Обеспечение заявки (участие)' if guarantee_type == 'participation' else 'Обеспечение исполнения контракта'}\n"
        f"Связь только по email: {'Да' if contact_by_email else 'Нет (звонить)'}\n"
        f"Время заявки: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    )

    url = f"{MAX_API_URL}?chat_id={MAX_CHAT_ID}"
    headers = {
        "Authorization": MAX_BOT_TOKEN,
        "Content-Type": "application/json"
    }
    payload = {"text": text}

    try:
        async with httpx.AsyncClient(timeout=10, verify=False) as client:
            response = await client.post(url, headers=headers, json=payload)
            if response.status_code == 200:
                logger.info(f"Уведомление в MAX отправлено (тендер {reg_number})")
            else:
                logger.error(f"Ошибка MAX API: {response.status_code} - {response.text}")
    except Exception as e:
        logger.error(f"Исключение при отправке в MAX: {e}")

# ================= ФУНКЦИИ ЧТЕНИЯ ФАЙЛОВ =================
def read_docx(file_path):
    try:
        doc = Document(file_path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except Exception as e:
        return f"Ошибка чтения docx: {e}"

def read_txt(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except:
        try:
            with open(file_path, 'r', encoding='cp1251') as f:
                return f.read()
        except Exception as e:
            return f"Ошибка чтения txt: {e}"

def read_excel(file_path):
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active
        rows = sheet.iter_rows(values_only=True)
        return "\n".join(str(cell) for row in rows for cell in row if cell)
    except:
        try:
            wb = xlrd.open_workbook(file_path)
            sheet = wb.sheet_by_index(0)
            text = ""
            for row in range(sheet.nrows):
                row_text = [str(sheet.cell_value(row, col)) for col in range(sheet.ncols) if sheet.cell_value(row, col)]
                if row_text:
                    text += " ".join(row_text) + "\n"
            return text
        except Exception as e:
            return f"Ошибка чтения Excel: {e}"

def read_pdf(file_path):
    try:
        text = ""
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        text = text.strip()
        if not text:
            logger.info(f"⚠️ [PDF] {file_path} не содержит текста (возможно, сканированный)")
        return text if text else "PDF не содержит текста (возможно, сканированный документ)"
    except Exception as e:
        logger.error(f"❌ [PDF] Ошибка при открытии {file_path}: {e}")
        return f"Ошибка чтения PDF: {e}"

# ================= ОПРЕДЕЛЕНИЕ ТИПА ФАЙЛА =================
def detect_file_type(content: bytes, filename: str = "") -> str:
    ext = os.path.splitext(filename)[1].lower() if filename else ""

    if content.startswith(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'):
        if b'WorkBook' in content[:2000] or b'BOUNDSHEET' in content[:2000]:
            return 'xls'
        return 'doc'

    if ext == '.doc':
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                if any(f.startswith('word/') for f in zf.namelist()):
                    return 'docx'
        except:
            pass
        return 'doc'

    if content.startswith(b'%PDF'):
        return 'pdf'

    if content.startswith(b'PK\x03\x04') or content.startswith(b'PK\x05\x06') or content.startswith(b'PK\x07\x08'):
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                files = zf.namelist()
                if any(f.startswith('word/') for f in files):
                    return 'docx'
                if any(f.startswith('xl/') for f in files):
                    return 'xlsx'
                return 'zip'
        except:
            return 'zip'

    if content.startswith(b'Rar!\x1a\x07\x00') or content.startswith(b'Rar!\x1a\x07\x01\x00') or content.startswith(b'Rar!'):
        return 'rar'
    if content.startswith(b'7z\xbc\xaf\x27\x1c'):
        return '7z'
    if content.startswith(b'\x89PNG\r\n\x1a\n'):
        return 'png'
    if content.startswith(b'\xff\xd8\xff'):
        return 'jpg'
    if content.startswith(b'{\\rtf'):
        return 'rtf'

    if ext == '.doc':
        return 'doc'
    return 'unknown'

# ================= ЗАПРОСЫ К DEEPSEEK =================
def query_deepseek(prompt, license_key=None, device_id=None):
    messages = [{"role": "system", "content": "Ты — эксперт по анализу тендерной документации. Отвечай чётко, по делу, без воды."},
                {"role": "user", "content": prompt}]
    payload = {"model": MODEL_NAME, "messages": messages, "temperature": 0.3, "max_tokens": 2500, "stream": False}
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    try:
        response = requests.post(OLLAMA_API_URL, json=payload, headers=headers, timeout=180)
        total_tokens = 0
        if response.status_code == 200:
            result = response.json()
            usage = result.get("usage", {})
            total_tokens = usage.get("total_tokens", 0)
            answer = result["choices"][0]["message"]["content"]
        else:
            answer = f"Ошибка HTTP {response.status_code}: {response.text}"
        if license_key or device_id:
            asyncio.create_task(database.increment_usage(license_key=license_key, device_id=device_id, tokens_used=total_tokens))
        return answer
    except Exception as e:
        error_msg = f"Ошибка: {e}"
        if license_key or device_id:
            asyncio.create_task(database.increment_usage(license_key=license_key, device_id=device_id, tokens_used=0))
        return error_msg

def analyze_tender_text(text, selected_fields, license_key=None, device_id=None, max_text_len=8000):
    if not selected_fields:
        selected_fields = [
            "НАЗВАНИЕ АУКЦИОНА", "Начальная цена (НМЦ)", "ДОПОЛНИТЕЛЬНЫЕ ТРЕБОВАНИЯ К УЧАСТНИКУ",
            "ДАТА ОКОНЧАНИЯ/ПРОВЕДЕНИЯ", "Аванс", "Обеспечение заявки", "Обеспечение контракта",
            "Обеспечение гарантийных обязательств", "Контакты", "Место исполнения", "ДАТА ОКОНЧАНИЯ КОНТРАКТА"
        ]
    fields_str = "\n".join(f"{field}: " for field in selected_fields)
    truncated = text[:max_text_len]
    prompt = f"""Ты анализируешь тендерную документацию. Извлеки из текста следующие данные. Если информации нет, напиши "Информация отсутствует".
Ответ должен быть строго в таком формате (каждый пункт с новой строки):
{fields_str}
Вот текст для анализа:
{truncated}
Извлеки данные и напиши в указанном формате."""
    answer = query_deepseek(prompt, license_key=license_key, device_id=device_id)
    result = {}
    for line in answer.split('\n'):
        if ':' in line:
            k, v = line.split(':', 1)
            result[k.strip()] = v.strip()
    return result

def critical_analysis(text, license_key=None, device_id=None, custom_prompt=None):
    if custom_prompt and custom_prompt.strip():
        prompt = custom_prompt.strip()
        prompt += f"\n\nТекст тендера и прикреплённых документов:\n{text}"
    else:
        prompt = f"""Ты — эксперт по тендерной документации. Проанализируй текст тендера и выяви потенциальные риски, сложности, скрытые требования, которые могут помешать выполнению контракта. Оцени, насколько выполним данный тендер для среднестатистического поставщика.

    Текст тендера и документов:
    {text}

    Ответ должен быть в виде связного текста (5–7 предложений), где будут перечислены:
    - основные риски и сложности,
    - неочевидные требования,
    - рекомендации по успешному выполнению.
    Не используй маркированные списки, пиши сплошным текстом.
    """
    answer = query_deepseek(prompt, license_key=license_key, device_id=device_id)
    return answer

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
    return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="oferta.docx")

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
    amount_value = data.get("amount", 2500) if data else 2500
    device_id = request.headers.get("X-Device-ID", "unknown")
    email = data.get("email", "client@example.com") if data else "client@example.com"
    try:
        payment = Payment.create({
            "amount": {"value": str(amount_value), "currency": "RUB"},
            "confirmation": {
                "type": "redirect",
                "return_url": f"https://csb24-tender.ru/success?device_id={device_id}",
                "cancel_url": "https://csb24-tender.ru/"
            },
            "capture": True,
            "description": "Лицензия для Тендерного парсера (1 месяц)",
            "metadata": {"device_id": device_id},
            "receipt": {
                "customer": {"email": email},
                "items": [{
                    "description": "Лицензия на использование Тендерного парсера (1 месяц)",
                    "quantity": "1.00",
                    "amount": {"value": str(amount_value), "currency": "RUB"},
                    "vat_code": 4,
                    "payment_mode": "full_payment",
                    "payment_subject": "service"
                }]
            }
        })
        await database.save_payment(payment.id, device_id)
        logger.info(f"Платёж создан: id={payment.id}, сумма={payment.amount.value} RUB")
        return {"id": payment.id, "confirmation_url": payment.confirmation.confirmation_url, "amount": payment.amount.value}
    except Exception as e:
        logger.error(f"Ошибка создания платежа: {e}")
        raise HTTPException(500, f"Не удалось создать платёж: {str(e)}")

@app.post("/yookassa-webhook")
@limiter.limit("10/minute")
async def yookassa_webhook(request: Request):
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
        license_key = await database.create_license(email=device_id, days_valid=30)
        if license_key:
            await database.update_payment_license(payment_id, license_key)
            logger.info(f"Лицензия {license_key} создана для device_id {device_id}")
        else:
            logger.warning(f"Не удалось создать лицензию для {device_id}")
        return {"status": "ok"}
    return {"status": "ok"}

@app.get("/success", response_class=HTMLResponse)
async def success_page(request: Request):
    device_id = request.query_params.get("device_id")
    if not device_id:
        return HTMLResponse("<html><body><h1>Ошибка</h1><p>Не указан идентификатор устройства.</p></body></html>")
    if device_id == "unknown":
        return HTMLResponse("<html><body><h1>Ошибка</h1><p>Идентификатор устройства не определён.</p></body></html>")
    license_key = await database.get_license_by_device(device_id)
    if not license_key:
        return HTMLResponse(f"""
        <!DOCTYPE html>
        <html>
        <head><meta charset="UTF-8"><title>Оплата успешна</title>
        <style>body{{font-family:Arial;text-align:center;padding:40px}}.key{{background:#f1f5f9;padding:15px;border-radius:8px;font-size:20px;font-weight:bold;letter-spacing:2px}}</style>
        </head>
        <body>
            <h1 style="color:#10b981;">✅ Оплата прошла успешно!</h1>
            <p>Ваш лицензионный ключ:</p>
            <div class="key" id="licenseKey">Загрузка...</div>
            <p><a href="/">На главную</a></p>
            <script>
                (async function(){{
                    const deviceId="{device_id}";
                    let attempts=0;
                    while(attempts<15){{
                        try{{
                            const resp=await fetch(`/api/get-license-by-device?device_id=${{deviceId}}`);
                            const data=await resp.json();
                            if(data.license_key){{
                                document.getElementById('licenseKey').textContent=data.license_key;
                                return;
                            }}
                        }}catch(e){{}}
                        attempts++;
                        await new Promise(r=>setTimeout(r,2000));
                    }}
                    document.getElementById('licenseKey').textContent='Не удалось получить ключ. Обратитесь в поддержку.';
                }})();
            </script>
        </body>
        </html>
        """)
    return HTMLResponse(f"""
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8"><title>Оплата успешна</title>
    <style>body{{font-family:Arial;text-align:center;padding:40px}}.key{{background:#f1f5f9;padding:15px;border-radius:8px;font-size:20px;font-weight:bold;letter-spacing:2px}}</style>
    </head>
    <body>
        <h1 style="color:#10b981;">✅ Оплата прошла успешно!</h1>
        <p>Ваш лицензионный ключ:</p>
        <div class="key">{license_key}</div>
        <p><a href="/">На главную</a></p>
    </body>
    </html>
    """)

@app.get("/api/get-license-by-device")
@limiter.limit("10/minute")
async def get_license_by_device(request: Request, device_id: str):
    if not device_id:
        raise HTTPException(400, "Не указан device_id")
    license_key = await database.get_license_by_device(device_id)
    if not license_key:
        raise HTTPException(404, "Лицензия ещё не создана")
    return {"license_key": license_key}

# ================= ЛИЦЕНЗИИ =================
@app.post("/api/create-order")
@limiter.limit("10/minute")
async def create_order(request: Request):
    admin_token = request.headers.get("X-Admin-Token")
    if admin_token != ADMIN_TOKEN:
        raise HTTPException(403, "Неверный admin token")
    license_key = await database.create_license(days_valid=30)
    if not license_key:
        raise HTTPException(500, "Не удалось создать лицензию")
    result = await database.verify_license(license_key)
    expires_at = result.get("expires_at") if result and result.get("valid") else None
    return {"status": "success", "license_key": license_key, "expires_at": expires_at}

app.mount("/static", StaticFiles(directory="static"), name="static")

# ================= ЭНДПОЙНТ АНАЛИЗА ТЕКСТОВ (МАССОВЫЙ) =================
@app.post("/analyze_texts")
@limiter.limit("10/minute")
async def analyze_texts(request: Request, data: AnalyzeRequest):
    await check_access(request)
    tenders_data = data.tenders
    if not tenders_data:
        raise HTTPException(400, "Нет данных для анализа")
    selected_fields = data.fields
    if not selected_fields:
        selected_fields = [
            "НАЗВАНИЕ АУКЦИОНА", "Начальная цена (НМЦ)", "ДОПОЛНИТЕЛЬНЫЕ ТРЕБОВАНИЯ К УЧАСТНИКУ",
            "ДАТА ОКОНЧАНИЯ/ПРОВЕДЕНИЯ", "Аванс", "Обеспечение заявки", "Обеспечение контракта",
            "Обеспечение гарантийных обязательств", "Контакты", "Место исполнения", "ДАТА ОКОНЧАНИЯ КОНТРАКТА"
        ]
    license_key = request.headers.get("X-License-Key")
    device_id = request.headers.get("X-Device-ID")

    async def analyze_one(tender):
        reg_number = tender.regNumber
        cached = await database.get_cached_analysis(reg_number)
        if cached:
            await database.increment_cache_usage(reg_number)
            logger.info(f"Кэш для {reg_number} использован")
            return {"url": tender.url, "reg_number": reg_number, "analysis": cached}
        tender_text = tender.text
        if not tender_text or len(tender_text) < 100:
            return {"url": tender.url, "reg_number": reg_number, "error": "Недостаточно текста для анализа"}
        start = time.time()
        analysis_result = analyze_tender_text(tender_text, selected_fields, license_key, device_id)
        await database.save_analysis_cache(reg_number, analysis_result)
        logger.info(f"DeepSeek обработал {reg_number} за {time.time()-start:.2f} сек")
        return {"url": tender.url, "reg_number": reg_number, "analysis": analysis_result}

    tasks = [analyze_one(t) for t in tenders_data]
    results = await asyncio.gather(*tasks)

    doc = Document()
    doc.add_heading('РЕЗУЛЬТАТЫ АНАЛИЗА ТЕНДЕРОВ', 0)
    doc.add_paragraph(f'Дата: {datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')
    doc.add_paragraph('=' * 50)
    for item in results:
        doc.add_heading(f'Тендер: {item.get("reg_number", "Неизвестно")}', level=1)
        if "error" in item:
            doc.add_paragraph(f'Ошибка: {item["error"]}')
        else:
            analysis = item.get("analysis", {})
            if isinstance(analysis, dict):
                for k, v in analysis.items():
                    doc.add_paragraph(f'{k}: {v}')
            else:
                doc.add_paragraph(str(analysis))
        doc.add_page_break()

    word_buffer = io.BytesIO()
    doc.save(word_buffer)
    word_buffer.seek(0)
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zf:
        zf.writestr(f'результаты_анализа_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx', word_buffer.getvalue())
    zip_buffer.seek(0)
    filename = f'результаты_анализа_{datetime.now().strftime("%Y%m%d_%H%M%S")}.zip'
    encoded = quote(filename)
    return Response(zip_buffer.getvalue(), media_type="application/zip", headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded}"})

# ================= ЭНДПОЙНТ ДЛЯ АНАЛИЗА ТЕКУЩЕГО ТЕНДЕРА С ДОКУМЕНТАМИ =================
@app.post("/analyze_tender_with_files")
@limiter.limit("10/minute")
async def analyze_tender_with_files(
    request: Request,
    regNumber: str = Form(...),
    printFormText: str = Form(...),
    fields: str = Form(...),
    files: list[UploadFile] = File(...),
    custom_critical_prompt: str = Form("")
):
    start_total = time.time()
    logger.info(f"🚀 [START] Тендер {regNumber}, файлов: {len(files)}")

    try:
        await check_access(request)
    except Exception as e:
        logger.error(f"❌ [ACCESS] Ошибка доступа для {regNumber}: {e}")
        raise

    selected_fields = json.loads(fields)
    if not selected_fields:
        selected_fields = [
            "НАЗВАНИЕ АУКЦИОНА", "Начальная цена (НМЦ)", "ДОПОЛНИТЕЛЬНЫЕ ТРЕБОВАНИЯ К УЧАСТНИКУ",
            "ДАТА ОКОНЧАНИЯ/ПРОВЕДЕНИЯ", "Аванс", "Обеспечение заявки", "Обеспечение контракта",
            "Обеспечение гарантийных обязательств", "Контакты", "Место исполнения", "ДАТА ОКОНЧАНИЯ КОНТРАКТА"
        ]

    device_id = request.headers.get("X-Device-ID", "unknown")
    license_key = request.headers.get("X-License-Key", None)

    prompt_info = "стандартного"
    if custom_critical_prompt and custom_critical_prompt.strip():
        prompt_info = "вашего личного"

    combined_text = printFormText
    original_files = []

    temp_dir = Path(f"/tmp/tender_{regNumber}_{datetime.now().strftime('%Y%m%d%H%M%S')}")
    temp_dir.mkdir(parents=True, exist_ok=True)

    try:
        start_files = time.time()
        for idx, file in enumerate(files):
            logger.info(f"📄 [FILE] Обработка файла {idx+1}/{len(files)}: {file.filename}")
            content = await file.read()

            file_type = detect_file_type(content, file.filename)
            ext_map = {
                'pdf': '.pdf', 'xlsx': '.xlsx', 'xls': '.xls', 'docx': '.docx',
                'rar': '.rar', 'zip': '.zip', '7z': '.7z', 'png': '.png',
                'jpg': '.jpg', 'rtf': '.rtf', 'doc': '.doc'
            }
            new_ext = ext_map.get(file_type, '')
            base, old_ext = os.path.splitext(file.filename)
            corrected_filename = base + new_ext if new_ext else file.filename

            file_path = temp_dir / corrected_filename
            with open(file_path, "wb") as f:
                f.write(content)

            ext = file_path.suffix.lower()
            text = ""
            if ext == ".docx":
                try:
                    text = read_docx(str(file_path))
                except Exception as e:
                    logger.error(f"❌ [DOCX] Ошибка чтения {file.filename}: {e}")
                    text = f"[Ошибка чтения DOCX: {e}]"
                else:
                    logger.info(f"📄 [DOCX] {file.filename} -> {len(text)} символов извлечено")
            elif ext == ".txt":
                text = read_txt(str(file_path))
                if text:
                    logger.info(f"📄 [TXT] {file.filename} -> {len(text)} символов извлечено")
            elif ext in (".xlsx", ".xls"):
                try:
                    text = read_excel(str(file_path))
                except Exception as e:
                    logger.error(f"❌ [EXCEL] Ошибка чтения {file.filename}: {e}")
                    text = f"[Ошибка чтения Excel: {e}]"
                else:
                    logger.info(f"📄 [EXCEL] {file.filename} -> {len(text)} символов извлечено
