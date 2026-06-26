import os
import shutil
import zipfile
import io
import json
import time
import asyncio
from pathlib import Path
from datetime import datetime
from docx import Document
import requests
from fastapi import FastAPI, UploadFile, File, HTTPException, Form, Request, Response
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from urllib.parse import quote
import openpyxl
import xlrd
import uvicorn
import database
import parser

app = FastAPI()

# ================= НАСТРОЙКИ =================
OLLAMA_API_URL = "https://api.deepseek.com/v1/chat/completions"
API_KEY = "sk-a1866f43ed134eb48d617185cda7cd56" # замени на свой
MODEL_NAME = "deepseek-chat"
MAX_TENDERS = 15
ADMIN_TOKEN = os.getenv("ADMIN_TOKEN", "ваш_секретный_токен")
# ============================================

# ================= ИНИЦИАЛИЗАЦИЯ БД (с защитой от падений) =================
@app.on_event("startup")
async def startup():
    try:
        await database.init_db()
        print("✅ База данных PostgreSQL инициализирована")
    except Exception as e:
        print(f"❌ Ошибка инициализации БД: {e}")

# ================= ПРОВЕРКА ДОСТУПА (ЛИЦЕНЗИЯ ИЛИ ПРОБНЫЙ ПЕРИОД) =================
async def check_access(request: Request):
    """
    Проверяет наличие либо валидной лицензии (X-License-Key),
    либо активного пробного периода (X-Device-ID).
    """
    # 1. Проверяем лицензию
    license_key = request.headers.get("X-License-Key")
    if license_key:
        result = await database.verify_license(license_key)
        if result and result.get("valid"):
            return True

    # 2. Проверяем пробный период
    device_id = request.headers.get("X-Device-ID")
    if device_id:
        is_active = await database.check_trial_by_device(device_id)
        if is_active:
            return True

    # Если ни то, ни другое — доступа нет
    raise HTTPException(401, detail="Unauthorized: valid license or active trial required")

# ================= ЭНДПОЙНТЫ HEALTH CHECK =================
@app.get("/")
async def root():
    return {"status": "ok", "message": "Tender Parser API is running"}

@app.get("/health")
async def health():
    return {"status": "ok"}

# ================= СТРАНИЦА СКАЧИВАНИЯ И ПОКУПКИ =================
@app.get("/download", response_class=HTMLResponse)
async def download_page():
    with open("templates/download.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

# ================= ЭНДПОЙНТ ДЛЯ СКАЧИВАНИЯ ФАЙЛОВ =================
@app.get("/download-file/{filename}")
async def download_file(filename: str):
    # Безопасность: разрешаем скачивать только наши файлы
    allowed_files = ["tender-parser-extension.crx", "tender-parser-extension.zip"]
    if filename not in allowed_files:
        raise HTTPException(404, "Файл не найден")
    file_path = os.path.join(os.path.dirname(__file__), filename)
    if not os.path.exists(file_path):
        raise HTTPException(404, "Файл не найден на сервере")
    # Определяем MIME-тип
    if filename.endswith(".crx"):
        media_type = "application/x-chrome-extension"
    else:
        media_type = "application/zip"
    return FileResponse(file_path, media_type=media_type, filename=filename)

# ================= ЭНДПОЙНТ ДЛЯ ГЕНЕРАЦИИ ЛИЦЕНЗИИ (через админ-токен) =================
@app.post("/api/create-order")
async def create_order(request: Request):
    admin_token = request.headers.get("X-Admin-Token")
    if admin_token != ADMIN_TOKEN:
        raise HTTPException(403, "Invalid admin token")
    license_key = await database.create_license(days_valid=30)
    if not license_key:
        raise HTTPException(500, "Не удалось создать лицензию")
    result = await database.verify_license(license_key)
    expires_at = result.get("expires_at") if result and result.get("valid") else None
    return {"status": "success", "license_key": license_key, "expires_at": expires_at}

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

def query_deepseek(prompt):
    messages = [
        {"role": "system", "content": "Ты — эксперт по анализу тендерной документации. Отвечай чётко, по делу, без воды."},
        {"role": "user", "content": prompt}
    ]
    payload = {
        "model": MODEL_NAME,
        "messages": messages,
        "temperature": 0.3,
        "max_tokens": 2500,
        "stream": False
    }
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    try:
        response = requests.post(OLLAMA_API_URL, json=payload, headers=headers, timeout=60)
        if response.status_code == 200:
            result = response.json()
            return result["choices"][0]["message"]["content"]
        else:
            return f"Ошибка HTTP {response.status_code}: {response.text}"
    except Exception as e:
        return f"❌ Ошибка: {e}"

def analyze_file(file_path, selected_fields):
    ext = Path(file_path).suffix.lower()
    if ext == ".docx":
        content = read_docx(file_path)
    elif ext == ".txt":
        content = read_txt(file_path)
    elif ext in (".xlsx", ".xls"):
        content = read_excel(file_path)
    else:
        return f"⚠️ Неподдерживаемый формат: {ext}"
    if not content or "Ошибка чтения" in content:
        return f"⚠️ Не удалось прочитать файл"
    content = content[:30000]
    if not content.strip():
        return "⚠️ Файл пустой"
    if not selected_fields:
        selected_fields = [
            "НАЗВАНИЕ АУКЦИОНА", "Начальная цена (НМЦ)", "ДОПОЛНИТЕЛЬНЫЕ ТРЕБОВАНИЯ К УЧАСТНИКУ",
            "ДАТА ОКОНЧАНИЯ/ПРОВЕДЕНИЯ", "Аванс", "Обеспечение заявки", "Обеспечение контракта",
            "Обеспечение гарантийных обязательств", "Контакты", "Место исполнения", "ДАТА ОКОНЧАНИЯ КОНТРАКТА"
        ]
    fields_str = "\n".join(f"{field}: " for field in selected_fields)
    prompt = f"""Ты анализируешь тендерную документацию. Извлеки из текста следующие данные. Если информации нет, напиши "Информация отсутствует".
Ответ должен быть строго в таком формате (каждый пункт с новой строки):
{fields_str}
Вот текст для анализа:
{content}
Извлеки данные и напиши в указанном формате."""
    answer = query_deepseek(prompt)
    lines = answer.split('\n')
    filtered = [line for line in lines if any(line.strip().startswith(f) for f in selected_fields)]
    return "\n".join(filtered) if filtered else answer

# ================= ЭНДПОЙНТЫ (защищённые) =================

# --- Проверка лицензии (для расширения) ---
@app.post("/api/verify-license")
async def verify_license_endpoint(request: Request):
    try:
        await check_access(request)
        return {"valid": True}
    except HTTPException:
        return {"valid": False, "reason": "Unauthorized"}

# --- АКТИВАЦИЯ ЛИЦЕНЗИИ ---
@app.post("/api/activate-license")
async def activate_license(data: dict):
    key = data.get("key")
    if not key:
        raise HTTPException(400, "Ключ не указан")
    result = await database.verify_and_activate_license(key)
    if not result["valid"]:
        return {"valid": False, "reason": result["reason"]}
    return {"valid": True, "expires_at": result["expires_at"]}

# --- Анализ текстов (печатных форм) ---
@app.post("/analyze_texts")
async def analyze_texts(request: Request, data: dict):
    await check_access(request)
    tenders_data = data.get("tenders", [])
    if not tenders_data:
        raise HTTPException(400, "Нет данных для анализа")
    selected_fields = data.get("fields", [])
    if not selected_fields:
        selected_fields = [
            "НАЗВАНИЕ АУКЦИОНА", "Начальная цена (НМЦ)", "ДОПОЛНИТЕЛЬНЫЕ ТРЕБОВАНИЯ К УЧАСТНИКУ",
            "ДАТА ОКОНЧАНИЯ/ПРОВЕДЕНИЯ", "Аванс", "Обеспечение заявки", "Обеспечение контракта",
            "Обеспечение гарантийных обязательств", "Контакты", "Место исполнения", "ДАТА ОКОНЧАНИЯ КОНТРАКТА"
        ]
    async def analyze_one(tender):
        reg_number = tender.get("regNumber", "")
        cached = await database.get_cached_analysis(reg_number)
        if cached:
            print(f"📦 Кэш для {reg_number} использован")
            return {"url": tender.get("url", ""), "reg_number": reg_number, "analysis": cached}
        tender_text = tender.get("text", "")
        if not tender_text or len(tender_text) < 100:
            return {"url": tender.get("url", ""), "reg_number": reg_number, "error": "Недостаточно текста"}
        start = time.time()
        analysis_result = analyze_tender_text(tender_text, selected_fields)
        await database.save_analysis_cache(reg_number, analysis_result)
        print(f"⏱️ DeepSeek обработал {reg_number} за {time.time()-start:.2f} сек")
        return {"url": tender.get("url", ""), "reg_number": reg_number, "analysis": analysis_result}
    tasks = [analyze_one(t) for t in tenders_data]
    results = await asyncio.gather(*tasks)
    doc = Document()
    doc.add_heading('РЕЗУЛЬТАТЫ АНАЛИЗА ТЕНДЕРОВ', 0)
    doc.add_paragraph(f'Дата: {datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')
    doc.add_paragraph('=' * 50)
    for item in results:
        doc.add_heading(f'Тендер: {item.get("reg_number", "Неизвестно")}', level=1)
        if "error" in item:
            doc.add_paragraph(f'❌ Ошибка: {item["error"]}')
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

def analyze_tender_text(text, selected_fields):
    if not selected_fields:
        selected_fields = [
            "НАЗВАНИЕ АУКЦИОНА", "Начальная цена (НМЦ)", "ДОПОЛНИТЕЛЬНЫЕ ТРЕБОВАНИЯ К УЧАСТНИКУ",
            "ДАТА ОКОНЧАНИЯ/ПРОВЕДЕНИЯ", "Аванс", "Обеспечение заявки", "Обеспечение контракта",
            "Обеспечение гарантийных обязательств", "Контакты", "Место исполнения", "ДАТА ОКОНЧАНИЯ КОНТРАКТА"
        ]
    fields_str = "\n".join(f"{field}: " for field in selected_fields)
    prompt = f"""Ты анализируешь тендерную документацию. Извлеки из текста следующие данные. Если информации нет, напиши "Информация отсутствует".
Ответ должен быть строго в таком формате (каждый пункт с новой строки):
{fields_str}
Вот текст для анализа:
{text[:8000]}
Извлеки данные и напиши в указанном формате."""
    answer = query_deepseek(prompt)
    result = {}
    for line in answer.split('\n'):
        if ':' in line:
            k, v = line.split(':', 1)
            result[k.strip()] = v.strip()
    return result

# --- Упаковка файлов ---
def detect_file_type(content: bytes) -> str:
    if content.startswith(b'%PDF'): return 'pdf'
    if content.startswith(b'PK\x03\x04') or content.startswith(b'PK\x05\x06') or content.startswith(b'PK\x07\x08'):
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                files = zf.namelist()
                if any(f.startswith('word/') for f in files): return 'docx'
                if any(f.startswith('xl/') for f in files): return 'xlsx'
                return 'zip'
        except:
            return 'zip'
    if content.startswith(b'Rar!\x1a\x07\x00') or content.startswith(b'Rar!\x1a\x07\x01\x00') or content.startswith(b'Rar!'): return 'rar'
    if content.startswith(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'):
        if b'WorkBook' in content[:2000] or b'BOUNDSHEET' in content[:2000]: return 'xls'
        return 'doc'
    if content.startswith(b'7z\xbc\xaf\x27\x1c'): return '7z'
    if content.startswith(b'\x89PNG\r\n\x1a\n'): return 'png'
    if content.startswith(b'\xff\xd8\xff'): return 'jpg'
    if content.startswith(b'{\\rtf'): return 'rtf'
    return 'unknown'

@app.post("/package_files")
async def package_files(request: Request, files: list[UploadFile] = File(...), analysis_text: str = Form("")):
    await check_access(request)
    if not files:
        raise HTTPException(400, "Нет файлов")
    tenders = {}
    for file in files:
        content = await file.read()
        parts = file.filename.split('_', 1)
        tender_id = parts[0] if len(parts) == 2 else "без_тендера"
        original_name = parts[1] if len(parts) == 2 else file.filename
        file_type = detect_file_type(content)
        base = os.path.splitext(original_name)[0] or 'file'
        ext_map = {
            'pdf': '.pdf', 'xlsx': '.xlsx', 'xls': '.xls', 'docx': '.docx',
            'rar': '.rar', 'zip': '.zip', '7z': '.7z', 'png': '.png',
            'jpg': '.jpg', 'rtf': '.rtf'
        }
        new_ext = ext_map.get(file_type)
        if new_ext:
            original_name = base + new_ext
        else:
            if '.' not in original_name:
                original_name = base + '.bin'
        print(f"🔍 Тип: {file_type}, имя: {original_name}")
        tenders.setdefault(tender_id, []).append((original_name, content))
    for tender_id, file_list in tenders.items():
        seen = set()
        new_list = []
        for name, content in file_list:
            base, ext = os.path.splitext(name)
            counter = 1
            new_name = name
            while new_name in seen:
                new_name = f"{base}_{counter}{ext}"
                counter += 1
            seen.add(new_name)
            new_list.append((new_name, content))
        tenders[tender_id] = new_list
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zf:
        for tender_id, file_list in tenders.items():
            folder = f"Тендер_{tender_id}"
            for name, content in file_list:
                zf.writestr(os.path.join(folder, name), content)
        if analysis_text:
            doc = Document()
            doc.add_heading('РЕЗУЛЬТАТЫ АНАЛИЗА ТЕНДЕРОВ', 0)
            doc.add_paragraph(f'Дата: {datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')
            doc.add_paragraph('=' * 50)
            sections = analysis_text.split('=== Тендер')
            for s in sections:
                if s.strip():
                    doc.add_paragraph(s.strip())
            word_buf = io.BytesIO()
            doc.save(word_buf)
            word_buf.seek(0)
            zf.writestr("анализ_тендеров.docx", word_buf.getvalue())
    zip_buffer.seek(0)
    filename = f"результаты_анализа_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    encoded = quote(filename)
    return Response(zip_buffer.getvalue(), media_type="application/zip", headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded}"})

# ================= ЭНДПОЙНТЫ ДЛЯ ПРОБНОГО ПЕРИОДА =================
@app.post("/api/trial/start")
async def start_trial(data: dict):
    device_id = data.get("device_id")
    if not device_id:
        raise HTTPException(400, "device_id не указан")
    result = await database.start_trial(device_id, trial_days=2)
    return result

@app.post("/api/trial/status")
async def check_trial(data: dict):
    device_id = data.get("device_id")
    if not device_id:
        raise HTTPException(400, "device_id не указан")
    result = await database.get_trial_status(device_id)
    return result

# ================= ОСТАЛЬНЫЕ ЭНДПОЙНТЫ =================
@app.post("/search_tenders")
async def search_tenders(request: Request, data: dict):
    await check_access(request)
    query = data.get("query", "").strip()
    limit = data.get("limit", MAX_TENDERS)
    if not query:
        raise HTTPException(400, "Введите ключевые слова для поиска")
    tender_urls = parser.search_tenders_zakupki(query, limit)
    if not tender_urls:
        return {"detail": "Тендеры по вашему запросу не найдены"}
    base_dir = f"search_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    os.makedirs(base_dir, exist_ok=True)
    results = []
    for idx, tender_url in enumerate(tender_urls[:MAX_TENDERS], 1):
        tender_name = f"Тендер_{idx}"
        tender_dir = os.path.join(base_dir, tender_name)
        os.makedirs(tender_dir, exist_ok=True)
        files = parser.download_files_from_tender(tender_url, tender_dir)
        if files:
            combined_text = ""
            for file_path in files:
                text = read_docx(file_path) if file_path.endswith('.docx') else read_txt(file_path) if file_path.endswith('.txt') else read_excel(file_path)
                if text and not text.startswith("Ошибка"):
                    combined_text += text + "\n"
            results.append({
                "tender_name": tender_name,
                "files": files,
                "text": combined_text[:5000]
            })
        time.sleep(2)
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zf:
        for res in results:
            doc = Document()
            doc.add_heading(f'Анализ тендера: {res["tender_name"]}', 0)
            doc.add_paragraph(f'Дата анализа: {datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')
            doc.add_paragraph('=' * 50)
            doc.add_paragraph(res["text"] if res["text"] else "Не удалось извлечь текст из документов")
            word_buf = io.BytesIO()
            doc.save(word_buf)
            word_buf.seek(0)
            zf.writestr(f"{res['tender_name']}_результат.docx", word_buf.getvalue())
    zip_buffer.seek(0)
    filename = f"тендеры_по_запросу_{query[:20]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    encoded = quote(filename)
    if os.path.exists(base_dir):
        shutil.rmtree(base_dir)
    return Response(zip_buffer.getvalue(), media_type="application/zip", headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded}"})

@app.post("/suggest_keywords")
async def suggest_keywords(request: Request, data: dict):
    await check_access(request)
    description = data.get("description", "").strip()
    if not description:
        raise HTTPException(400, "Описание не может быть пустым")
    prompt = f"""Ты — помощник по тендерам. Пользователь описал свою деятельность:
"{description}"
Выдели 5–7 ключевых слов для поиска тендеров на zakupki.gov.ru.
Ключевые слова должны быть конкретными (например, "строительство школы", "поставка медоборудования", "ремонт дорог").
Выдай ТОЛЬКО список слов через запятую, без лишнего текста."""
    answer = query_deepseek(prompt)
    keywords = [kw.strip() for kw in answer.replace('\n', ',').split(',') if kw.strip()]
    return {"keywords": keywords[:7]}

@app.post("/ask_ai")
async def ask_ai(request: Request, data: dict):
    await check_access(request)
    question = data.get("question", "").strip()
    if not question:
        raise HTTPException(400, "Введите вопрос")
    prompt = f"""Ты — консультант по тендерам и бизнес-процессам. Ответь на вопрос пользователя чётко, по делу, без воды.
Вопрос пользователя: {question}
Дай ответ в виде текста (2-4 предложения), который будет полезен для бизнеса."""
    answer = query_deepseek(prompt)
    return {"answer": answer}

# ================= ЗАПУСК =================
if __name__ == "__main__":
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
