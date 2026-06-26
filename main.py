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
from fastapi.responses import HTMLResponse, RedirectResponse
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

# ================= ИНИЦИАЛИЗАЦИЯ БД =================
@app.on_event("startup")
async def startup():
    await database.init_db()
    print("✅ База данных PostgreSQL инициализирована")

# ================= ПРОВЕРКА ЛИЦЕНЗИИ =================
async def verify_license_from_request(request: Request):
    license_key = request.headers.get("X-License-Key")
    if not license_key:
        raise HTTPException(401, detail="License key required")
    result = await database.verify_license(license_key)
    if result is None:
        raise HTTPException(401, detail="License not found")
    if not result.get("valid"):
        raise HTTPException(401, detail=result.get("reason", "Invalid license"))
    return license_key

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

# ================= ЭНДПОЙНТЫ =================

# --- Создание лицензии (для Tilda) ---
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

# --- Проверка лицензии (для расширения) ---
@app.post("/api/verify-license")
async def verify_license_endpoint(request: Request):
    try:
        await verify_license_from_request(request)
        return {"valid": True}
    except HTTPException as e:
        return {"valid": False, "reason": e.detail}

# --- Анализ текстов (печатных форм) ---
@app.post("/analyze_texts")
async def analyze_texts(request: Request, data: dict):
    await verify_license_from_request(request)
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
        # Проверяем кэш
        cached = await database.get_cached_analysis(reg_number)
        if cached:
            print(f"📦 Кэш для {reg_number} использован")
            return {"url": tender.get("url", ""), "reg_number": reg_number, "analysis": cached}
        tender_text = tender.get("text", "")
        if not tender_text or len(tender_text) < 100:
            return {"url": tender.get("url", ""), "reg_number": reg_number, "error": "Недостаточно текста"}
        start = time.time()
        analysis_result = analyze_tender_text(tender_text, selected_fields)
        # Сохраняем в кэш
        await database.save_analysis_cache(reg_number, analysis_result)
        print(f"⏱️ DeepSeek обработал {reg_number} за {time.time()-start:.2f} сек")
        return {"url": tender.get("url", ""), "reg_number": reg_number, "analysis": analysis_result}
    tasks = [analyze_one(t) for t in tenders_data]
    results = await asyncio.gather(*tasks)
    # Формируем DOCX
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
    await verify_license_from_request(request)
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
    # Дедупликация имён внутри каждого тендера
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
    # Создаём ZIP
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

# ================= ЗАПУСК =================
if __name__ == "__main__":
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
