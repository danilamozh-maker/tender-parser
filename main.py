import os
import shutil
import zipfile
import io
import json
import re
import time
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

app = FastAPI()

# Инициализация базы данных
database.init_db()

# ================= НАСТРОЙКИ =================
OLLAMA_API_URL = "https://api.kodikrouter.ru/v1/chat/completions"
API_KEY = "sk-kr_live_E-JvaZzvEh-AnkSjO6d35qcAJ7RCysKt" # ← ВСТАВЬ СВОЙ КЛЮЧ!
MODEL_NAME = "deepseek/deepseek-chat"
MAX_TENDERS = 15
# ============================================

# ================= АВТОРИЗАЦИЯ =================
def get_current_user(request: Request):
    email = request.cookies.get("user_email")
    if not email:
        return None
    user = database.get_user(email)
    return user

# ================= СТРАНИЦЫ АВТОРИЗАЦИИ (без Jinja2) =================
@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    error = request.query_params.get("error", "")
    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Вход</title>
        <style>
            body { font-family: Arial, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; display: flex; justify-content: center; align-items: center; }
            .card { background: white; padding: 40px; border-radius: 20px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); max-width: 400px; width: 100%; }
            h1 { margin-bottom: 20px; color: #333; }
            input { width: 100%; padding: 12px; margin-bottom: 15px; border: 2px solid #e2e8f0; border-radius: 8px; font-size: 16px; }
            button { width: 100%; padding: 14px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 10px; font-size: 18px; font-weight: 600; cursor: pointer; }
            .error { color: #ef4444; margin-bottom: 15px; padding: 10px; background: #fee2e2; border-radius: 8px; }
            .link { margin-top: 15px; text-align: center; font-size: 14px; }
            a { color: #667eea; text-decoration: none; font-weight: 600; }
        </style>
    </head>
    <body>
        <div class="card">
            <h1>🔐 Вход</h1>
            """ + (f'<div class="error">{error}</div>' if error else "") + """
            <form method="post">
                <input type="email" name="email" placeholder="Почта" required>
                <input type="password" name="password" placeholder="Пароль" required>
                <button type="submit">Войти</button>
            </form>
            <div class="link">Нет аккаунта? <a href="/register">Зарегистрироваться</a></div>
        </div>
    </body>
    </html>
    """
    return HTMLResponse(content=html)

@app.post("/login", response_class=HTMLResponse)
async def login_post(request: Request, email: str = Form(...), password: str = Form(...)):
    user = database.get_user(email)
    if not user:
        return RedirectResponse(url="/login?error=Пользователь не найден", status_code=302)
    if not database.verify_password(password, user["hashed_password"]):
        return RedirectResponse(url="/login?error=Неверный пароль", status_code=302)
    response = RedirectResponse(url="/", status_code=302)
    response.set_cookie(key="user_email", value=email, httponly=True, max_age=3600*24*7)
    return response

@app.get("/register", response_class=HTMLResponse)
async def register_page(request: Request):
    error = request.query_params.get("error", "")
    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Регистрация</title>
        <style>
            body { font-family: Arial, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; display: flex; justify-content: center; align-items: center; }
            .card { background: white; padding: 40px; border-radius: 20px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); max-width: 400px; width: 100%; }
            h1 { margin-bottom: 20px; color: #333; }
            input { width: 100%; padding: 12px; margin-bottom: 15px; border: 2px solid #e2e8f0; border-radius: 8px; font-size: 16px; }
            button { width: 100%; padding: 14px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 10px; font-size: 18px; font-weight: 600; cursor: pointer; }
            .error { color: #ef4444; margin-bottom: 15px; padding: 10px; background: #fee2e2; border-radius: 8px; }
            .link { margin-top: 15px; text-align: center; font-size: 14px; }
            a { color: #667eea; text-decoration: none; font-weight: 600; }
        </style>
    </head>
    <body>
        <div class="card">
            <h1>📝 Регистрация</h1>
            """ + (f'<div class="error">{error}</div>' if error else "") + """
            <form method="post">
                <input type="email" name="email" placeholder="Почта" required>
                <input type="password" name="password" placeholder="Пароль" required>
                <button type="submit">Зарегистрироваться</button>
            </form>
            <div class="link">Уже есть аккаунт? <a href="/login">Войти</a></div>
        </div>
    </body>
    </html>
    """
    return HTMLResponse(content=html)

@app.post("/register", response_class=HTMLResponse)
async def register_post(request: Request, email: str = Form(...), password: str = Form(...)):
    if len(password) < 6:
        return RedirectResponse(url="/register?error=Пароль должен быть не менее 6 символов", status_code=302)
    success = database.create_user(email, password)
    if not success:
        return RedirectResponse(url="/register?error=Пользователь с такой почтой уже существует", status_code=302)
    return RedirectResponse(url="/login", status_code=302)

@app.get("/logout")
async def logout():
    response = RedirectResponse(url="/login", status_code=302)
    response.delete_cookie("user_email")
    return response

# ================= ФУНКЦИИ ЧТЕНИЯ ФАЙЛОВ =================
def read_docx(file_path):
    try:
        doc = Document(file_path)
        full_text = []
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                full_text.append(paragraph.text)
        return "\n".join(full_text)
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
        text = "\n".join([str(cell) for row in rows for cell in row if cell])
        return text
    except:
        try:
            wb = xlrd.open_workbook(file_path)
            sheet = wb.sheet_by_index(0)
            text = ""
            for row in range(sheet.nrows):
                row_text = []
                for col in range(sheet.ncols):
                    cell = sheet.cell_value(row, col)
                    if cell:
                        row_text.append(str(cell))
                if row_text:
                    text += " ".join(row_text) + "\n"
            return text
        except Exception as e:
            return f"Ошибка чтения Excel: {e}"

# ================= ЗАПРОС К KODIKROUTER =================
def query_kodik(prompt):
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
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    try:
        response = requests.post(OLLAMA_API_URL, json=payload, headers=headers, timeout=600)
        if response.status_code == 200:
            result = response.json()
            return result["choices"][0]["message"]["content"]
        else:
            return f"Ошибка HTTP {response.status_code}: {response.text}"
    except requests.exceptions.ConnectionError:
        return "❌ Не удаётся подключиться к KodikRouter! Проверь интернет."
    except Exception as e:
        return f"❌ Ошибка: {e}"

# ================= АНАЛИЗ ФАЙЛА =================
def analyze_file(file_path, selected_fields):
    ext = Path(file_path).suffix.lower()
    if ext == ".docx":
        content = read_docx(file_path)
    elif ext == ".txt":
        content = read_txt(file_path)
    elif ext == ".xlsx" or ext == ".xls":
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
            "НАЗВАНИЕ АУКЦИОНА",
            "Начальная цена (НМЦ)",
            "ДОПОЛНИТЕЛЬНЫЕ ТРЕБОВАНИЯ К УЧАСТНИКУ",
            "ДАТА ОКОНЧАНИЯ/ПРОВЕДЕНИЯ",
            "Аванс",
            "Обеспечение заявки",
            "Обеспечение контракта",
            "Обеспечение гарантийных обязательств",
            "Контакты",
            "Место исполнения",
            "ДАТА ОКОНЧАНИЯ КОНТРАКТА"
        ]
    fields_str = "\n".join([f"{field}: " for field in selected_fields])
    prompt = f"""Ты анализируешь тендерную документацию. Извлеки из текста следующие данные. Если информации нет, напиши "Информация отсутствует".

Ответ должен быть строго в таком формате (каждый пункт с новой строки):

{fields_str}

Вот текст для анализа:
{content}

Извлеки данные и напиши в указанном формате."""
    answer = query_kodik(prompt)
    lines = answer.split('\n')
    filtered_lines = []
    for line in lines:
        for field in selected_fields:
            if line.strip().startswith(field):
                filtered_lines.append(line)
                break
    if not filtered_lines:
        return answer
    return "\n".join(filtered_lines)

# ================= НОВЫЕ ФУНКЦИИ: ПАРСИНГ ZAKUPKI.GOV.RU =================
def search_tenders_on_zakupki(query, limit=MAX_TENDERS):
    """Ищет тендеры на zakupki.gov.ru по ключевому слову и возвращает список ссылок."""
    tenders = []
    search_url = "https://zakupki.gov.ru/epz/order/extendedsearch/results.html"
    params = {
        "searchString": query,
        "pageNumber": 1,
        "recordsPerPage": limit,
        "fz44": "on",
        "fz223": "on",
        "sortDirection": "false"
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    try:
        response = requests.get(search_url, params=params, headers=headers, timeout=30)
        if response.status_code != 200:
            print(f"Ошибка поиска: HTTP {response.status_code}")
            return []
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(response.text, 'html.parser')
        # Ищем ссылки на карточки тендеров
        for link in soup.find_all('a', href=True):
            href = link.get('href')
            if href and 'order/view/' in href:
                full_url = href if href.startswith('http') else f"https://zakupki.gov.ru{href}"
                if full_url not in tenders:
                    tenders.append(full_url)
        print(f"Найдено {len(tenders)} ссылок на тендеры")
        return tenders[:limit]
    except Exception as e:
        print(f"Ошибка при поиске тендеров: {e}")
        return []

def download_files_from_tender(tender_url, download_dir):
    """Скачивает все файлы по тендеру и сохраняет их в указанную папку."""
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        response = requests.get(tender_url, headers=headers, timeout=30)
        if response.status_code != 200:
            return []
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(response.text, 'html.parser')
        downloaded_files = []
        # Ищем ссылки на файлы
        for link in soup.find_all('a', href=True):
            href = link.get('href')
            if href and any(href.endswith(ext) for ext in ['.pdf', '.docx', '.xlsx', '.doc', '.xls', '.rtf', '.txt']):
                # Пропускаем слишком большие файлы (больше 20 МБ)
                full_url = href if href.startswith('http') else f"https://zakupki.gov.ru{href}"
                try:
                    file_response = requests.get(full_url, headers=headers, stream=True, timeout=60)
                    if file_response.status_code == 200:
                        # Генерируем имя файла
                        filename = f"{len(downloaded_files)+1}_{Path(full_url).name}"
                        file_path = os.path.join(download_dir, filename)
                        with open(file_path, 'wb') as f:
                            for chunk in file_response.iter_content(chunk_size=8192):
                                f.write(chunk)
                        downloaded_files.append(file_path)
                        time.sleep(0.5) # пауза между скачиваниями
                except Exception as e:
                    print(f"Ошибка при скачивании файла: {e}")
        return downloaded_files
    except Exception as e:
        print(f"Ошибка при загрузке страницы тендера: {e}")
        return []

# ================= НОВЫЙ ЭНДПОЙНТ: ПОИСК ТЕНДЕРОВ =================
@app.post("/search_tenders")
async def search_tenders(request: Request, data: dict):
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Не авторизован")
    
    query = data.get("query", "").strip()
    limit = data.get("limit", MAX_TENDERS)
    if not query:
        raise HTTPException(status_code=400, detail="Введите ключевые слова для поиска")
    
    # 1. Ищем тендеры
    tender_urls = search_tenders_on_zakupki(query, limit)
    if not tender_urls:
        return {"detail": "Тендеры по вашему запросу не найдены"}
    
    # 2. Создаём временную папку для всех тендеров
    base_dir = f"search_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    os.makedirs(base_dir, exist_ok=True)
    
    results = []
    for idx, tender_url in enumerate(tender_urls[:MAX_TENDERS], 1):
        tender_name = f"Тендер_{idx}"
        tender_dir = os.path.join(base_dir, tender_name)
        os.makedirs(tender_dir, exist_ok=True)
        files = download_files_from_tender(tender_url, tender_dir)
        if files:
            # 3. Анализируем скачанные файлы
            combined_text = ""
            for file_path in files:
                text = read_docx(file_path) if file_path.endswith('.docx') else read_txt(file_path) if file_path.endswith('.txt') else read_excel(file_path)
                if text and not text.startswith("Ошибка"):
                    combined_text += text + "\n"
            # Сохраняем результат для отчёта
            results.append({
                "tender_name": tender_name,
                "files": files,
                "text": combined_text[:5000] # ограничиваем длину
            })
        time.sleep(2) # пауза между тендерами
    
    # 4. Формируем ZIP-архив с отчётами
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for res in results:
            doc = Document()
            doc.add_heading(f'Анализ тендера: {res["tender_name"]}', 0)
            doc.add_paragraph(f'Дата анализа: {datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')
            doc.add_paragraph('=' * 50)
            doc.add_paragraph(res["text"] if res["text"] else "Не удалось извлечь текст из документов")
            word_buffer = io.BytesIO()
            doc.save(word_buffer)
            word_buffer.seek(0)
            zip_file.writestr(f"{res['tender_name']}_результат.docx", word_buffer.getvalue())
    
    zip_buffer.seek(0)
    filename = f"тендеры_по_запросу_{query[:20]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    encoded_filename = quote(filename)
    # Удаляем временную папку
    if os.path.exists(base_dir):
        shutil.rmtree(base_dir)
    return Response(
        zip_buffer.getvalue(),
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"}
    )

# ================= НОВЫЙ ЭНДПОЙНТ: ПОДБОР КЛЮЧЕВЫХ СЛОВ ЧЕРЕЗ ИИ =================
@app.post("/suggest_keywords")
async def suggest_keywords(request: Request, data: dict):
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Не авторизован")
    
    description = data.get("description", "").strip()
    if not description:
        raise HTTPException(status_code=400, detail="Описание не может быть пустым")
    
    prompt = f"""Ты — помощник по тендерам. Пользователь описал свою деятельность:
"{description}"

Выдели 5–7 ключевых слов для поиска тендеров на zakupki.gov.ru.
Ключевые слова должны быть конкретными (например, "строительство школы", "поставка медоборудования", "ремонт дорог").
Выдай ТОЛЬКО список слов через запятую, без лишнего текста."""
    
    answer = query_kodik(prompt)
    # Парсим ответ ИИ в список ключевых слов
    keywords = [kw.strip() for kw in answer.replace('\n', ',').split(',') if kw.strip()]
    return {"keywords": keywords[:7]}

# ================= ГЛАВНАЯ СТРАНИЦА =================
@app.get("/", response_class=HTMLResponse)
async def main(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=302)
    with open("templates/index.html", "r", encoding="utf-8") as f:
        html_content = f.read()
    return HTMLResponse(content=html_content)

# ================= ЭНДПОЙНТ ДЛЯ АНАЛИЗА ФАЙЛОВ (ручной) =================
@app.post("/analyze")
async def analyze_files(
    request: Request,
    files: list[UploadFile] = File(...),
    fields: str = Form("")
):
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Не авторизован")
    if not files:
        raise HTTPException(400, "Нет файлов")
    try:
        selected_fields = json.loads(fields) if fields else []
    except:
        selected_fields = []
    temp_dir = f"temp_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    os.makedirs(temp_dir, exist_ok=True)
    results = []
    error_files = []
    try:
        for file in files:
            if not file.filename.endswith(('.docx', '.txt', '.xlsx', '.xls')):
                error_files.append(f"{file.filename} (неподдерживаемый формат)")
                continue
            file_path = os.path.join(temp_dir, file.filename)
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            result = analyze_file(file_path, selected_fields)
            results.append({"filename": file.filename, "result": result})
        if error_files:
            results.append({"filename": "ОШИБКИ", "result": "\n".join(error_files)})
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            doc = Document()
            doc.add_heading('РЕЗУЛЬТАТЫ АНАЛИЗА ТЕНДЕРОВ', 0)
            doc.add_paragraph(f'Дата анализа: {datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')
            doc.add_paragraph('=' * 50)
            for r in results:
                doc.add_heading(f'Файл: {r["filename"]}', level=1)
                doc.add_paragraph(r['result'])
                doc.add_page_break()
            word_buffer = io.BytesIO()
            doc.save(word_buffer)
            word_buffer.seek(0)
            zip_file.writestr(
                f"тендеры_результат_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                word_buffer.getvalue()
            )
            for file in files:
                if file.filename.endswith(('.docx', '.txt', '.xlsx', '.xls')):
                    file_path = os.path.join(temp_dir, file.filename)
                    if os.path.exists(file_path):
                        zip_file.write(file_path, file.filename)
        zip_buffer.seek(0)
        filename = f"результаты_анализа_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        encoded_filename = quote(filename)
        return Response(
            zip_buffer.getvalue(),
            media_type="application/zip",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"}
        )
    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

if __name__ == "__main__":
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
