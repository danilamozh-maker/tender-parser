import os
import shutil
import zipfile
import io
import json
from pathlib import Path
from datetime import datetime
from docx import Document
import requests
from fastapi import FastAPI, UploadFile, File, HTTPException, Form, Request, Response, Depends
from fastapi.responses import HTMLResponse, Response as FastAPIResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from urllib.parse import quote
import openpyxl
import xlrd
import uvicorn
import database

app = FastAPI()

# Инициализация базы данных
database.init_db()

# Шаблоны
templates = Jinja2Templates(directory="templates")
import os
print("Текущая директория:", os.getcwd())
print("Содержимое папки templates:", os.listdir("templates") if os.path.exists("templates") else "❌ Папка templates ОТСУТСТВУЕТ!")

# ================= НАСТРОЙКИ =================
OLLAMA_API_URL = "https://api.kodikrouter.ru/v1/chat/completions"
API_KEY = "sk-kr_live_E-JvaZzvEh-AnkSjO6d35qcAJ7RCysKt" # ← ВСТАВЬ СВОЙ КЛЮЧ!
MODEL_NAME = "deepseek/deepseek-chat"
# ============================================

# ================= АВТОРИЗАЦИЯ =================
def get_current_user(request: Request):
    email = request.cookies.get("user_email")
    if not email:
        return None
    user = database.get_user(email)
    return user

@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    return templates.TemplateResponse("login.html", {"request": {}, "error": None})

@app.post("/login", response_class=HTMLResponse)
async def login_post(request: Request, email: str = Form(...), password: str = Form(...)):
    user = database.get_user(email)
    if not user:
        return templates.TemplateResponse("login.html", {"request": {}, "error": "Пользователь не найден"})
    
    if not database.verify_password(password, user["hashed_password"]):
        return templates.TemplateResponse("login.html", {"request": {}, "error": "Неверный пароль"})
    
    response = RedirectResponse(url="/", status_code=302)
    response.set_cookie(key="user_email", value=email, httponly=True, max_age=3600*24*7)
    return response

@app.get("/register", response_class=HTMLResponse)
async def register_page(request: Request):
    return templates.TemplateResponse("register.html", {"request": {}, "error": None})

@app.post("/register", response_class=HTMLResponse)
async def register_post(request: Request, email: str = Form(...), password: str = Form(...)):
    if len(password) < 6:
        return templates.TemplateResponse("register.html", {"request": {}, "error": "Пароль должен быть не менее 6 символов"})
    
    success = database.create_user(email, password)
    if not success:
        return templates.TemplateResponse("register.html", {"request": {}, "error": "Пользователь с такой почтой уже существует"})
    
    response = RedirectResponse(url="/login", status_code=302)
    return response

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

# ================= ГЛАВНАЯ СТРАНИЦА =================
@app.get("/", response_class=HTMLResponse)
async def main(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=302)
    with open("templates/index.html", "r", encoding="utf-8") as f:
        html_content = f.read()
    return HTMLResponse(content=html_content)

# ================= ЭНДПОЙНТ ДЛЯ АНАЛИЗА =================
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
            results.append({
                "filename": file.filename,
                "result": result
            })
        
        if error_files:
            results.append({
                "filename": "ОШИБКИ",
                "result": "\n".join(error_files)
            })
        
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
