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
import parser
from bs4 import BeautifulSoup

app = FastAPI()

# Инициализация базы данных
database.init_db()

# ================= НАСТРОЙКИ =================
OLLAMA_API_URL = "https://api.kodikrouter.ru/v1/chat/completions"
API_KEY = "sk-kr_live_E-JvaZzvEh-AnkSjO6d35qcAJ7RCysKt"
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

# ================= СТРАНИЦЫ АВТОРИЗАЦИИ =================
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

# ================= ПОИСК ТЕНДЕРОВ =================
@app.post("/search_tenders")
async def search_tenders(request: Request, data: dict):
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Не авторизован")
    
    query = data.get("query", "").strip()
    limit = data.get("limit", MAX_TENDERS)
    if not query:
        raise HTTPException(status_code=400, detail="Введите ключевые слова для поиска")
    
    tender_urls = parser.search_tenders_zakupki(query, limit)
    if not tender_urls:
        return {"detail": "Тендеры по ваш
