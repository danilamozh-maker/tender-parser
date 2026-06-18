import os
import shutil
import zipfile
import io
from pathlib import Path
from datetime import datetime
from docx import Document
import requests
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, Response
from starlette.requests import Request
import uvicorn

app = FastAPI()

# ================= НАСТРОЙКИ =================
OLLAMA_API_URL = "https://api.kodikrouter.ru/v1/chat/completions"
API_KEY = "sk-kr_live_E-JvaZzvEh-AnkSjO6d35qcAJ7RCysKt" # ← ВСТАВЬ СВОЙ КЛЮЧ!
MODEL_NAME = "deepseek/deepseek-chat"
# ============================================

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
        response = requests.post(OLLAMA_API_URL, json=payload, headers=headers, timeout=300)
        
        if response.status_code == 200:
            result = response.json()
            return result["choices"][0]["message"]["content"]
        else:
            return f"Ошибка HTTP {response.status_code}: {response.text}"
            
    except requests.exceptions.ConnectionError:
        return "❌ Не удаётся подключиться к KodikRouter! Проверь интернет."
    except Exception as e:
        return f"❌ Ошибка: {e}"

def analyze_file(file_path):
    ext = Path(file_path).suffix.lower()
    
    if ext == ".docx":
        content = read_docx(file_path)
    elif ext == ".txt":
        content = read_txt(file_path)
    else:
        return f"⚠️ Неподдерживаемый формат: {ext}"
    
    if not content or "Ошибка чтения" in content:
        return f"⚠️ Не удалось прочитать файл"
    
    content = content[:8000]
    
    if not content.strip():
        return "⚠️ Файл пустой"
    
    prompt = f"""Ты анализируешь тендерную документацию. Извлеки из текста следующие данные. Если информации нет, напиши "Информация отсутствует".

Ответ должен быть строго в таком формате (каждый пункт с новой строки):

НАЗВАНИЕ АУКЦИОНА: 
Начальная цена (НМЦ): 
ДОПОЛНИТЕЛЬНЫЕ ТРЕБОВАНИЯ К УЧАСТНИКУ: 
ДАТА ОКОНЧАНИЯ/ПРОВЕДЕНИЯ: 
Аванс: 
Обеспечение заявки: 
Обеспечение контракта: 
Обеспечение гарантийных обязательств: 
Контакты: 
Место исполнения: 
ДАТА ОКОНЧАНИЯ КОНТРАКТА: 

Вот текст для анализа:
{content}

Извлеки данные и напиши в указанном формате."""
    
    answer = query_kodik(prompt)
    return answer

# ================= ГЛАВНАЯ СТРАНИЦА (без Jinja2) =================
@app.get("/", response_class=HTMLResponse)
async def main():
    with open("templates/index.html", "r", encoding="utf-8") as f:
        html_content = f.read()
    return HTMLResponse(content=html_content)

# ================= ЭНДПОЙНТ ДЛЯ АНАЛИЗА =================
@app.post("/analyze")
async def analyze_files(files: list[UploadFile] = File(...)):
    if not files:
        raise HTTPException(400, "Нет файлов")
    
    temp_dir = f"temp_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    os.makedirs(temp_dir, exist_ok=True)
    
    results = []
    error_files = []
    
    try:
        for file in files:
            if not file.filename.endswith(('.docx', '.txt')):
                error_files.append(f"{file.filename} (неподдерживаемый формат)")
                continue
            
            file_path = os.path.join(temp_dir, file.filename)
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            
            result = analyze_file(file_path)
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
            result_text = ""
            for r in results:
                result_text += f"\n{'='*60}\n"
                result_text += f"ФАЙЛ: {r['filename']}\n"
                result_text += f"{'='*60}\n\n"
                result_text += r['result']
                result_text += "\n\n"
            
            zip_file.writestr(
                f"тендеры_результат_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                result_text
            )
            
            for file in files:
                if file.filename.endswith(('.docx', '.txt')):
                    file_path = os.path.join(temp_dir, file.filename)
                    if os.path.exists(file_path):
                        zip_file.write(file_path, file.filename)
        
        zip_buffer.seek(0)
        return Response(
            zip_buffer.getvalue(),
            media_type="application/zip",
            headers={"Content-Disposition": f"attachment; filename=результаты_анализа_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"}
        )
    
    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

if __name__ == "__main__":
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
