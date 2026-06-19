import os
import time
import requests
from bs4 import BeautifulSoup
from urllib.parse import quote
from pathlib import Path

# ===== НАСТРОЙКИ =====
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
    "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
}

# ===== ФУНКЦИЯ ДЛЯ ЗАГРУЗКИ ЧЕРЕЗ ALLORIGINS =====
def fetch_via_allorigins(url):
    """
    Загружает страницу через AllOrigins (бесплатный прокси-сервис).
    Возвращает HTML-код страницы.
    """
    try:
        encoded_url = quote(url, safe='')
        proxy_url = f"https://api.allorigins.win/raw?url={encoded_url}"
        response = requests.get(proxy_url, headers=HEADERS, timeout=30)
        if response.status_code == 200:
            return response.text
        else:
            print(f"❌ Ошибка AllOrigins: {response.status_code}")
            return None
    except Exception as e:
        print(f"❌ Ошибка при запросе через AllOrigins: {e}")
        return None

# ===== ПОИСК ТЕНДЕРОВ НА ZAKUPKI.GOV.RU =====
def search_tenders_zakupki(query, limit=10):
    """
    Ищет тендеры на zakupki.gov.ru по ключевым словам.
    Использует AllOrigins для обхода блокировки.
    """
    print(f"🔍 Ищем тендеры по запросу: {query}")
    
    search_url = "https://zakupki.gov.ru/epz/order/extendedsearch/results.html"
    params = {
        "searchString": query,
        "pageNumber": "1",
        "recordsPerPage": str(limit),
        "fz44": "on",
        "fz223": "",
        "orderPlacementType": "ALL",
        "sortBy": "P_DATE",
        "sortDirection": "false"
    }
    
    # Формируем полный URL с параметрами
    full_url = search_url + "?" + "&".join([f"{k}={v}" for k, v in params.items()])
    
    # Загружаем через AllOrigins
    html = fetch_via_allorigins(full_url)
    if not html:
        print("❌ Не удалось загрузить страницу через AllOrigins")
        return []
    
    try:
        soup = BeautifulSoup(html, 'html.parser')
        tenders = []
        
        for link in soup.find_all('a', href=True):
            href = link.get('href')
            if href and 'epz/order/view/' in href:
                full_url = href if href.startswith('http') else f"https://zakupki.gov.ru{href}"
                if full_url not in tenders:
                    tenders.append(full_url)
                    print(f"📎 Найдена ссылка: {full_url}")
        
        print(f"✅ Найдено {len(tenders)} тендеров")
        return tenders[:limit]
        
    except Exception as e:
        print(f"❌ Ошибка при парсинге: {e}")
        return []

# ===== СКАЧИВАНИЕ ФАЙЛОВ ПО ТЕНДЕРУ =====
def download_files_from_tender(tender_url, download_dir):
    """
    Скачивает все доступные файлы по тендеру и сохраняет в указанную папку.
    Возвращает список путей к скачанным файлам.
    """
    print(f"📥 Скачиваем файлы из: {tender_url}")
    
    # Загружаем страницу тендера через AllOrigins
    html = fetch_via_allorigins(tender_url)
    if not html:
        print("❌ Не удалось загрузить страницу тендера")
        return []
    
    soup = BeautifulSoup(html, 'html.parser')
    downloaded_files = []
    
    for link in soup.find_all('a', href=True):
        href = link.get('href')
        if href and any(href.endswith(ext) for ext in ['.pdf', '.docx', '.xlsx', '.doc', '.xls', '.rtf', '.txt']):
            full_url = href if href.startswith('http') else f"https://zakupki.gov.ru{href}"
            try:
                # Пробуем скачать напрямую (без прокси, так как сам файл может быть доступен)
                file_response = requests.get(full_url, headers=HEADERS, stream=True, timeout=60)
                if file_response.status_code == 200:
                    filename = f"{len(downloaded_files)+1}_{Path(full_url).name}"
                    file_path = os.path.join(download_dir, filename)
                    with open(file_path, 'wb') as f:
                        for chunk in file_response.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)
                    downloaded_files.append(file_path)
                    print(f" ✅ Скачан: {filename}")
                    time.sleep(0.3)
                else:
                    print(f" ❌ Ошибка скачивания: {file_response.status_code}")
            except Exception as e:
                print(f" ❌ Ошибка скачивания: {e}")
    
    print(f"📦 Всего скачано файлов: {len(downloaded_files)}")
    return downloaded_files
