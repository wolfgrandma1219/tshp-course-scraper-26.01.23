import json
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

def setup_driver():
    """設定 Chrome driver"""
    chrome_options = Options()
    chrome_options.add_argument('--headless=new')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36')
    
    # 直接使用系統的 ChromeDriver
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def convert_to_roc_date(date_str):
    """將西元年轉換為民國年 (YYYY-MM-DD -> YYY/MM/DD)"""
    try:
        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        roc_year = date_obj.year - 1911
        return f"{roc_year}/{date_obj.month:02d}/{date_obj.day:02d}"
    except:
        return date_str

def scrape_courses(start_date=None, end_date=None):
    """
    爬取台灣臨床藥學會的課程資訊
    start_date, end_date: 格式 'YYY/MM/DD' (民國年) 或 'YYYY-MM-DD' (西元年)
    """
    
    # 如果沒有提供日期，使用預設值
    if not start_date or not end_date:
        today = datetime.now()
        roc_year = today.year - 1911
        start_date = f"{roc_year}/{today.month:02d}/{today.day:02d}"
        
        future = today + timedelta(days=90)
        future_roc_year = future.year - 1911
        end_date = f"{future_roc_year}/{future.month:02d}/{future.day:02d}"
    
    # 如果是西元年格式，轉換為民國年
    if '-' in start_date:
        start_date = convert_to_roc_date(start_date)
    if '-' in end_date:
        end_date = convert_to_roc_date(end_date)
    
    print(f"查詢日期範圍: {start_date} ~ {end_date}")
    
    url = "https://www.tshp.org.tw/ehc-tshp/s/w/edu/teachMst/teachMstB2"
    driver = setup_driver()
    courses_data = []
    
    try:
        driver.get(url)
        wait = WebDriverWait(driver, 20)
        
        print("正在填寫查詢條件...")
        
        # 解析日期
        start_parts = start_date.split("/")
        end_parts = end_date.split("/")
        
        # 填寫開課日期起
        year_input = wait.until(EC.presence_of_element_located((By.ID, "courseStartYear")))
        year_input.clear()
        year_input.send_keys(start_parts[0])
        
        Select(driver.find_element(By.ID, "courseStartMonth")).select_by_value(start_parts[1])
        Select(driver.find_element(By.ID, "courseStartDay")).select_by_value(start_parts[2])
        
        # 填寫開課日期迄
        end_year_input = driver.find_element(By.ID, "showEndYear")
        end_year_input.clear()
        end_year_input.send_keys(end_parts[0])
        
        Select(driver.find_element(By.ID, "courseEndMonth")).select_by_value(end_parts[1])
        Select(driver.find_element(By.ID, "courseEndDay")).select_by_value(end_parts[2])
        
        # 勾選「開放報名」
        regist_checkbox = driver.find_element(By.ID, "registType1")
        if not regist_checkbox.is_selected():
            driver.execute_script("arguments[0].click();", regist_checkbox)
        print("✓ 已勾選「開放報名」")
        
        # 點擊查詢按鈕
        search_btn = driver.find_element(By.ID, "searchCourse")
        driver.execute_script("arguments[0].click();", search_btn)
        print("✓ 已點擊查詢按鈕")
        
        time.sleep(5)  # 等待結果載入
        
        # 解析結果頁面
        print("正在解析課程列表...")
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        rows = soup.find_all("tr", onclick=True)
        print(f"找到 {len(rows)} 個課程")
        
        # 提取課程詳細頁面連結
        course_links = []
        for row in rows:
            onclick = row.get("onclick", "")
            if "selectEduCourse" in onclick:
                course_id = onclick.split("'")[1]
                detail_url = f"https://www.tshp.org.tw/ehc-tshp/s/w/edu/scheduleInfo1/schedule/{course_id}"
                course_links.append(detail_url)
        
        # 去重
        course_links = list(set(course_links))
        print(f"共 {len(course_links)} 個不重複課程")
        
        # 訪問每個課程詳細頁面
        for idx, link in enumerate(course_links, 1):
            try:
                print(f"\n處理第 {idx}/{len(course_links)} 個課程...")
                driver.get(link)
                time.sleep(2)
                
                detail_soup = BeautifulSoup(driver.page_source, 'html.parser')
                
                # 提取課程資訊
                course_info = {
                    '課程主題': '',
                    '課程期間': '',
                    '積分': '',
                    '報名資訊': '',
                    '課程連結': link
                }
                
                # 使用更穩健的方式提取資訊
                text_content = detail_soup.get_text()
                lines = [line.strip() for line in text_content.split('\n') if line.strip()]
                
                for i, line in enumerate(lines):
                    # 提取標題
                    if ('標題' in line or '課程名稱' in line or '主題' in line) and i + 1 < len(lines):
                        next_line = lines[i + 1]
                        if '|' in line:
                            course_info['課程主題'] = line.split('|')[-1].strip()
                        elif next_line and len(next_line) > 3:
                            course_info['課程主題'] = next_line
                    
                    # 提取課程日期
                    elif ('課程日期' in line or '日期' in line) and i + 1 < len(lines):
                        next_line = lines[i + 1]
                        if '|' in line:
                            course_info['課程期間'] = line.split('|')[-1].strip()
                        elif next_line and ('/' in next_line or '~' in next_line or '-' in next_line):
                            course_info['課程期間'] = next_line
                    
                    # 提取積分
                    elif ('積分' in line or '學分' in line) and i + 1 < len(lines):
                        next_line = lines[i + 1]
                        if '|' in line:
                            course_info['積分'] = line.split('|')[-1].strip()
                        elif next_line and any(char.isdigit() for char in next_line):
                            course_info['積分'] = next_line
                    
                    # 提取報名資訊
                    elif '報名資訊' in line and i + 1 < len(lines):
                        next_line = lines[i + 1]
                        if '|' in line:
                            course_info['報名資訊'] = line.split('|')[-1].strip()
                        elif next_line and len(next_line) > 3:
                            course_info['報名資訊'] = next_line
                
                # 如果標題還是空的，嘗試從 h1, h2, h3 等標籤提取
                if not course_info['課程主題']:
                    title_tag = detail_soup.find(['h1', 'h2', 'h3', 'h4'])
                    if title_tag:
                        course_info['課程主題'] = title_tag.get_text(strip=True)
                
                # 只保存有標題的課程
                if course_info['課程主題']:
                    courses_data.append(course_info)
                    print(f"✓ {course_info['課程主題']}")
                else:
                    print(f"⚠ 跳過：無法提取課程標題")
                
            except Exception as e:
                print(f"✗ 處理課程時發生錯誤: {e}")
                continue
        
    except Exception as e:
        print(f"爬蟲執行錯誤: {e}")
        driver.save_screenshot('error_screenshot.png')
        print("已保存錯誤截圖: error_screenshot.png")
    
    finally:
        driver.quit()
    
    return courses_data

def save_to_json(data, filename='data.json'):
    """將資料存成 JSON 檔案"""
    output = {
        'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'total_courses': len(data),
        'courses': data
    }
    
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    
    print(f"\n{'='*50}")
    print(f"✓ 資料已儲存至 {filename}")
    print(f"✓ 共爬取 {len(data)} 筆課程資料")
    print(f"✓ 更新時間: {output['last_updated']}")
    print(f"{'='*50}")

def load_date_settings():
    """從設定檔讀取日期區間"""
    try:
        with open('date_settings.json', 'r', encoding='utf-8') as f:
            settings = json.load(f)
            start = settings.get('start_date', '')
            end = settings.get('end_date', '')
            
            # 如果是西元年格式，轉換為民國年
            if start and '-' in start:
                start = convert_to_roc_date(start)
            if end and '-' in end:
                end = convert_to_roc_date(end)
            
            return start, end
    except:
        # 預設值：今天到90天後（民國年格式）
        today = datetime.now()
        roc_year = today.year - 1911
        start = f"{roc_year}/{today.month:02d}/{today.day:02d}"
        
        future = today + timedelta(days=90)
        future_roc_year = future.year - 1911
        end = f"{future_roc_year}/{future.month:02d}/{future.day:02d}"
        
        return start, end

if __name__ == '__main__':
    print("="*50)
    print("台灣臨床藥學會課程爬蟲程式")
    print("="*50)
    
    # 讀取日期設定
    start_date, end_date = load_date_settings()
    print(f"日期設定: {start_date} ~ {end_date}")
    print("="*50 + "\n")
    
    # 執行爬蟲
    courses = scrape_courses(start_date, end_date)
    
    # 儲存結果
    if courses:
        save_to_json(courses)
    else:
        print("\n⚠ 警告：未抓取到任何課程資料")
        print("請檢查:")
        print("1. 網站是否正常運作")
        print("2. 日期範圍內是否有開放報名的課程")
        print("3. 查看 error_screenshot.png（如果有生成）")
        
        # 仍然保存空資料
        save_to_json(courses)
