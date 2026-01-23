import json
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def scrape_courses(start_date, end_date):
    """
    爬取台灣外科醫學會的課程資訊
    start_date, end_date: 格式 'YYYY/MM/DD'
    """
    # 設定 Chrome options
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # 背景執行
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    
    # 使用 webdriver-manager 自動管理 ChromeDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    courses_data = []
    
    try:
        # 1. 開啟目標網站
        driver.get('https://www.tshp.org.tw/ehc-tshp/s/w/edu/teachMst/teachMstB2')
        time.sleep(3)
        
        # 2. 填入日期區間
        start_date_input = driver.find_element(By.ID, 'g6')  # 開課日期起
        end_date_input = driver.find_element(By.ID, 'g7')    # 開課日期迄
        
        start_date_input.clear()
        start_date_input.send_keys(start_date)
        
        end_date_input.clear()
        end_date_input.send_keys(end_date)
        
        # 3. 勾選"開放報名"
        open_registration_checkbox = driver.find_element(By.CSS_SELECTOR, 'input[value="1"]')  # 報名資訊：開放報名
        if not open_registration_checkbox.is_selected():
            open_registration_checkbox.click()
        
        # 4. 點選查詢按鈕
        search_button = driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]')
        search_button.click()
        
        # 等待結果載入
        time.sleep(3)
        
        # 5. 抓取查詢結果
        rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')
        
        for row in rows:
            try:
                cols = row.find_elements(By.TAG_NAME, 'td')
                if len(cols) >= 5:
                    course_period = cols[0].text.strip()  # 課程期間
                    course_title_element = cols[2].find_element(By.TAG_NAME, 'a')
                    course_title = course_title_element.text.strip()  # 課程主題
                    credits = cols[4].text.strip()  # 積分
                    
                    # 點進課程主題連結，抓取報名資訊
                    course_link = course_title_element.get_attribute('href')
                    
                    # 開新分頁
                    driver.execute_script("window.open('');")
                    driver.switch_to.window(driver.window_handles[1])
                    driver.get(course_link)
                    time.sleep(2)
                    
                    # 抓取報名資訊
                    try:
                        registration_info = driver.find_element(By.XPATH, "//th[contains(text(), '報名資訊')]/following-sibling::td").text.strip()
                    except:
                        registration_info = "無法取得"
                    
                    # 關閉分頁，回到主頁
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    
                    courses_data.append({
                        '課程期間': course_period,
                        '課程主題': course_title,
                        '積分': credits,
                        '報名資訊': registration_info,
                        '課程連結': course_link
                    })
                    
            except Exception as e:
                print(f"處理課程時發生錯誤: {e}")
                continue
        
    except Exception as e:
        print(f"爬蟲執行錯誤: {e}")
    
    finally:
        driver.quit()
    
    return courses_data

def save_to_json(data, filename='data.json'):
    """將資料存成 JSON 檔案"""
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump({
            'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'courses': data
        }, f, ensure_ascii=False, indent=2)
    print(f"資料已儲存至 {filename}")

def load_date_settings():
    """從設定檔讀取日期區間"""
    try:
        with open('date_settings.json', 'r', encoding='utf-8') as f:
            settings = json.load(f)
            return settings.get('start_date'), settings.get('end_date')
    except:
        # 預設值：今天到30天後
        from datetime import timedelta
        today = datetime.now()
        start = today.strftime('%Y/%m/%d')
        end = (today + timedelta(days=30)).strftime('%Y/%m/%d')
        return start, end

if __name__ == '__main__':
    # 讀取日期設定
    start_date, end_date = load_date_settings()
    print(f"開始爬取課程資料... 日期區間: {start_date} ~ {end_date}")
    
    # 執行爬蟲
    courses = scrape_courses(start_date, end_date)
    
    # 儲存結果
    save_to_json(courses)
    print(f"共爬取 {len(courses)} 筆課程資料")
