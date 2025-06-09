import os
import re
import time
import datetime
import jdatetime
import pandas as pd
import win32com.client as win32
import unicodedata
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager

url = "https://tsetmc.com/"

chrome_options = webdriver.ChromeOptions()
# chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--log-level=3")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

now = jdatetime.date.fromgregorian(date=datetime.date.today())
now_str = now.strftime(f"%Y-%m-%d")

def clean_name(s):
    if not s:
        return ""
    s = str(s).replace("\u200c", "").strip()
    return re.sub(r'[\\/*?:"<>|]', "_", s)

def normalize_farsi(text):
    return unicodedata.normalize('NFC', text).replace("\u200c", "").replace("ی", "ي").replace("ک", "ك").strip()

def create_folder(excel_file = "لیست_شرکت_ها.xlsx"):
    companies = []
    industries = set()

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(os.path.abspath(excel_file))
    sheet = wb.Sheets(1)
    row_count = sheet.UsedRange.Rows.Count

 
    for r in range(2, row_count + 1):
        name = sheet.Cells(r, 2).Value
        ind  = sheet.Cells(r, 6).Value
        if name and ind and name != "مقدار یافت نشد":
            industries.add(clean_name(ind))

    for ind in industries:
        os.makedirs(ind, exist_ok=True)

  
    for r in range(2, row_count + 1):
        name = sheet.Cells(r, 2).Value
        ind  = sheet.Cells(r, 6).Value
        if name and ind and name != "مقدار یافت نشد":
            companies.append({
                "name": name,
                "industry": ind
            })

    wb.Close(False)
    excel.Quit()
    print(f"تعداد شرکت‌های استخراج‌شده: {len(companies)}")
    return companies

# def get_webpage(companies): 
#     driver.get(url) 
#     search_icon = WebDriverWait(driver, 10).until( EC.presence_of_element_located((By.CSS_SELECTOR, "a#search")) ) 
    
#     search_icon.click() 
#     search = WebDriverWait(driver, 10).until( EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='search']")) ) 
#     for company in companies[:4]: 
#         try: 
#             search.click() 
#             search.clear() 
#             search.send_keys(company["name"], Keys.ENTER)

#             WebDriverWait(driver, 10).until( EC.text_to_be_present_in_element( (By.CSS_SELECTOR, "div.box1.grey.tbl.z2_4"), company["name"] ) ) 

#             rows = WebDriverWait(driver, 10).until( EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.box1.grey.tbl.z2_4 div[role='row'] a"))) 

#             for row in rows: 
#                 link = row.get_attribute("href") 
#                 normalized_text = normalize_farsi(row.text) 
#                 before_dash = normalized_text.split('-')[0].strip() 
#                 print(company["name"], before_dash)
                    
#                 # if link and normalize_farsi(company["name"]).strip() == before_dash \
#                 # and not row.find_elements(By.XPATH, ".//span[contains(text(),'حذف')]"):
                
#                 #     # ذخیره شناسه پنجره اصلی
#                 #     original_window = driver.current_window_handle
                    
#                 #     # کلیک روی لینک (که پنجره جدید باز می‌کند)
#                 #     row.click()
                    
#                 #     # انتظار برای اینکه تعداد پنجره‌ها به 2 برسد
#                 #     WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
                    
#                 #     # سوئیچ به پنجره جدید
#                 #     new_window = [window for window in driver.window_handles if window != original_window][0]
#                 #     driver.switch_to.window(new_window)
                    
#                 #     # انجام پردازش مورد نظر (برای مثال، استخراج اطلاعات)
#                 #     # ...
                    
#                 #     # بستن پنجره جدید
#                 #     driver.close()
                    
#                 #     # انتظار تا فقط یک پنجره باقی بماند
#                 #     WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(1))
                    
#                 #     # سوئیچ مجدد به پنجره اصلی
#                 #     driver.switch_to.window(original_window)
                    
#                 #     # (در صورت لزوم، ادامه کار برای شرکت فعلی یا پایان حلقه)
#                 #     break
        
#         except StaleElementReferenceException as e:
#             print(e)


def get_webpage(companies): 
    driver.get(url) 
    search_icon = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "a#search"))
    ) 
    search_icon.click() 

    for company in companies[11:14]: 
        try: 
            search = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='search']"))
            )
            search.click() 
            search.clear() 
            time.sleep(5) 

            search.send_keys(company["name"], Keys.ENTER)
            time.sleep(5)

            WebDriverWait(driver, 30).until(
                EC.text_to_be_present_in_element(
                    (By.CSS_SELECTOR, "div.box1.grey.tbl.z2_4"), company["name"]
                )
            ) 
            
            attempts = 3
            while attempts > 0:
                try:
                    rows = WebDriverWait(driver, 30).until(
                        EC.presence_of_all_elements_located(
                            (By.CSS_SELECTOR, "div.box1.grey.tbl.z2_4 div[role='row'] a:not(.expiredInstrument)")
                        )
                    )
                    
                    for row in rows: 
                        link = row.get_attribute("href") 
                        normalized_text = normalize_farsi(row.text) 
                        before_dash = normalized_text.split('-')[0].strip()

                        if link and normalize_farsi(company["name"]).strip() == before_dash:
                            original_window = driver.current_window_handle
                        
                            row.click()
                            
                            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
                            
                            new_window = [window for window in driver.window_handles if window != original_window][0]
                            driver.switch_to.window(new_window)
                            
                            
                            
                            driver.close()
                            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(1))
                            driver.switch_to.window(original_window)
                    break  
                except StaleElementReferenceException:
                    attempts -= 1
                    time.sleep(1)  

            if attempts == 0:
                print(f"Failed to retrieve rows for {company['name']} due to stale elements.")

        except Exception as e:
            print(f"Error searching for {company['name']}: {e}")

companies = create_folder()
webpage = get_webpage(companies)
