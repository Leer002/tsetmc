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
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

url = "https://tsetmc.com/"

chrome_options = webdriver.ChromeOptions()
# chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--log-level=3")

driver = webdriver.Chrome(service=Service(ChromeDriverManager(driver_version="137.0.7151.104").install()), options=chrome_options)

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

def get_webpage(companies): 
    driver.get(url) 
    search_icon = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "a#search"))
    ) 
    search_icon.click() 

    for company in companies[76:]: 
        try: 
            search = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='search']"))
            )
            search.click() 
            search.clear() 
            time.sleep(5) 

            search.send_keys(company["name"], Keys.ENTER)
            time.sleep(5)

            WebDriverWait(driver, 60).until(
                EC.text_to_be_present_in_element(
                    (By.CSS_SELECTOR, "div.box1.grey.tbl.z2_4"), company["name"]
                )
            ) 
            
            attempts = 3
            while attempts > 0:
                try:
                    rows = WebDriverWait(driver, 60).until(
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

                            is_new_tab = False
                            try:
                                WebDriverWait(driver, 5).until(EC.number_of_windows_to_be(2))
                                new_window = [w for w in driver.window_handles if w != original_window][0]
                                driver.switch_to.window(new_window)
                                is_new_tab = True
                            except TimeoutException:
                                pass
                            try:
                                tables = WebDriverWait(driver, 60).until(
                                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table"))
                                )
                            except TimeoutException:
                                print("عنصر tables یافت نشد، لیست خالی ایجاد می‌شود.")
                                tables = []

                            try:
                                xpath_0 = "//*[@id='Section_relco']/div[2]//div[@role='gridcell']"
                                elements_0 = WebDriverWait(driver, 60).until(
                                    EC.presence_of_all_elements_located((By.XPATH, xpath_0))
                                )
                            except TimeoutException:
                                print("عنصر elements_0 یافت نشد، لیست خالی ایجاد می‌شود.")
                                elements_0 = []

                            try:
                                xpath = '//*[@id="Section_codal"]/div[2]//div[@role="gridcell"] | //*[@id="Section_codal"]/div[2]//span[@class="ag-header-cell-text"]'
                                elements = WebDriverWait(driver, 60).until(
                                    EC.presence_of_all_elements_located((By.XPATH, xpath))
                                )
                            except TimeoutException:
                                print("عنصر elements یافت نشد، لیست خالی ایجاد می‌شود.")
                                elements = []

                            try:
                                xpath_2 = "//*[@id='Section_history']/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]//div[@role='gridcell']"
                                elements_2 = WebDriverWait(driver, 60).until(
                                    EC.presence_of_all_elements_located((By.XPATH, xpath_2))
                                )
                            except TimeoutException:
                                print("عنصر elements_2 یافت نشد، لیست خالی ایجاد می‌شود.")
                                elements_2 = []

                            try:
                                elements_2_date = WebDriverWait(driver, 60).until(
                                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.ag-pinned-right-cols-container a"))
                                )
                            except TimeoutException:
                                print("عنصر elements_2_date یافت نشد، لیست خالی ایجاد می‌شود.")
                                elements_2_date = []

                            dfs = []
                            tds = driver.find_elements(By.TAG_NAME, "td")

                            extra_row = []
                            for i in range(len(tds) - 1):
                                label = tds[i].text.strip()
                                value = tds[i+1].text.strip()
                                
                                if label == "قیمت پایانی":
                                    extra_row = [label, value]
                                    break 
                            extra_added = False
                            for table in tables:
                                rows = []
                                for row in table.find_elements(By.CSS_SELECTOR, "tr"):
                                    cells = [cell.text.strip() for cell in row.find_elements(By.TAG_NAME, "th")]
                                    cells += [cell.text.strip() for cell in row.find_elements(By.TAG_NAME, "td")]

                                    if cells:
                                        if extra_row:
                                            rows.append(cells)
                                            if extra_row and not extra_added:
                                                rows.append(extra_row)
                                                extra_added = True
                                        else:
                                            rows.append(cells)

                                df = pd.DataFrame(rows)
                                dfs.append(df)
                            
                            texts_0 = [
                                element.text.strip()
                                for element in elements_0
                                if element.text.strip() or any(c.isalnum() for c in element.text)
                            ]
                            texts = [element.text for element in elements]

                            texts_2 = [element.text for element in elements_2]
                            texts_2_date = [element.text for element in elements_2_date]

                            data_0 = []
                            for i in range(0, len(texts_0), 8):
                                group = texts_0[i:i+8]
                                if group not in data_0:
                                    data_0.append(group)

                            df_elements_0 = pd.DataFrame(data_0, columns=
                                                        ["نماد", "پایانی", " ", "آخرین", " ", "تعداد", "حجم", "ارزش"])

                            data = []
                            for i in range(2, len(texts), 2):
                                first = texts[i]
                                second = texts[i + 1] if i + 1 < len(texts) else "" 
                                if [first, second] not in data:
                                    data.append([first, second])

                            df_elements = pd.DataFrame(data, columns=["تاریخ", "عنوان"])

                            data_2 = []
                            j = 0
                            for i in range(0, len(texts_2), 7):
                                items = texts_2[i:i + 7]
                                if len(items) < 7:
                                    items += [""] * (7 - len(items))
                                
                                date = texts_2_date[j] if j < len(texts_2_date) else ""
                                j += 1
                                
                                group_2 = [date] + items
                                data_2.append(group_2)


                            df_elements_2 = pd.DataFrame(data_2, columns=["تاریخ", "پایانی", "تغییر%", "کمترین", "بیشترین", "تعداد", "حجم", "ارزش"])

                            industry = clean_name(company["industry"])
                            name = clean_name(company["name"])

                            file_path = os.path.join(industry, f"{name}_{now_str}.xlsx")
                            os.makedirs(os.path.dirname(file_path), exist_ok=True)

                            with pd.ExcelWriter(file_path) as writer:
                                start_row = 0 
                                for idx, df in enumerate(dfs, start=1):
                                    df.to_excel(writer, sheet_name="AllData", startrow=start_row, index=False, header=False)

                                    start_row += len(df) + 2

                                df_elements_0.to_excel(writer, sheet_name="AllData", startrow=start_row, index=False)

                                start_row += len(df_elements_0) + 2
                                
                                df_elements.to_excel(writer, sheet_name="AllData", startrow=start_row, index=False)

                                start_row += len(df_elements) + 2

                                df_elements_2.to_excel(writer, sheet_name="AllData", startrow=start_row, index=False)
                                print("ja")

                            if is_new_tab:
                                driver.close()
                                driver.switch_to.window(original_window)
                            else:
                                driver.back()
                    break    
                        
                except StaleElementReferenceException:
                    attempts -= 1
                    time.sleep(1)  

            if attempts == 0:
                print(f"Failed to retrieve rows for {company['name']} due to stale elements.")

        except Exception as e:    
            print(f"Error searching for {company['name']}: {e}")

companies = create_folder()
get_webpage(companies)
