import os # ساخت پوشه
import re # کار با Regex
import time # برای استفاده از زمان‌سنجی یا مکث
import datetime # کار با تاریخ 
import jdatetime # برای تبدیل تاریخ ها
import pandas as pd # خواندن و پردازش داده‌های جدولی 
import win32com.client as win32  # کار با برنامه‌ های ویندوز مثل اکسل 
import unicodedata # برای نرمال‌سازی و پاک‌سازی کاراکترهای یونیکد مثل فارسی
from selenium import webdriver # اجرای خودکار تعاملات با صفحات وب
from selenium.webdriver.chrome.service import Service # راه‌اندازی مرورگر Chrome با استفاده از
from selenium.webdriver.common.by import By  # یافتن عناصر صفحه
from selenium.webdriver.support.ui import WebDriverWait # اعمال زمان انتظار برای بارگذاری عناصر
from selenium.webdriver.support import expected_conditions as EC # برای بررسی وضعیت عناصر صفحه
from webdriver_manager.chrome import ChromeDriverManager # مدیریت درایور مرورگر 
from selenium.webdriver.common.keys import Keys # شبیه‌سازی فشار دادن کلیدهای صفحه‌کلید
from selenium.common.exceptions import StaleElementReferenceException  # مدیریت خطایی که وقتی عنصر صفحه دیگه معتبر نیست پیش میاد
from selenium.common.exceptions import TimeoutException # مدیریت خطایی که وقتی یک عملیات به‌موقع انجام نشه اتفاق می‌ افته

url = "https://tsetmc.com/"

# تنظیمات مرورگر
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")  # اجرای مرورگر بدون نمایش دادن 
chrome_options.add_argument("--disable-gpu")  # غیرفعال کردن پردازش گرافیکی (برای جلوگیری از برخی خطاها)
chrome_options.add_argument("--log-level=3")  # جلوگیری از نمایش پیام‌ های کم‌اهمیت‌ تر

# ساخت شیء درایور برای کنترل مرورگر
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

now = jdatetime.date.fromgregorian(date=datetime.date.today())
now_str = now.strftime(f"%Y-%m-%d")

def normalize_and_clean_filename(text):
    """ نرمال‌ سازی یونیکد, حذف نیم‌ فاصله‌ ها
        جایگزینی کاراکترهای غیرمجاز در نام فایل با "_"
    """
    if not text:
        return ""

    # نرمال‌ سازی یونیکد و حذف نیم‌ فاصله‌ ها
    text = unicodedata.normalize('NFC', text)
    text = text.replace("\u200c", "").replace("ی", "ي").replace("ک", "ك").strip()

    # جایگزینی کاراکترهای غیرمجاز در نام فایل با "_"
    return re.sub(r'[\\/*?:"<>|]', "_", text)

def create_folder(excel_file = "لیست_شرکت_ها.xlsx"):
    """ساخت پوشه‌ بر اساس لیست شرکت‌ها و صنایع از فایل اکسل"""
    companies = []         # لیست شرکت‌ها
    industries = set()     # مجموعه‌ای از صنایع 
    
    # باز کردن اکسل به صورت مخفی
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(os.path.abspath(excel_file))
    sheet = wb.Sheets(1) # انتخاب شیت اول
    row_count = sheet.UsedRange.Rows.Count # شمارش ردیف‌های استفاده‌شده

    # استخراج صنایع از فایل اکسل و ساخت پوشه برای هر صنعت
    for r in range(2, row_count + 1): # از ردیف دوم شروع می‌کنیم (چون ردیف اول عنوان‌ است)
        name = sheet.Cells(r, 2).Value # ستون دوم: نماد
        ind  = sheet.Cells(r, 6).Value # ستون ششم: صنعت
        if name and ind and name != "مقدار یافت نشد":
            industries.add(normalize_and_clean_filename(ind))
    
    # ساخت پوشه برای هر صنعت
    for ind in industries:
        os.makedirs(ind, exist_ok=True)

    # ساخت لیست شرکت‌ها + صنعت
    for r in range(2, row_count + 1):
        name = sheet.Cells(r, 2).Value
        ind  = sheet.Cells(r, 6).Value
        if name and ind and name != "مقدار یافت نشد":
            companies.append({
                "name": name,
                "industry": ind
            })
    # بستن فایل و خارج شدن از اکسل
    wb.Close(False)
    excel.Quit()

    print(f"تعداد شرکت‌های استخراج‌شده: {len(companies)}")
    return companies

def get_webpage(companies): 
    # باز کردن صفحه اصلی (TSETMC)
    driver.get(url) 

    # صبر تا زمانی که آیکون جست‌وجو در صفحه لود شود
    search_icon = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "a#search"))
    ) 
    # کلیک روی آیکون جست‌وجو
    search_icon.click() 

    for company in companies: 
        try: 
            # منتظر می‌ ماند تا فیلد ورودی جست‌وجو در دسترس باشد
            search = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='search']"))
            )

            # کلیک روی فیلد ورودی و پاک کردن محتوای قبلی
            search.click() 
            search.clear() 
            time.sleep(3) 
            
            # جست‌وجو کردن نماد
            search.send_keys(company["name"], Keys.ENTER)
            
            # منتظر می‌ ماند تا نماد در بخشی از صفحه ظاهر شود
            WebDriverWait(driver, 60).until(
                EC.text_to_be_present_in_element((By.CSS_SELECTOR, "div.box1.grey.tbl.z2_4"), company["name"])) 
            
            attempts = 3
            while attempts > 0:
                try:
                    # صبر می‌کنه تا تمام ردیف‌های شرکت‌هایی که منقضی نشده‌اند در جدول بارگذاری شوند
                    rows = WebDriverWait(driver, 60).until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.box1.grey.tbl.z2_4 div[role='row'] a:not(.expiredInstrument)")))
                    
                    for row in rows:
                        link = row.get_attribute("href") # گرفتن لینک مربوط به هر ردیف
                        normalized_text = normalize_and_clean_filename(row.text) # نرمال‌ سازی و پاک‌ سازی نام نمایش‌ داده‌ شده در ردیف
                        before_dash = normalized_text.split('-')[0].strip() # گرفتن بخش قبل از "-"
                        
                        # بررسی تطبیق نماد نوشته شده در اکسل و در سایت
                        if link and normalize_and_clean_filename(company["name"]).strip() == before_dash:
                            original_window = driver.current_window_handle # ذخیره تب فعلی مرورگر که در آن جست و جو انجام می شود
                            row.click()

                            is_new_tab = False
                            try:
                                # بررسی می‌کنه آیا تب جدید باز شده وقتی روی لینک مربوط به یک نماد کلیک کردیم
                                WebDriverWait(driver, 5).until(EC.number_of_windows_to_be(2))
                                # یک لیست ساخته می‌شه از تمام تب‌هایی که با تب اصلی فرق دارن 
                                new_window = [w for w in driver.window_handles if w != original_window][0]
                                driver.switch_to.window(new_window) # رفتن به تب جدید
                                is_new_tab = True

                            except TimeoutException:
                                # اگر تب جدید باز نشد همان تب فعلی باقی می‌ ماند
                                pass

                            try:
                                # صبر برای پیدا کردن تمام جدول‌ ها در صفحه
                                tables = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table")))
                            except TimeoutException:
                                print("جدولی یافت نشد")
                                tables = []

                            try:
                                # پیدا کردن اطلاعات مربوط به بخش مقایسه ی شرکت ها
                                xpath_0 = "//*[@id='Section_relco']/div[2]//div[@role='gridcell']"
                                elements_0 = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, xpath_0)))
                                
                            except TimeoutException:
                                print("بخش مقایسه ی شرکت ها یافت نشد")
                                elements_0 = []

                            try:
                                # پیدا کردن اطلاعات مربوط به بخش اطلاعیه
                                xpath = '//*[@id="Section_codal"]/div[2]//div[@role="gridcell"] | //*[@id="Section_codal"]/div[2]//span[@class="ag-header-cell-text"]'
                                elements = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
                                
                            except TimeoutException:
                                print("بخش اطلاعیه یافت نشد")
                                elements = []

                            try:
                                #  پیدا کردن اطلاعات مربوط به بخش سابقه معاملات
                                xpath_2 = "//*[@id='Section_history']/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]//div[@role='gridcell']"
                                elements_2 = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, xpath_2)))

                            except TimeoutException:
                                print("بخش سابقه معاملات یافت نشد")
                                elements_2 = []

                            try:
                                # پیدا کردن ستون مربوط به تاریخ بخش سابقه معاملات
                                elements_2_date = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.ag-pinned-right-cols-container a")))
                                
                            except TimeoutException:
                                print("ستون مربوط به تاریخ بخش سابقه معاملات یافت نشد")
                                elements_2_date = []

                            dfs = [] # لیستی برای دیتا‌فریم‌ های استخراج‌ شده

                            tds = driver.find_elements(By.TAG_NAME, "td")
                            # پیدا کردن مقدار قیمت پایانی به صورت جدا چون تگ tr برای آن وجود ندارد
                            extra_row = []
                            for i in range(len(tds) - 1):
                                label = tds[i].text.strip()
                                value = tds[i+1].text.strip()
                                
                                if label == "قیمت پایانی":
                                    extra_row = [label, value]
                                    break 

                            # علامتی برای اطمینان از اینکه فقط یک بار مقدار قیمت پایانی به لیست اضافه شه
                            extra_added = False
                            # گرفتن اطلاعات داخل جدول ها
                            for table in tables:
                                rows = []
                                for row in table.find_elements(By.CSS_SELECTOR, "tr"):
                                    cells = [cell.text.strip() for cell in row.find_elements(By.TAG_NAME, "th")]
                                    cells += [cell.text.strip() for cell in row.find_elements(By.TAG_NAME, "td")]

                                    if cells:
                                        if extra_row:
                                            rows.append(cells)
                                            if not extra_added: # اگر ردیف extra قبلاً اضافه نشده بود،  اضافه‌اش کن
                                                rows.append(extra_row)
                                                extra_added = True
                                        else:
                                            rows.append(cells)

                                df = pd.DataFrame(rows) # تبدیل جدول به DataFrame
                                dfs.append(df)
                            
                            # لیست کردن داده های بخش مقایسه ی شرکت ها
                            texts_0 = [
                                element.text.strip()
                                for element in elements_0
                                if element.text.strip() or any(c.isalnum() for c in element.text) # حداقل یک کاراکتر الفبایی یا عددی در متن وجود دارد یا بعد از حذف فاصله خالی نباشد
                            ]
                            # هر 8 آیتم بشه یک سطر
                            data_0 = []
                            for i in range(0, len(texts_0), 8):
                                group = texts_0[i:i+8]
                                if group not in data_0:
                                    data_0.append(group)

                            df_elements_0 = pd.DataFrame(data_0, columns=
                                                        ["نماد", "پایانی", " ", "آخرین", " ", "تعداد", "حجم", "ارزش"])
                            
                            # لیست کردن داده های بخش اطلاعیه
                            texts = [element.text for element in elements]
                            # هر 2 آیتم بشه یک سطر
                            data = []
                            for i in range(2, len(texts), 2):
                                first = texts[i]
                                second = texts[i + 1] if i + 1 < len(texts) else "" 
                                if [first, second] not in data:
                                    data.append([first, second])

                            df_elements = pd.DataFrame(data, columns=["تاریخ", "عنوان"])

                            # لیست کردن داده های بخش سابقه معاملات
                            texts_2 = [element.text for element in elements_2]
                            # لیست کردن تاریخ های مربوط به سابقه معاملات
                            texts_2_date = [element.text for element in elements_2_date]
                            
                            # هر 7 آیتم بشه یک سطر
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

                            industry = normalize_and_clean_filename(company["industry"])
                            name = normalize_and_clean_filename(company["name"])
                            
                            # مسیر فایل اکسل خروجی
                            file_path = os.path.join(industry, f"{name}_{now_str}.xlsx")
                            os.makedirs(os.path.dirname(file_path), exist_ok=True)

                            # ذخیره‌سازی همه دیتا‌فریم‌ها در فایل اکسل در یک شیت 
                            with pd.ExcelWriter(file_path) as writer:
                                start_row = 0 

                                for df in dfs:
                                    df.to_excel(writer, sheet_name="AllData", startrow=start_row, index=False, header=False)

                                    start_row += len(df) + 2 # فاصله بین جدول‌ ها

                                df_elements_0.to_excel(writer, sheet_name="AllData", startrow=start_row, index=False)

                                start_row += len(df_elements_0) + 2
                                
                                df_elements.to_excel(writer, sheet_name="AllData", startrow=start_row, index=False)

                                start_row += len(df_elements) + 2

                                df_elements_2.to_excel(writer, sheet_name="AllData", startrow=start_row, index=False)
                                print(f"{company['name']}")

                            if is_new_tab: # اگر تب جدید باز شده بود آن را ببند و برگرد به تب اصلی
                                driver.close()
                                driver.switch_to.window(original_window)
                            else:
                                driver.back()
                    break    
                        
                except StaleElementReferenceException: # ر عنصر صفحه معتبر نبود دوباره امتحان کند
                    attempts -= 1
                    time.sleep(1)  

            if attempts == 0:
                print(f"بازیابی ردیف‌ های {company['name']} ناموفق بود")

        except Exception as e:    
            print(f"خطا در جستجو {company['name']}: {e}")

companies = create_folder()
get_webpage(companies)
