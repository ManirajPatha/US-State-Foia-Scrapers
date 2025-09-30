# Requirements:
# pip install selenium webdriver-manager pandas openpyxl

import time
import pandas as pd
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    NoSuchElementException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
    WebDriverException,
)


INPUT_XLSX = "connecticut_awarded_bids.xlsx"   
OUTPUT_XLSX = "ct_govqa_results.xlsx"          

LOGIN_URL = ("https://dasct.govqa.us/WEBAPP/_rs/"
             "(S(fjdcph5ifhqhheaao1wovxjm))/RequestLogin.aspx"
             "?sSessionID=&rqst=1&target=YpURA3m6cNU+N1K9kEqQhqz8yC2ZLKNdSdB4wnowVJ5S8CGTBp2GIItHg4/"
             "I0pUM8Jvp1AAd4YheCcTrA795fG9P3xL5LmB/wFQjiIoSWN7tLnJa+Bm/oEirHbO2IQAI")


ID_EMAIL = "RequesLoginFormLayout_txtUsername_I"
ID_PASSWORD = "RequesLoginFormLayout_txtPassword_I"
ID_LOGIN_SUBMIT = "RequesLoginFormLayout_btnLogin_I"

ID_TEXTAREA_REQ = "requestData_CustomFieldsFormLayout_cf_DeflectionTextContainer_2_I"
ID_DATE_FROM = "requestData_CustomFieldsFormLayout_cf_56_I"
ID_DATE_TO = "requestData_CustomFieldsFormLayout_cf_57_I"
ID_KEYWORD1 = "requestData_CustomFieldsFormLayout_cf_59_I"
ID_FORM_SUBMIT = "btnSaveData_I"


PORTAL_EMAIL = "akchrao@gmail.com"
PORTAL_PASSWORD = "Password@1234"


RUN_LIMIT = None  # e.g., 10


def start_driver():
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--window-size=1280,900")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=opts)
    driver.set_page_load_timeout(60)
    return driver

def wait_ready(driver, timeout=30):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

def js_set_with_events(driver, element, value):
    driver.execute_script("""
        const el = arguments[0], val = arguments[1];
        el.value = val;
        el.dispatchEvent(new Event('input', {bubbles:true}));
        el.dispatchEvent(new Event('change', {bubbles:true}));
    """, element, value)

def safe_fill_input(driver, by, value, timeout=20):
    
    el = WebDriverWait(driver, timeout).until(EC.presence_of_element_located(by))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try:
        el.clear()
    except Exception:
        pass
    try:
        el.send_keys(value)
    except (ElementNotInteractableException, StaleElementReferenceException, NoSuchElementException, Exception):
        el = WebDriverWait(driver, timeout).until(EC.presence_of_element_located(by))
        js_set_with_events(driver, el, value)

def mmddyyyy(x):
    if pd.isna(x) or str(x).strip() == "":
        return ""
    try:
        
        if isinstance(x, str):
            x = x.strip()
            try:
                dt = datetime.strptime(x, "%Y-%m-%d")
            except ValueError:
                
                for fmt in ("%m/%d/%Y", "%Y/%m/%d", "%m-%d-%Y"):
                    try:
                        dt = datetime.strptime(x, fmt)
                        break
                    except ValueError:
                        dt = None
                if dt is None:
                    
                    dt = pd.to_datetime(x, errors="coerce")
                    if pd.isna(dt):
                        return ""
        else:
            dt = pd.to_datetime(x, errors="coerce")
            if pd.isna(dt):
                return ""
        if isinstance(dt, pd.Timestamp):
            dt = dt.to_pydatetime()
        return dt.strftime("%m/%d/%Y")
    except Exception:
        return ""

def goto_login(driver):
    for i in range(1, 4):
        try:
            print(f"[NAV] Open login attempt {i}")
            driver.get(LOGIN_URL)
            wait_ready(driver, 30)
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, ID_EMAIL)))
            return True
        except (TimeoutException, WebDriverException):
            time.sleep(2)
    return False

def login(driver):
    if not goto_login(driver):
        return False
    safe_fill_input(driver, (By.ID, ID_EMAIL), PORTAL_EMAIL, 30)
    safe_fill_input(driver, (By.ID, ID_PASSWORD), PORTAL_PASSWORD, 30)
    
    try:
        btn = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, ID_LOGIN_SUBMIT)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        try:
            btn.click()
        except ElementClickInterceptedException:
            driver.execute_script("arguments[0].click();", btn)
    except Exception:
        pass
    
    try:
        WebDriverWait(driver, 30).until(
            EC.any_of(
                EC.presence_of_element_located((By.ID, ID_TEXTAREA_REQ)),
                EC.url_contains("RequestOpen.aspx")
            )
        )
        return True
    except TimeoutException:
        return False

def fill_and_submit_request(driver, row, row_idx):
    
    notice_id = str(row.get("notice_id", "") or "")
    title = str(row.get("title", "") or "")
    source = row.get("source", "")
    publish_date = row.get("publish_date", "")
    closing_date = row.get("closing_date", "")
    issuer = row.get("issuer", "")
    email = row.get("email", "")
    industry = str(row.get("industry", "") or "")

    request_text = (
        f"I am requesting a copy of the winning and shortlisted proposals for {title} ."
        f"I am requesting a copy of the winning and shortlisted proposals for the referenced award. "
        f"The solicitation/contract number is {notice_id}."
    )

    
    safe_fill_input(driver, (By.ID, ID_TEXTAREA_REQ), request_text, 30)

    
    safe_fill_input(driver, (By.ID, ID_DATE_FROM), mmddyyyy(publish_date), 20)
    safe_fill_input(driver, (By.ID, ID_DATE_TO), mmddyyyy(closing_date), 20)

   
    safe_fill_input(driver, (By.ID, ID_KEYWORD1), industry, 20)

    
    print(f"Row {row_idx}: Solve CAPTCHA in the browser window, then press Enter here...")
    input("Press Enter to submit...")

    
    try:
        btn = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, ID_FORM_SUBMIT)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        try:
            btn.click()
        except ElementClickInterceptedException:
            driver.execute_script("arguments[0].click();", btn)
    except Exception:
        pass

    
    try:
        WebDriverWait(driver, 30).until(
            EC.any_of(
                EC.text_to_be_present_in_element((By.TAG_NAME, "body"), "Thank"),
                EC.text_to_be_present_in_element((By.TAG_NAME, "body"), "received"),
                EC.presence_of_element_located((By.CSS_SELECTOR, ".dxpc-content, .alert-success, .validationSummary"))
            )
        )
        body_text = driver.find_element(By.TAG_NAME, "body").text.lower()
        ok = ("thank" in body_text) or ("received" in body_text) or ("request number" in body_text)
        success = "success" if ok else "fail"
    except TimeoutException:
        success = "fail"

    return {
        "notice_id": notice_id,
        "title": title,
        "source": source,
        "publish_date": publish_date,
        "closing_date": closing_date,
        "issuer": issuer,
        "email": email,
        "industry": industry,
        "success": success,
    }


def main():
    df = pd.read_excel(INPUT_XLSX)

    
    if RUN_LIMIT:
        df = df.head(RUN_LIMIT)

    
    expected_cols = [
        "notice_id", "title", "source", "publish_date", "closing_date",
        "issuer", "email", "industry"
    ]
    for c in expected_cols:
        if c not in df.columns:
            df[c] = ""

    driver = start_driver()
    results = []
    try:
        if not login(driver):
            print("Login failed; aborting this run.")
            
            for idx, row in df.iterrows():
                out = {
                    "notice_id": row.get("notice_id", ""),
                    "title": row.get("title", ""),
                    "source": row.get("source", ""),
                    "publish_date": row.get("publish_date", ""),
                    "closing_date": row.get("closing_date", ""),
                    "issuer": row.get("issuer", ""),
                    "email": row.get("email", ""),
                    "industry": row.get("industry", ""),
                    "success": "fail",
                }
                results.append(out)
        else:
            for idx, row in df.iterrows():
                print(f"=== Processing row {idx+1} of {len(df)} ===")
                result = fill_and_submit_request(driver, row, idx + 1)
                results.append(result)
                time.sleep(1)
    finally:
        driver.quit()

    out_df = pd.DataFrame(results)[
        ["notice_id", "title", "source", "publish_date", "closing_date", "issuer", "email", "industry", "success"]
    ]
    out_df.to_excel(OUTPUT_XLSX, index=False)
    print(f"Saved results to {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
