# Requirements:
# pip install selenium webdriver-manager pandas openpyxl

import time
import pandas as pd
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    TimeoutException,
    StaleElementReferenceException,
    NoSuchElementException,
    WebDriverException,
)


INPUT_XLSX = "az_app_awarded_achieved1.xlsx"   
OUTPUT_XLSX = "az_app_awarded_results.xlsx"    
FORM_URL = "https://doa.az.gov/public-information-and-records-request-form"


FIRST_NAME = "Raaj"
LAST_NAME = "-Thipparthy"
EMAIL = "raajnrao@gmail.com"
PHONE = "8325197135"


REQUEST_TYPE_VALUE = "416"


MAX_CAPTCHA_TRIES = 5


def build_summary(label, agency, code):
    return (
        f"I am requesting a copy of the winning and shortlisted proposals for {label} by {agency}."
        f"I am requesting a copy of the winning and shortlisted proposals for the referenced award. "
        f"The solicitation/contract number is {code}."
    )

def start_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--window-size=1280,900")
    
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.set_page_load_timeout(60)
    return driver

def wait_present(driver, by, value, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))

def goto_form(driver, url, tries=3):
    last_err = None
    for i in range(1, tries + 1):
        try:
            print(f"[NAV] Opening form (attempt {i}) -> {url}")
            driver.get(url)
            WebDriverWait(driver, 30).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "edit-first-name")))
            print(f"[NAV] Landed on: {driver.current_url}")
            return True
        except (TimeoutException, WebDriverException) as e:
            last_err = e
            print(f"[NAV] Retry due to: {e}")
            time.sleep(2)
    print(f"[NAV] Failed to open form after retries: {last_err}")
    return False

def safe_click_submit(driver, timeout=20):
    btn = wait_present(driver, By.ID, "edit-submit", timeout)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    driver.execute_script("window.scrollBy(0, -160);")
    try:
        WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "edit-submit")))
        btn.click()
        return
    except ElementClickInterceptedException:
        try:
            driver.execute_script("arguments[0].click();", btn)
            return
        except Exception:
            pass
        try:
            ActionChains(driver).move_to_element_with_offset(btn, 1, 1).click().perform()
            return
        except Exception:
            time.sleep(1)
            driver.execute_script("arguments[0].click();", btn)

def safe_type(driver, locator, text, timeout=30):
    
    ignored = (StaleElementReferenceException, NoSuchElementException)
    el = WebDriverWait(driver, timeout, ignored_exceptions=ignored).until(
        EC.presence_of_element_located(locator)
    )
    for _ in range(2):
        try:
            el.clear()
            el.send_keys(text)
            return
        except (StaleElementReferenceException, NoSuchElementException):
            el = WebDriverWait(driver, timeout, ignored_exceptions=ignored).until(
                EC.presence_of_element_located(locator)
            )

def fill_form_fields(driver, row):
    wait_present(driver, By.ID, "edit-first-name", 30)
    safe_type(driver, (By.ID, "edit-first-name"), FIRST_NAME, 30)
    safe_type(driver, (By.ID, "edit-last-name"), LAST_NAME, 30)
    safe_type(driver, (By.ID, "edit-email-mail-1"), EMAIL, 30)
    safe_type(driver, (By.ID, "edit-email-mail-2"), EMAIL, 30)
    safe_type(driver, (By.ID, "edit-phone-number"), PHONE, 30)

    
    Select(WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "edit-request-type"))
    )).select_by_value(REQUEST_TYPE_VALUE)

    label = str(row.get("Label", "") or "")
    agency = str(row.get("Agency", "") or "")
    code = str(row.get("Code", "") or "")
    summary_text = build_summary(label, agency, code)
    safe_type(driver, (By.ID, "edit-description"), summary_text, 30)

def wait_for_result_or_error(driver, timeout=25):
    try:
        WebDriverWait(driver, timeout).until(
            EC.any_of(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".webform-confirmation, .messages--status")),
                EC.presence_of_element_located((By.CSS_SELECTOR, ".messages--error, .error"))
            )
        )
    except TimeoutException:
        return ("timeout", "")

    body_text = driver.find_element(By.TAG_NAME, "body").text.lower()
    has_success = ("thank" in body_text) or ("received" in body_text)
    has_error = len(driver.find_elements(By.CSS_SELECTOR, ".messages--error, .error")) > 0
    captcha_error = "captcha" in body_text

    if has_success and not has_error:
        return ("success", "")
    if captcha_error:
        return ("captcha", "CAPTCHA incorrect")
    if has_error:
        return ("error", "Form validation error")
    return ("unknown", "")

def submit_one_row(driver, row, row_num):
    if not goto_form(driver, FORM_URL, tries=3):
        return "fail"

    fill_form_fields(driver, row)

    tries = 0
    while tries < MAX_CAPTCHA_TRIES:
        tries += 1
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        print(f"Row {row_num} - Attempt {tries}: Solve the CAPTCHA in the browser window, then press Enter here...")
        input("Press Enter to click Submit...")

        
        try:
            safe_click_submit(driver, timeout=20)
        except (ElementClickInterceptedException, StaleElementReferenceException):
            time.sleep(1)
            safe_click_submit(driver, timeout=20)

        status, _ = wait_for_result_or_error(driver, timeout=30)
        if status == "success":
            return "success"
        if status == "captcha":
            print("CAPTCHA incorrect detected. Reloading form and retrying...")
            
            if not goto_form(driver, FORM_URL, tries=3):
                return "fail"
            fill_form_fields(driver, row)
            continue
        if status in ("error", "timeout", "unknown"):
            print(f"Submission did not succeed (status={status}).")
            return "fail"

    return "fail"

def main():
    df = pd.read_excel(INPUT_XLSX)
    print(f"[DATA] Loaded rows: {df.shape[0]}")

    expected_cols = [
        "Label", "Commodity", "Agency", "Status",
        "RFx Awarded", "Begin (UTC-7)", "End (UTC-7)", "Code"
    ]
    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""

    driver = start_driver()
    results = []
    try:
        for idx, row in df.iterrows():
            result_flag = submit_one_row(driver, row, idx + 1)
            results.append({
                "Label": row.get("Label", ""),
                "Commodity": row.get("Commodity", ""),
                "Agency": row.get("Agency", ""),
                "Status": row.get("Status", ""),
                "RFx Awarded": row.get("RFx Awarded", ""),
                "Begin (UTC-7)": row.get("Begin (UTC-7)", ""),
                "End (UTC-7)": row.get("End (UTC-7)", ""),
                "success": result_flag
            })
            time.sleep(1)
    finally:
        driver.quit()

    out_df = pd.DataFrame(results)[
        ["Label", "Commodity", "Agency", "Status", "RFx Awarded", "Begin (UTC-7)", "End (UTC-7)", "success"]
    ]
    out_df.to_excel(OUTPUT_XLSX, index=False)
    print(f"Saved results to {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
