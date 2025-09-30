# Requirements:
# pip install selenium webdriver-manager pandas openpyxl

import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    NoSuchElementException,
    ElementClickInterceptedException,
    WebDriverException,
    ElementNotInteractableException,
)


INPUT_XLSX = "pa_archived_closed_bids.xlsx"          
OUTPUT_XLSX = "pa_rtk_results.xlsx"                  
FORM_URL = "https://www.openrecords.pa.gov/RTKL/RequestForm.cfm"  

RUN_LIMIT = None  

SALUTATION = "Mr."
FIRST_NAME = "Raaj"
LAST_NAME = "Thipparthy"
ADDRESS1 = "8181 Fannin St"
CITY = "Houston"
ZIP = "77054"
EMAIL = "raajnrao@gmail.com"


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

def goto_form(driver, tries=3):
    last_err = None
    for i in range(1, tries + 1):
        try:
            print(f"[NAV] Opening ({i}) -> {FORM_URL}")
            driver.get(FORM_URL)
            wait_ready(driver, 30)
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "nameFirst")))
            print(f"[NAV] Landed: {driver.current_url}")
            return True
        except (TimeoutException, WebDriverException) as e:
            last_err = e
            print(f"[NAV] Retry: {e}")
            time.sleep(2)
    print(f"[NAV] Failed to open form: {last_err}")
    return False

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

def safe_set_textarea(driver, locator, text, timeout=30):
    
    el = WebDriverWait(driver, timeout).until(EC.presence_of_element_located(locator))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try:
        el.clear()
        el.send_keys(text)
    except (ElementNotInteractableException, StaleElementReferenceException, NoSuchElementException, Exception):
        try:
            el = driver.find_element(*locator)
            driver.execute_script("""
                const el = arguments[0];
                el.value = arguments[1];
                el.dispatchEvent(new Event('input', {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
            """, el, text)
        except Exception:
           
            el = WebDriverWait(driver, timeout).until(EC.presence_of_element_located(locator))
            driver.execute_script("arguments[0].value = arguments[1];", el, text)

def select_by_visible_text_even_if_select2(driver, select_id, visible_text, timeout=20):
    sel = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, select_id)))
    Select(sel).select_by_visible_text(visible_text)

def select2_search_and_pick(driver, container_id_suffix, term, results_ul_id, timeout=20):
    
    container = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.ID, f"select2-{container_id_suffix}-container"))
    )
    container.click()
    search = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field"))
    )
    search.clear()
    search.send_keys(term)
    match_xpath = f"//ul[@id='{results_ul_id}']//li[contains(normalize-space(.), {repr(term)})]"
    result = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, match_xpath))
    )
    result.click()

def compose_request_text(title, agency, bid_no):
    return (
        f"I am requesting a copy of the winning and shortlisted proposals for {title} by {agency}."
        f"I am requesting a copy of the winning and shortlisted proposals for the referenced award. "
        f"The solicitation/contract number is {bid_no}."
    )

def wait_submit_enabled(driver, timeout=60):
    def enabled(d):
        btn = d.find_element(By.ID, "btn-submit")
        dis = btn.get_attribute("disabled")
        return btn.is_enabled() and (dis is None or dis is False or dis == "false")
    WebDriverWait(driver, timeout).until(enabled)

def detect_submission_status(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            EC.any_of(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".alert-success, .messages--status, .webform-confirmation")),
                EC.presence_of_element_located((By.CSS_SELECTOR, ".alert-danger, .messages--error, .error"))
            )
        )
    except TimeoutException:
        pass
    body = driver.find_element(By.TAG_NAME, "body").text.lower()
    if "thank" in body or "received" in body or "success" in body:
        return "success"
    if "error" in body or "invalid" in body or "failed" in body:
        return "fail"
    return "success" if "RequestForm.cfm" not in driver.current_url else "fail"


def submit_one_row(driver, row, row_num):
    
    bid_no = str(row.get("Bid No", "") or "")
    bid_type = str(row.get("Bid Type", "") or "")
    title = str(row.get("Title", "") or "")
    description = str(row.get("Description", "") or "")
    agency = str(row.get("Agency", "") or "")
    county = str(row.get("County", "") or "")
    bid_start = row.get("Bid Start Date", "")
    bid_end = row.get("Bid End Date", "")
    bid_open = row.get("Bid Open Date", "")
    status_txt = str(row.get("Status", "") or "")

    if not goto_form(driver, tries=3):
        return {
            "Bid No": bid_no, "Bid Type": bid_type, "Title": title, "Description": description,
            "Agency": agency, "County": county, "Bid Start Date": bid_start, "Bid End Date": bid_end,
            "Bid Open Date": bid_open, "Status": status_txt, "success": "fail"
        }

    
    try:
        select_by_visible_text_even_if_select2(driver, "nameSal", SALUTATION, timeout=20)
    except Exception:
        pass

    
    safe_type(driver, (By.ID, "nameFirst"), FIRST_NAME, 20)
    safe_type(driver, (By.ID, "nameLast"), LAST_NAME, 20)
    safe_type(driver, (By.ID, "address1"), ADDRESS1, 20)
    safe_type(driver, (By.ID, "city"), CITY, 20)
    safe_type(driver, (By.ID, "zip"), ZIP, 20)
    safe_type(driver, (By.ID, "email"), EMAIL, 20)

    
    try:
        select2_search_and_pick(driver, container_id_suffix="agency", term=agency, results_ul_id="select2-agency-results", timeout=30)
    except Exception:
        try:
            container = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.ID, "select2-agency-container"))
            )
            container.click()
            first = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#select2-agency-results li.select2-results__option"))
            )
            first.click()
        except Exception:
            pass

    
    records_text = compose_request_text(title, agency, bid_no)
    safe_set_textarea(driver, (By.ID, "records"), records_text, 30)

    
    print(f"Row {row_num}: Complete CAPTCHA in the browser, then press Enter here to continue...")
    input("Press Enter to submit...")


    try:
        wait_submit_enabled(driver, timeout=120)
    except TimeoutException:
        pass

    try:
        btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btn-submit")))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        driver.execute_script("window.scrollBy(0, -120);")
        try:
            btn.click()
        except ElementClickInterceptedException:
            try:
                driver.execute_script("arguments[0].click();", btn)
            except Exception:
                ActionChains(driver).move_to_element_with_offset(btn, 2, 2).click().perform()
    except Exception:
        pass

    result = detect_submission_status(driver, timeout=30)

    return {
        "Bid No": bid_no, "Bid Type": bid_type, "Title": title, "Description": description,
        "Agency": agency, "County": county, "Bid Start Date": bid_start, "Bid End Date": bid_end,
        "Bid Open Date": bid_open, "Status": status_txt, "success": result
    }

def main():
    df = pd.read_excel(INPUT_XLSX)
    if RUN_LIMIT:
        df = df.head(RUN_LIMIT)

    expected = [
        "Bid No", "Bid Type", "Title", "Description", "Agency", "County",
        "Bid Start Date", "Bid End Date", "Bid Open Date", "Status"
    ]
    for col in expected:
        if col not in df.columns:
            df[col] = ""

    driver = start_driver()
    rows_out = []
    try:
        for idx, row in df.iterrows():
            print(f"=== Processing row {idx+1} ===")
            result = submit_one_row(driver, row, idx + 1)
            rows_out.append(result)
            time.sleep(1)
    finally:
        driver.quit()

    out_df = pd.DataFrame(rows_out)[
        ["Bid No", "Bid Type", "Title", "Description", "Agency", "County",
         "Bid Start Date", "Bid End Date", "Bid Open Date", "Status", "success"]
    ]
    out_df.to_excel(OUTPUT_XLSX, index=False)
    print(f"Saved results to {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
