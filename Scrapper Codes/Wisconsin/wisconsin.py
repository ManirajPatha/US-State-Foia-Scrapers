
# pip install selenium pandas python-dateutil openpyxl

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
)
from dateutil import parser as dateparser
from datetime import datetime
import pandas as pd
import time
import re

START_URL = "https://vendornet.wi.gov/Contracts.aspx"
OUTPUT_XLSX = "vendornet_bids_first_100_2.xlsx"
MAX_ROWS = 100  

def new_driver():
    options = webdriver.ChromeOptions()
    
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(60)
    return driver

def wait_for_grid(driver, timeout=20):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "table.rgMasterTable"))
    )

def click_bids_tab(driver):
    try:
        link = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href,'Bids.aspx') and normalize-space()='Bids']"))
        )
        link.click()
    except TimeoutException:
        link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Bids"))
        )
        link.click()

def click_with_staleness_wait(driver, clickable_locator, container_locator=(By.CSS_SELECTOR, "table.rgMasterTable"), max_attempts=6):
    attempts = 0
    while attempts < max_attempts:
        attempts += 1
        try:
            old = WebDriverWait(driver, 20).until(EC.presence_of_element_located(container_locator))
            el = WebDriverWait(driver, 20).until(EC.element_to_be_clickable(clickable_locator))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            driver.execute_script("arguments[0].click();", el)
            WebDriverWait(driver, 20).until(EC.staleness_of(old))
            WebDriverWait(driver, 20).until(EC.presence_of_element_located(container_locator))
            return True
        except (StaleElementReferenceException, ElementClickInterceptedException, TimeoutException):
            time.sleep(0.3)
            continue
    return False

def ensure_checkbox_state_by_label_text(driver, label_text, should_be_checked: bool):
    label = WebDriverWait(driver, 12).until(
        EC.presence_of_element_located((By.XPATH, f"//label[contains(., '{label_text}')]"))
    )
    cls = (label.get_attribute("class") or "").strip()
    is_checked = "rfdCheckboxChecked" in cls
    if (should_be_checked and not is_checked) or ((not should_be_checked) and is_checked):
        try:
            old_table = driver.find_element(By.CSS_SELECTOR, "table.rgMasterTable")
            driver.execute_script("arguments[0].click();", label)
            WebDriverWait(driver, 15).until(EC.staleness_of(old_table))
            wait_for_grid(driver, 15)
        except Exception:
            driver.execute_script("arguments[0].click();", label)
            time.sleep(0.5)

def apply_filters(driver):
    
    try:
        ensure_checkbox_state_by_label_text(driver, "Include eSupplier", False)
    except Exception:
        pass
    try:
        ensure_checkbox_state_by_label_text(driver, "Awarded/ Canceled", True)
    except Exception:
        pass
    try:
        search_btn = driver.find_element(By.XPATH, "//*[self::input or self::button][@value='Search' or normalize-space()='Search']")
        old_table = driver.find_element(By.CSS_SELECTOR, "table.rgMasterTable")
        driver.execute_script("arguments[0].click();", search_btn)
        WebDriverWait(driver, 15).until(EC.staleness_of(old_table))
        wait_for_grid(driver, 15)
    except Exception:
        wait_for_grid(driver, 15)

def sort_by_available_date_desc(driver):
    header_locator = (By.XPATH, "//th[.//a[normalize-space()='Available Date']]//a")
    if not click_with_staleness_wait(driver, header_locator):
        raise TimeoutException("Failed to click Available Date header (1)")
    time.sleep(0.2)
    if not click_with_staleness_wait(driver, header_locator):
        raise TimeoutException("Failed to click Available Date header (2)")

def parse_row(tr):
    tds = tr.find_elements(By.TAG_NAME, "td")
    if len(tds) < 6:
        return None
    try:
        ref_anchor = tds[0].find_element(By.TAG_NAME, "a")
        reference = ref_anchor.text.strip()
        bid_url = ref_anchor.get_attribute("href") or ""
    except Exception:
        reference = tds[0].text.strip()
        bid_url = ""
    title = tds[1].text.strip()
    agency = tds[2].text.strip()
    available_date_raw = tds[3].text.strip()
    due_date_raw = tds[4].text.strip()
    e_supplier = False
    try:
        label = tds[5].find_element(By.XPATH, ".//label")
        cls = label.get_attribute("class") or ""
        e_supplier = "rfdCheckboxChecked" in cls
    except Exception:
        pass
    available_dt = None
    due_dt = None
    try:
        if available_date_raw:
            available_dt = dateparser.parse(available_date_raw, dayfirst=False)
    except Exception:
        pass
    try:
        if due_date_raw:
            due_dt = dateparser.parse(due_date_raw, dayfirst=False)
    except Exception:
        pass
    return {
        "Solicitation Reference #": reference,
        "Title": title,
        "Agency": agency,
        "Available Date": available_dt,
        "Due Date": due_dt,
        "Available in eSupplier": e_supplier,
        "Bid URL": bid_url,
    }

def collect_page_rows(driver):
    rows_data = []
    table = driver.find_element(By.CSS_SELECTOR, "table.rgMasterTable")
    tbody = table.find_element(By.TAG_NAME, "tbody")
    rows = tbody.find_elements(By.XPATH, "./tr[contains(@class,'rgRow') or contains(@class,'rgAltRow')]")
    for tr in rows:
        item = parse_row(tr)
        if item:
            rows_data.append(item)
    return rows_data


def _get_first_ref_text(driver):
    try:
        table = driver.find_element(By.CSS_SELECTOR, "table.rgMasterTable")
        tbody = table.find_element(By.TAGNAME, "tbody")
    except Exception:
        table = driver.find_element(By.CSS_SELECTOR, "table.rgMasterTable")
        tbody = table.find_element(By.TAG_NAME, "tbody")
    try:
        first_anchor = tbody.find_element(
            By.XPATH,
            "./tr[(contains(@class,'rgRow') or contains(@class,'rgAltRow'))][1]/td[1]//a"
        )
        return first_anchor.text.strip()
    except Exception:
        return ""

def _get_first_ref_text(driver):
    try:
        table = driver.find_element(By.CSS_SELECTOR, "table.rgMasterTable")
        tbody = table.find_element(By.TAG_NAME, "tbody")
        first_anchor = tbody.find_element(
            By.XPATH,
            "./tr[(contains(@class,'rgRow') or contains(@class,'rgAltRow'))][1]/td[1]//a"
        )
        return first_anchor.text.strip()
    except Exception:
        return ""

def _get_current_page_index(driver):
    try:
        span = driver.find_element(By.CSS_SELECTOR, ".rgPager span.rgCurrentPage")
        return int(span.text.strip())
    except Exception:
        pass
    try:
        inp = driver.find_element(By.CSS_SELECTOR, ".rgPager input.rgCurrentPageBox")
        return int(inp.get_attribute("value") or "0")
    except Exception:
        pass
    return None

def _find_enabled_next(driver):
    candidates = driver.find_elements(By.CSS_SELECTOR, ".rgPager a.rgPageNext, .rgPager button.rgPageNext, .rgPager input.rgPageNext")
    for el in candidates:
        cls = (el.get_attribute("class") or "")
        disabled_attr = el.get_attribute("disabled")
        if disabled_attr or "rgDisabled" in cls:
            continue
        return el
    return None

def _click_next(driver):
    el = WebDriverWait(driver, 6).until(lambda d: _find_enabled_next(d))
    driver.execute_script("arguments[0].click();", el)

def _click_numeric_page(driver, target_index):
    try:
        link = driver.find_element(By.XPATH, f"//div[contains(@class,'rgPager')]//div[contains(@class,'rgNumPart')]//a[normalize-space()='{target_index}']")
        driver.execute_script("arguments[0].click();", link)
        return True
    except Exception:
        return False

def _set_page_by_input(driver, target_index, timeout=10):
    try:
        inp = driver.find_element(By.CSS_SELECTOR, ".rgPager input.rgCurrentPageBox")
    except Exception:
        return False
    before_first = _get_first_ref_text(driver)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
    inp.clear()
    inp.send_keys(str(target_index))
    inp.send_keys(Keys.ENTER)
    def page_changed(drv):
        try:
            cur_idx = _get_current_page_index(drv)
            if cur_idx == target_index:
                return True
            cur_first = _get_first_ref_text(drv)
            return cur_first and cur_first != before_first
        except StaleElementReferenceException:
            return False
    WebDriverWait(driver, timeout, poll_frequency=0.2).until(page_changed)
    return True

def go_to_next_page(driver, timeout=10):
    before_first = _get_first_ref_text(driver)
    cur_idx = _get_current_page_index(driver)
    
    if cur_idx is not None and _click_numeric_page(driver, cur_idx + 1):
        pass
    elif _find_enabled_next(driver) is not None:
        _click_next(driver)
    elif cur_idx is not None:
        if not _set_page_by_input(driver, cur_idx + 1, timeout=timeout):
            return False
        return True
    else:
        return False
    
    def changed(drv):
        try:
            new_idx = _get_current_page_index(drv)
            if cur_idx is not None and new_idx is not None and new_idx != cur_idx:
                return True
            return _get_first_ref_text(drv) not in ("", before_first) and _get_first_ref_text(drv) != before_first
        except StaleElementReferenceException:
            return False
    WebDriverWait(driver, timeout, poll_frequency=0.2).until(changed)
    return True

def try_set_page_size(driver, target_size=100):
    
    try:
        select_el = driver.find_element(By.CSS_SELECTOR, ".rgPager select")
        old_first = _get_first_ref_text(driver)
        Select(select_el).select_by_visible_text(str(target_size))
        WebDriverWait(driver, 12).until(lambda d: _get_first_ref_text(d) != old_first and _get_first_ref_text(d) != "")
        return True
    except Exception:
        return False

def main():
    driver = new_driver()
    records = []
    pages = 0
    try:
        driver.get(START_URL)
        click_bids_tab(driver)
        wait_for_grid(driver)
        apply_filters(driver)
        sort_by_available_date_desc(driver)

        
        increased = try_set_page_size(driver, 100)

        while len(records) < MAX_ROWS:
            page_rows = collect_page_rows(driver)
            remaining = MAX_ROWS - len(records)
            records.extend(page_rows[:remaining])
            pages += 1
            if len(records) >= MAX_ROWS:
                break
            
            if not go_to_next_page(driver):
                break

        if records:
            df = pd.DataFrame(records)
            for col in ["Available Date", "Due Date"]:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors="coerce")
                    df[col] = df[col].dt.strftime("%Y-%m-%d %H:%M:%S").fillna("")
            df.to_excel(OUTPUT_XLSX, index=False)
            print(f"Saved {len(records)} rows from {pages} pages to {OUTPUT_XLSX}")
        else:
            print("No rows collected; check filters.")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()