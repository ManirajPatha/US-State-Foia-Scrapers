#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Nevada ePro — Advanced Search (Status=Closed) → Excel
HARD-CODED Selenium version (robust Search click + full pagination)

Page:
  https://nevadaepro.com/bso/view/search/external/advancedSearchBid.xhtml

Outputs columns:
  Bid Solicitation #, Organization Name, Contract #, Buyer, Description,
  Bid Opening Date, Bid Holder List, Awarded Vendor(s), Status, Alternate Id, Row URL
"""

import argparse
import os
import re
import sys
from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, JavascriptException

ADV_URL = "https://nevadaepro.com/bso/view/search/external/advancedSearchBid.xhtml"

TARGET_HEADERS = [
    "Bid Solicitation #",
    "Organization Name",
    "Contract #",
    "Buyer",
    "Description",
    "Bid Opening Date",
    "Bid Holder List",
    "Awarded Vendor(s)",
    "Status",
    "Alternate Id",
]

def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())

def build_driver(headless: bool) -> webdriver.Chrome:
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1600,1100")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opts)

def wait_for_boot(driver: webdriver.Chrome, timeout: int = 30):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='javax.faces.ViewState']"))
    )

def select_status_closed(driver: webdriver.Chrome, timeout: int = 20):
    # Status dropdown name often ends with :status
    sel_el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "select[name$=':status']"))
    )
    Select(sel_el)  # ensure it's a <select>
    sel = Select(sel_el)

    # Hardcode: choose the option containing "Closed"
    chosen = False
    for opt in sel.options:
        if "closed" in (opt.text or "").lower():
            sel.select_by_visible_text(opt.text)
            chosen = True
            break
    if not chosen:
        # Fallback by value
        for opt in sel.options:
            val = (opt.get_attribute("value") or "").lower()
            if "closed" in val or "close" in val:
                sel.select_by_value(opt.get_attribute("value"))
                chosen = True
                break
    if not chosen:
        raise RuntimeError("Could not find 'Closed' in Status dropdown")

def _try_click(driver: webdriver.Chrome, by, locator) -> bool:
    try:
        el = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((by, locator)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        try:
            el.click()
        except Exception:
            driver.execute_script("arguments[0].click();", el)
        return True
    except Exception:
        return False

def click_search_hardcoded(driver: webdriver.Chrome):
    """
    Aggressive, hard-coded strategies to trigger Search:
      1) Try common IDs (advSearchForm:search, bidSearchForm:search, etc.)
      2) Try visible text //button[.='Search'] or role/button variants
      3) Try input[type=submit][value='Search']
      4) Trigger Search button's onclick via JS if present
      5) Submit the form directly
      6) Send ENTER on the Status select
    """
    # 1) Known ID candidates seen on ePro/PrimeFaces skins
    id_candidates = [
        "advSearchForm:search", "advSearchForm:searchBtn", "advSearchForm:searchButton",
        "bidSearchForm:search", "bidSearchForm:searchBtn", "bidSearchForm:searchButton",
        # generic command button ids sometimes look like these:
        "searchForm:search", "searchForm:searchBtn", "searchForm:searchButton",
    ]
    for cid in id_candidates:
        if _try_click(driver, By.ID, cid):
            return

    # 2) Visible text button
    xpath_btns = [
        "//button[normalize-space()='Search']",
        "//button[contains(normalize-space(.),'Search')]",
        "//*[@role='button' and (normalize-space()='Search' or contains(normalize-space(.),'Search'))]",
    ]
    for xp in xpath_btns:
        if _try_click(driver, By.XPATH, xp):
            return

    # 3) Input submit
    if _try_click(driver, By.XPATH, "//input[@type='submit' and translate(@value,'SEARCH','search')='search']"):
        return

    # 4) Execute the onclick javascript of a likely Search element, if present
    try:
        candidates = driver.find_elements(By.XPATH, "//*[contains(@onclick,'PrimeFaces.ab') or contains(@onclick,'submit')]")
        for el in candidates:
            txt = (el.text or "").lower()
            val = (el.get_attribute("value") or "").lower()
            title = (el.get_attribute("title") or "").lower()
            if "search" in txt or "search" in val or "search" in title:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                driver.execute_script("var f = arguments[0]; if (f.onclick) { f.onclick(); } else { f.click(); }", el)
                return
    except JavascriptException:
        pass

    # 5) Submit the form element directly
    try:
        form = driver.find_element(By.XPATH, "//form[contains(@id,'advSearchForm') or contains(@name,'advSearchForm') or contains(@id,'search') or contains(@name,'search')]")
        driver.execute_script("arguments[0].submit();", form)
        return
    except NoSuchElementException:
        pass

    # 6) Send ENTER on the dropdown (often triggers default command)
    try:
        sel = driver.find_element(By.CSS_SELECTOR, "select[name$=':status']")
        sel.send_keys(Keys.ENTER)
        return
    except Exception:
        pass

    raise RuntimeError("Failed to click the Search button with hard-coded strategies")

def wait_for_results_table(driver: webdriver.Chrome, timeout: int = 40):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "tbody[id$='_data']"))
    )

def map_header_indices(driver: webdriver.Chrome) -> Dict[str, int]:
    tbody = driver.find_element(By.CSS_SELECTOR, "tbody[id$='_data']")
    table = tbody.find_element(By.XPATH, "./ancestor::table[1]")
    ths = table.find_elements(By.CSS_SELECTOR, "thead th")
    headers = [normalize(th.text) for th in ths]

    def find_idx(label: str) -> Optional[int]:
        for i, h in enumerate(headers):
            if h.lower() == label.lower():
                return i
        want = re.sub(r"[^a-z0-9]+", "", label.lower())
        for i, h in enumerate(headers):
            if re.sub(r"[^a-z0-9]+", "", h.lower()) == want:
                return i
        for i, h in enumerate(headers):
            if label.lower() in h.lower():
                return i
        return None

    return {label: find_idx(label) for label in TARGET_HEADERS}

def extract_rows_current_page(driver: webdriver.Chrome, idx_map: Dict[str, int]) -> List[Dict]:
    tbody = driver.find_element(By.CSS_SELECTOR, "tbody[id$='_data']")
    trs = tbody.find_elements(By.CSS_SELECTOR, ":scope > tr")
    rows: List[Dict] = []
    for tr in trs:
        tds = tr.find_elements(By.CSS_SELECTOR, ":scope > td, :scope > th")
        if not tds:
            continue
        rec: Dict[str, str] = {}
        i0 = idx_map.get("Bid Solicitation #")
        if i0 is not None and i0 < len(tds):
            cell = tds[i0]
            try:
                link = cell.find_element(By.CSS_SELECTOR, "a[href]")
                rec["Bid Solicitation #"] = normalize(link.text)
                rec["Row URL"] = link.get_attribute("href")
            except Exception:
                rec["Bid Solicitation #"] = normalize(cell.text)
                rec["Row URL"] = ""
        else:
            rec["Bid Solicitation #"] = ""
            rec["Row URL"] = ""

        for col in TARGET_HEADERS:
            if col == "Bid Solicitation #":
                continue
            idx = idx_map.get(col)
            rec[col] = normalize(tds[idx].text) if idx is not None and idx < len(tds) else ""
        rows.append(rec)
    return rows

def paginator_has_next(driver: webdriver.Chrome) -> bool:
    # Typical PrimeFaces "next"
    for el in driver.find_elements(By.CSS_SELECTOR, "a.ui-paginator-next, span.ui-paginator-next"):
        cls = (el.get_attribute("class") or "").lower()
        aria = (el.get_attribute("aria-disabled") or "").lower()
        if "ui-state-disabled" in cls or aria in ("true", "yes", "1"):
            continue
        return True
    # If no explicit "next", try finding a higher page number than active
    active_nums = driver.find_elements(By.CSS_SELECTOR, ".ui-paginator-page.ui-state-active")
    if active_nums:
        try:
            cur = int(active_nums[0].text.strip())
            for a in driver.find_elements(By.CSS_SELECTOR, "a.ui-paginator-page"):
                try:
                    if int(a.text.strip()) > cur:
                        return True
                except ValueError:
                    continue
        except ValueError:
            pass
    return False

def click_next_page(driver: webdriver.Chrome, timeout: int = 25) -> bool:
    # Try standard "next"
    try:
        next_el = WebDriverWait(driver, timeout//2).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.ui-paginator-next"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", next_el)
        driver.execute_script("arguments[0].click();", next_el)
        # Wait for table body to refresh (staleness of any row)
        old_first = driver.find_element(By.CSS_SELECTOR, "tbody[id$='_data'] tr")
        WebDriverWait(driver, timeout).until(EC.staleness_of(old_first))
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody[id$='_data'] tr")))
        return True
    except Exception:
        # Fallback numeric page progression
        try:
            active = driver.find_element(By.CSS_SELECTOR, ".ui-paginator-page.ui-state-active")
            cur = int(active.text.strip())
            next_link = driver.find_element(By.XPATH, f"//a[contains(@class,'ui-paginator-page') and normalize-space()='{cur+1}']")
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", next_link)
            driver.execute_script("arguments[0].click();", next_link)
            WebDriverWait(driver, timeout).until(
                EC.text_to_be_present_in_element((By.CSS_SELECTOR, ".ui-paginator-page.ui-state-active"), str(cur+1))
            )
            WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody[id$='_data'] tr")))
            return True
        except Exception:
            return False

def scrape_all_closed(out_folder: str, headless: bool) -> Optional[str]:
    driver = build_driver(headless)
    try:
        driver.get(ADV_URL)
        wait_for_boot(driver)
        select_status_closed(driver)

        # HARD-CODED & MULTI-STRATEGY CLICK on Search
        click_search_hardcoded(driver)

        # Wait for grid
        try:
            wait_for_results_table(driver, timeout=40)
        except TimeoutException:
            # Some skins re-render the form then grid; wait again after a brief scroll
            driver.execute_script("window.scrollBy(0, 400);")
            wait_for_results_table(driver, timeout=40)

        # Map headers once
        idx_map = map_header_indices(driver)

        all_rows: List[Dict] = []
        page = 1
        while True:
            rows = extract_rows_current_page(driver, idx_map)
            print(f"[INFO] Page {page}: {len(rows)} row(s)")
            all_rows.extend(rows)

            if not paginator_has_next(driver):
                break
            if not click_next_page(driver):
                break
            page += 1

        if not all_rows:
            print("[INFO] No rows found.")
            return None

        os.makedirs(out_folder, exist_ok=True)
        out_path = os.path.join(out_folder, f"nevada_closed_results_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
        df = pd.DataFrame.from_records(all_rows, columns=TARGET_HEADERS + ["Row URL"])
        df.to_excel(out_path, index=False)
        print(f"[INFO] Wrote {len(df)} rows → {out_path}")
        return out_path
    finally:
        driver.quit()

def main():
    ap = argparse.ArgumentParser(description="Nevada ePro (Advanced) — Status=Closed → Excel (Selenium, hard-coded search)")
    ap.add_argument("--out", default=".", help="Output folder")
    ap.add_argument("--headless", action="store_true", help="Run Chrome headless")
    args = ap.parse_args()

    out_path = scrape_all_closed(args.out, headless=args.headless)
    if not out_path:
        sys.exit(2)

if __name__ == "__main__":
    main()
