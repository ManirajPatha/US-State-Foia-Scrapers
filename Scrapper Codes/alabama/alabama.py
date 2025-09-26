"""
Alabama VSS scraper — Excel output (dedup, max 20, count-first)

Flow
- Logs into Alabama STAARS VSS
- Business Opportunities → Advanced Search → Status: Open (ONLY)
- Pass 1: count how many Open rows exist across pages (no details opened)
- Pass 2: fetch up to MAX_RESULTS rows (default 20), skipping duplicates by notice_id
- Writes/updates a stable Excel file: alabama_vss_opportunities.xlsx (deduped)

Requirements:
  pip install selenium pandas
  Chrome + matching ChromeDriver on PATH
"""

import os
import re
import time
import tempfile
from datetime import datetime
from typing import List, Dict, Optional, Tuple, Set

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# -------------------
# 0) Config
# -------------------
ALABAMA_VSS_URL = "https://procurement.staars.alabama.gov/webapp/PRDVSS1X1/AltSelfService"
USERNAME = os.getenv("AL_VSS_USER", "sakesh_841")
PASSWORD = os.getenv("AL_VSS_PASS", "Password123")

# Hard cap for this run (per your request)
MAX_RESULTS = 20

# Stable master output (deduped across runs)
MASTER_XLSX = os.getenv("AL_VSS_OUTPUT", "alabama_vss_opportunities.xlsx")

# -------------------
# 1) WebDriver helpers
# -------------------
def create_chrome_driver(download_dir: str) -> webdriver.Chrome:
    chrome_options = Options()
    # chrome_options.add_argument("--headless=new")  # optional
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_experimental_option(
        "prefs",
        {
            "download.prompt_for_download": False,
            "download.default_directory": download_dir,
            "plugins.always_open_pdf_externally": True,
            "download.directory_upgrade": True,
            "profile.default_content_settings.popups": 0,
            "profile.default_content_setting_values.automatic_downloads": 1,
        },
    )
    return webdriver.Chrome(options=chrome_options)

def cleanup_chrome_driver(driver: webdriver.Chrome):
    try:
        driver.quit()
    except Exception:
        pass

# -------------------
# 2) Navigation helpers
# -------------------
def _ensure_frame(driver: webdriver.Chrome, wait: WebDriverWait, name: str):
    driver.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, name)))

def login(driver: webdriver.Chrome, wait: WebDriverWait, url: str) -> None:
    driver.get(url)
    wait.until(EC.presence_of_element_located((By.ID, "login")))
    driver.find_element(By.ID, "login").send_keys(USERNAME)
    driver.find_element(By.ID, "password").send_keys(PASSWORD)
    driver.find_element(By.ID, "vslogin").click()
    wait.until(EC.number_of_windows_to_be(2))
    driver.switch_to.window(driver.window_handles[-1])

def open_business_opportunities(driver: webdriver.Chrome, wait: WebDriverWait):
    time.sleep(2)
    _ensure_frame(driver, wait, "Startup")
    driver.find_element(By.LINK_TEXT, "Business Opportunities").click()

def apply_open_filter(driver: webdriver.Chrome, wait: WebDriverWait):
    driver.switch_to.default_content()
    _ensure_frame(driver, wait, "Display")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@title='Advanced Search']"))).click()
    # Status = Open only
    wait.until(EC.element_to_be_clickable((By.ID, "txtT1SO_STAQry"))).click()
    driver.find_element(By.XPATH, "//select[@id='txtT1SO_STAQry']/option[@value='O']").click()
    driver.find_element(By.XPATH, "//input[@type='submit' and @value='Go']").click()
    # Wait for grid rows
    WebDriverWait(driver, 20).until(
        EC.presence_of_all_elements_located(
            (By.XPATH, "//tr[contains(@class, 'ADVGridOddRow') or contains(@class, 'advgridevenrow')]")
        )
    )

def _wait_grid_rows(driver: webdriver.Chrome, wait: WebDriverWait):
    _ensure_frame(driver, wait, "Display")
    wait.until(
        EC.presence_of_all_elements_located(
            (By.XPATH, "//tr[contains(@class, 'ADVGridOddRow') or contains(@class, 'advgridevenrow')]")
        )
    )

def _find_rows(driver: webdriver.Chrome):
    return driver.find_elements(
        By.XPATH,
        "//tr[contains(@class, 'ADVGridOddRow') or contains(@class, 'advgridevenrow')]",
    )

def click_next_page_if_any(driver: webdriver.Chrome) -> bool:
    for by, selector in [
        (By.XPATH, "//input[@type='submit' and translate(@value,'next','NEXT')='NEXT']"),
        (By.XPATH, "//a[normalize-space(translate(.,'next','NEXT'))='NEXT']"),
        (By.XPATH, "//button[normalize-space(translate(.,'next','NEXT'))='NEXT']"),
        (By.XPATH, "//a[contains(@href,'PageNext') or contains(@title,'Next')]"),
        (By.XPATH, "//a[img[contains(@alt,'Next')]]"),
        (By.XPATH, "//*[self::a or self::button or self::input][contains(., 'Next')]"),
    ]:
        try:
            el = driver.find_element(by, selector)
            driver.execute_script("arguments[0].click();", el)
            return True
        except Exception:
            continue
    return False

def return_to_results_fresh(driver: webdriver.Chrome, wait: WebDriverWait, page_no: int):
    """Rebuild results view without browser back/new tab, then jump to page_no (1-based)."""
    driver.switch_to.default_content()
    try:
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.LINK_TEXT, "Home"))).click()
    except Exception:
        pass
    open_business_opportunities(driver, wait)
    apply_open_filter(driver, wait)
    for _ in range(max(0, page_no - 1)):
        if not click_next_page_if_any(driver):
            break
        _wait_grid_rows(driver, wait)

# -------------------
# 3) Excel de-dup helpers
# -------------------
def load_existing_master(path: str) -> Tuple[pd.DataFrame, Set[str]]:
    if os.path.exists(path):
        try:
            df = pd.read_excel(path, dtype=str)
            ids = set(df.get("notice_id", pd.Series(dtype=str)).dropna().astype(str).str.strip())
            return df, ids
        except Exception:
            pass
    # empty
    return pd.DataFrame(columns=[
        "notice_id","title","issued_date","closing","department","buyer","category","doc_type",
        "contact_email","details_url","attachments","source","page_no","row_index_on_page"
    ]), set()

def _safe_get_text(driver: webdriver.Chrome, by: By, selector: str) -> str:
    try:
        return driver.find_element(by, selector).text.strip()
    except Exception:
        return ""

def _guess_notice_id_from_row_text(txt: str) -> Optional[str]:
    # STAARS solicitation IDs are often long numeric strings
    m = re.search(r"\b\d{7,}\b", txt or "")
    return m.group(0) if m else None

# -------------------
# 4) Planning pass — count and plan rows to fetch
# -------------------
def plan_rows_to_fetch(driver: webdriver.Chrome, wait: WebDriverWait, limit: int = MAX_RESULTS) -> Tuple[int, List[Tuple[int, int, Optional[str]]]]:
    """
    Returns (total_open_rows, plan) where plan is a list of tuples:
      (page_no, row_index_on_page, guessed_notice_id)
    Only up to `limit` items are included in the plan.
    """
    open_business_opportunities(driver, wait)
    apply_open_filter(driver, wait)

    total = 0
    page = 1
    plan: List[Tuple[int, int, Optional[str]]] = []

    while True:
        _wait_grid_rows(driver, wait)
        rows = _find_rows(driver)
        count_here = len(rows)
        print(f"[DEBUG] Count pass — Page {page}: {count_here} row(s)")
        total += count_here

        # add row indices for this page until we hit the limit
        for idx in range(count_here):
            if len(plan) >= limit:
                break
            guess = _guess_notice_id_from_row_text(rows[idx].text)
            plan.append((page, idx, guess))

        if len(plan) >= limit:
            break

        # move to next page if exists; if not, stop
        if not click_next_page_if_any(driver):
            break
        page += 1

    return total, plan

# -------------------
# 5) Scraper core — execute the plan (dedup aware)
# -------------------
def scrape_plan(driver: webdriver.Chrome, wait: WebDriverWait, plan: List[Tuple[int, int, Optional[str]]], existing_ids: Set[str]) -> List[Dict]:
    records: List[Dict] = []
    fetched = 0

    for (page_no, row_idx, id_guess) in plan:
        # Rebuild grid and jump to target page/row
        return_to_results_fresh(driver, wait, page_no)
        _wait_grid_rows(driver, wait)
        rows = _find_rows(driver)
        if row_idx >= len(rows):
            print(f"[WARN] Planned row {row_idx+1} missing on page {page_no}; skipping")
            continue

        # quick dedup using guess
        if id_guess and id_guess in existing_ids:
            print(f"[DEBUG] Skip duplicate (grid guess): {id_guess}")
            continue

        # open details (POST)
        try:
            details_btn = rows[row_idx].find_element(By.XPATH, ".//input[@type='submit' and @value='Details']")
            driver.execute_script("arguments[0].click();", details_btn)
            _ensure_frame(driver, wait, "Display")
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "Solicitation1")))

            header_raw = _safe_get_text(driver, By.ID, "Solicitation1")
            clean = re.sub(r"^Solicitation:\s*", "", header_raw or "")
            parts = re.split(r"\s{2,}", clean, maxsplit=1)
            notice_id = (parts[0].strip() if parts else "") or id_guess or ""
            title = parts[1].strip() if len(parts) > 1 else ""

            if notice_id and notice_id in existing_ids:
                print(f"[DEBUG] Skip duplicate (detail actual): {notice_id}")
            else:
                publish_date_str = _safe_get_text(driver, By.XPATH, "//td[contains(., 'Issued:')]/span[@class='date'][1]")
                closing_date_str = _safe_get_text(driver, By.XPATH, "//div[@id='Solicitation6']/parent::td/following-sibling::td")
                dept = _safe_get_text(driver, By.XPATH, "//b[text()='Doc Dept:']/parent::td/following-sibling::td")
                buyer_name = _safe_get_text(driver, By.XPATH, "//div[@id='Solicitation3']/parent::td/following-sibling::td")
                category = _safe_get_text(driver, By.XPATH, "//b[text()='Category:']/parent::td/following-sibling::td")
                doc_type = _safe_get_text(driver, By.XPATH, "//b[text()='Type:']/parent::td/following-sibling::td")
                email = _safe_get_text(driver, By.XPATH, "//a[contains(@href, 'mailto:')]")
                current_url = driver.current_url

                attachments: List[Dict] = []
                try:
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "ATTACHMENTS_TAB"))).click()
                    table = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "tblT34IN_OBJ_ATT_CTLG"))
                    )
                    links = table.find_elements(By.XPATH, ".//a[@title='View Solicitation File']")
                    for link in links:
                        attachments.append({"name": link.text.strip(), "href": link.get_attribute("href") or ""})
                except Exception:
                    pass

                records.append({
                    "notice_id": notice_id,
                    "title": title,
                    "issued_date": publish_date_str,
                    "closing": closing_date_str,
                    "department": dept,
                    "buyer": buyer_name,
                    "category": category,
                    "doc_type": doc_type,
                    "contact_email": email,
                    "details_url": current_url,
                    "attachments": "; ".join([f"{a['name']}|{a['href']}" for a in attachments]) if attachments else "",
                    "source": "Alabama VSS",
                    "page_no": page_no,
                    "row_index_on_page": row_idx + 1,
                })
                fetched += 1
                print(f"[DEBUG] Collected {fetched}/{len(plan)}: {notice_id or title}")

        finally:
            # restore grid state for the next item in the plan
            return_to_results_fresh(driver, wait, page_no)

    return records

# -------------------
# 6) Orchestrator (count-first → scrape; Excel dedup)
# -------------------
def scrape_alabama_to_excel(url: str = ALABAMA_VSS_URL, limit: int = MAX_RESULTS) -> str:
    download_dir = tempfile.mkdtemp(prefix="downloads_")
    driver = create_chrome_driver(download_dir)
    wait = WebDriverWait(driver, 20)

    try:
        existing_df, existing_ids = load_existing_master(MASTER_XLSX)
        print(f"[INFO] Existing rows in master: {len(existing_df)}")

        login(driver, wait, url)

        total_open, plan = plan_rows_to_fetch(driver, wait, limit=limit)
        print(f"[INFO] Total 'Open' rows available now: {total_open}")
        print(f"[INFO] Planning to fetch up to {len(plan)} row(s)")

        new_rows = scrape_plan(driver, wait, plan, existing_ids)
        new_df = pd.DataFrame(new_rows) if new_rows else pd.DataFrame(columns=[
            "notice_id","title","issued_date","closing","department","buyer","category","doc_type",
            "contact_email","details_url","attachments","source","page_no","row_index_on_page"
        ])

        # Combine & dedup by notice_id (keep first)
        combined = pd.concat([existing_df, new_df], ignore_index=True)
        if not combined.empty:
            combined["notice_id"] = combined["notice_id"].astype(str).str.strip()
            combined = combined.drop_duplicates(subset=["notice_id"], keep="first").reset_index(drop=True)

        combined.to_excel(MASTER_XLSX, index=False)
        print(f"Excel written: {MASTER_XLSX} (rows: {len(combined)}, added_this_run: {len(new_df)})")
        return MASTER_XLSX

    finally:
        cleanup_chrome_driver(driver)

if __name__ == "__main__":
    print(f"Starting Alabama VSS scrape → {ALABAMA_VSS_URL}")
    outfile = scrape_alabama_to_excel()
    print(f"Done. Output: {outfile}")
