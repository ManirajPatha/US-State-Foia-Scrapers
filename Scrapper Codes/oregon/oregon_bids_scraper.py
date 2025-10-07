# filename: oregon_buys_scrape_bid_to_po_sturdy.py
import csv
import time
import logging
from typing import Dict, List

from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
)

from openpyxl import Workbook

URL = "https://oregonbuys.gov/bso/view/search/external/advancedSearchBid.xhtml?openBids=true"
CSV_PATH = "or_buys_results.csv"
XLSX_PATH = "or_buys_results.xlsx"

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

TARGET_HEADERS = {
    "bid_solicitation": "Bid Solicitation #",
    "organization": "Organization Name",
    "awarded": "Awarded Vendor(s)",
}

def norm(s: str) -> str:
    return " ".join((s or "").strip().split()).lower()

def scroll_into_view(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.2)

def soft_scroll(driver, px=400):
    driver.execute_script(f"window.scrollBy(0, {px});")
    time.sleep(0.15)

def open_browser():
    service = Service()  # uses geckodriver on PATH
    options = webdriver.FirefoxOptions()
    driver = webdriver.Firefox(service=service, options=options)
    try:
        driver.maximize_window()
    except Exception:
        pass
    driver.set_window_size(1600, 1200)
    return driver

def open_and_search(driver):
    wait = WebDriverWait(driver, 35)

    logging.info("Opening OregonBuys Advanced Search…")
    driver.get(URL)
    time.sleep(1)
    soft_scroll(driver, 300)

    # Expand Advanced Search (Periscope/S2G)
    adv_legend = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//legend[contains(@class,'ui-fieldset-legend')][contains(., 'Advanced Search')]")
        )
    )
    scroll_into_view(driver, adv_legend)
    adv_legend.click()

    # Wait for panel
    wait.until(EC.visibility_of_element_located((By.ID, "advSearchFormFields")))

    # Try to set Status = Bid to PO if present; otherwise proceed with defaults (e.g., open bids)
    try:
        status_select_el = wait.until(EC.element_to_be_clickable((By.ID, "bidSearchForm:status")))
        scroll_into_view(driver, status_select_el)
        select = Select(status_select_el)
        values = [o.get_attribute("value") for o in select.options]
        if "2BPO" in values:
            select.select_by_value("2BPO")
            logging.info("Selected Status = Bid to PO.")
        else:
            logging.info("Status option 'Bid to PO' not found; continuing with default Status.")
    except TimeoutException:
        logging.info("Status dropdown not found; continuing with default filters.")

    # Click Search
    search_btn = wait.until(EC.element_to_be_clickable((By.ID, "bidSearchForm:btnBidSearch")))
    scroll_into_view(driver, search_btn)
    search_btn.click()

    # Wait for results to load
    logging.info("Waiting 10 seconds for results to load…")
    time.sleep(10)

    # Confirm results tbody is present (or empty table still renders the tbody)
    wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, "tbody#bidSearchResultsForm\\:bidResultId_data.ui-datatable-data.ui-widget-content")
    ))

def map_header_indexes(driver) -> Dict[str, int]:
    wait = WebDriverWait(driver, 20)
    thead = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table[role='grid'] thead")))
    ths = thead.find_elements(By.CSS_SELECTOR, "th")

    header_map: Dict[str, int] = {}
    for idx, th in enumerate(ths):
        header_text = th.text.strip() or (th.get_attribute("aria-label") or th.get_attribute("title") or "").strip()
        header_text_norm = norm(header_text)
        for key, label in TARGET_HEADERS.items():
            if header_text_norm == norm(label):
                header_map[key] = idx

    if len(header_map) < 3:
        for idx, th in enumerate(ths):
            try:
                span = th.find_element(By.CSS_SELECTOR, "span")
                txt = norm(span.text)
                for key, label in TARGET_HEADERS.items():
                    if txt == norm(label) and key not in header_map:
                        header_map[key] = idx
            except Exception:
                continue

    missing = [k for k in TARGET_HEADERS if k not in header_map]
    if missing:
        logging.warning(f"Could not map all headers, missing: {missing}. Will use in-cell labels as fallback.")
    return header_map

def get_cell_value(driver, td) -> str:
    label_txt = ""
    try:
        label_el = td.find_element(By.CSS_SELECTOR, "span.ui-column-title")
        label_txt = label_el.text.strip()
    except Exception:
        pass

    try:
        a = td.find_element(By.TAG_NAME, "a")
        link_txt = a.text.strip()
        if not link_txt:
            link_txt = (a.get_attribute("title") or "").strip()
        if link_txt:
            return link_txt
    except Exception:
        pass

    try:
        spans = td.find_elements(By.TAG_NAME, "span")
        for sp in spans:
            t = (sp.get_attribute("title") or "").strip()
            if t:
                return t
    except Exception:
        pass

    try:
        inner_text = driver.execute_script("return arguments[0].innerText;", td).strip()
    except Exception:
        inner_text = td.text.strip()

    if label_txt and inner_text.startswith(label_txt):
        inner_text = inner_text[len(label_txt):].strip(" \n:\u00A0\t")

    inner_text = " ".join(inner_text.split())
    return inner_text

def scrape_current_page(driver, header_map: Dict[str, int]) -> List[dict]:
    wait = WebDriverWait(driver, 20)
    tbody = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, "tbody#bidSearchResultsForm\\:bidResultId_data.ui-datatable-data.ui-widget-content")
    ))
    scroll_into_view(driver, tbody)

    rows = tbody.find_elements(By.CSS_SELECTOR, "tr.ui-widget-content")
    logging.info(f"Rows on page: {len(rows)}")

    # Early out if the page is empty
    if not rows:
        return []

    records = []
    for r in rows:
        for attempt in range(3):
            try:
                scroll_into_view(driver, r)
                tds = r.find_elements(By.CSS_SELECTOR, "td[role='gridcell']")

                bid = org = awarded = ""

                if "bid_solicitation" in header_map and header_map["bid_solicitation"] < len(tds):
                    bid = get_cell_value(driver, tds[header_map["bid_solicitation"]]).strip()
                if "organization" in header_map and header_map["organization"] < len(tds):
                    org = get_cell_value(driver, tds[header_map["organization"]]).strip()
                if "awarded" in header_map and header_map["awarded"] < len(tds):
                    awarded = get_cell_value(driver, tds[header_map["awarded"]]).strip()

                if not (org and awarded):
                    for td in tds:
                        try:
                            lbl = td.find_element(By.CSS_SELECTOR, "span.ui-column-title").text.strip().lower()
                        except Exception:
                            lbl = ""
                        if not lbl:
                            continue
                        val = get_cell_value(driver, td)
                        if not org and lbl == "organization name":
                            org = val
                        elif not awarded and lbl == "awarded vendor(s)":
                            awarded = val

                if not bid:
                    try:
                        a = r.find_element(By.TAG_NAME, "a")
                        bid = a.text.strip() or (a.get_attribute("title") or "").strip()
                    except Exception:
                        pass

                records.append({
                    "Bid Solicitation #": bid,
                    "Organization Name": org,
                    "Awarded Vendor(s)": awarded,
                })

                soft_scroll(driver, 120)
                break
            except StaleElementReferenceException:
                logging.warning("Stale row encountered, retrying (%d/3)", attempt + 1)
                time.sleep(0.12)
                if attempt == 2:
                    logging.warning("Stale row encountered, skipping after retries.")
                    break
    return records

def is_next_disabled(driver) -> bool:
    try:
        next_a = driver.find_element(By.XPATH, "//a[contains(@class,'ui-paginator-next') and @aria-label='Next Page']")
    except Exception:
        return True
    cls = (next_a.get_attribute("class") or "")
    if "ui-state-disabled" in cls:
        return True
    aria_dis = (next_a.get_attribute("aria-disabled") or "").lower()
    if aria_dis == "true":
        return True
    return False

def click_next(driver):
    next_a = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//a[contains(@class,'ui-paginator-next') and @aria-label='Next Page']"))
    )
    scroll_into_view(driver, next_a)
    try:
        next_a.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", next_a)
    time.sleep(1.0)

def write_csv(path: str, rows: List[dict]):
    fields = ["Bid Solicitation #", "Organization Name", "Awarded Vendor(s)"]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow(r)

def write_xlsx(path: str, rows: List[dict]):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bid to PO"
    headers = ["Bid Solicitation #", "Organization Name", "Awarded Vendor(s)"]
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h, "") for h in headers])
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 42
    ws.column_dimensions['C'].width = 42
    wb.save(path)

def main():
    driver = open_browser()
    all_rows: List[dict] = []
    try:
        open_and_search(driver)

        header_map = map_header_indexes(driver)
        logging.info(f"Header index map: {header_map}")

        page_num = 0
        MAX_PAGES = 200

        while True:
            page_num += 1
            logging.info(f"Scraping page {page_num} …")
            page_rows = scrape_current_page(driver, header_map)
            if not page_rows and page_num == 1:
                # First page already empty — we can stop immediately
                logging.info("No rows found on the first page.")
                break

            all_rows.extend(page_rows)

            if is_next_disabled(driver):
                logging.info("Next page is disabled — end of results.")
                break

            try:
                click_next(driver)
            except TimeoutException:
                logging.info("No next button found — stopping.")
                break

            if page_num >= MAX_PAGES:
                logging.warning("Hit safety MAX_PAGES; stopping pagination.")
                break

        # === NEW BEHAVIOR: only write files when we actually have rows ===
        if not all_rows:
            print("No records")
            logging.info("No records — skipping CSV/XLSX creation.")
            return

        write_csv(CSV_PATH, all_rows)
        write_xlsx(XLSX_PATH, all_rows)
        logging.info(f"✅ Wrote {len(all_rows)} rows to {CSV_PATH} and {XLSX_PATH}")

    finally:
        try:
            driver.quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()