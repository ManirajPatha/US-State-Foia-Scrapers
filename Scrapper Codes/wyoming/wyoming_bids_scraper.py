# filename: wy_closed_bids_scraper.py
import csv
import logging
import time
from dataclasses import dataclass
from typing import List, Dict, Optional

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    StaleElementReferenceException,
    ElementClickInterceptedException,
    NoSuchElementException,
    TimeoutException,
)

START_URL = "https://www.publicpurchase.com/gems/wyominggsd,wy/buyer/public/publicClosedBidsInfo"
CSV_PATH = "wy_closed_bids.csv"       # final, filtered output
XLSX_PATH = "wy_closed_bids.xlsx"     # final, filtered output
PAGE_LOAD_TIMEOUT = 30
W = 15  # seconds

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


@dataclass
class BidRow:
    title: str
    status: str
    end_date: str


def build_driver() -> webdriver.Firefox:
    opts = Options()
    # Keep it visible so you can watch the activity:
    # opts.add_argument("-headless")
    driver = webdriver.Firefox(options=opts)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return driver


# -------------------- helpers --------------------

def smart_text(el) -> str:
    try:
        return " ".join(el.text.replace("\xa0", " ").split())
    except Exception:
        return ""


def switch_into_table_iframe(driver) -> None:
    """If results table isn't in top doc, scan iframes and switch to the one that contains it."""
    if driver.find_elements(By.CSS_SELECTOR, "table.tabHome"):
        return
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    for i, fr in enumerate(frames):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(fr)
            if driver.find_elements(By.CSS_SELECTOR, "table.tabHome"):
                logging.info(f"Switched into iframe index {i} containing the results table.")
                return
        except Exception:
            continue
    driver.switch_to.default_content()


def get_visible_tab_home(driver):
    tables = driver.find_elements(By.CSS_SELECTOR, "table.tabHome")
    for t in tables:
        if t.is_displayed():
            return t
    return None


def first_row_signature(table_el) -> str:
    try:
        tbody = table_el.find_element(By.TAG_NAME, "tbody")
        trs = tbody.find_elements(By.TAG_NAME, "tr")
        if not trs:
            return ""
        cells = trs[0].find_elements(By.TAG_NAME, "td")
        return " | ".join(smart_text(c) for c in cells)
    except Exception:
        return ""


def resolve_indices(table_el) -> Dict[str, int]:
    """
    Choose column indices, preferring Award Status for 'status'.
    Fallbacks cover missing/unknown headers.
    """
    # Map headers if present
    header_texts: Dict[int, str] = {}
    try:
        thead = table_el.find_element(By.TAG_NAME, "thead")
        ths = thead.find_elements(By.TAG_NAME, "th")
        for i, th in enumerate(ths):
            name = smart_text(th).lower()
            if name:
                header_texts[i] = name
    except Exception:
        pass

    # Probe first row for fallback
    try:
        tbody = table_el.find_element(By.TAG_NAME, "tbody")
        trs = tbody.find_elements(By.TAG_NAME, "tr")
        td_count = len(trs[0].find_elements(By.TAG_NAME, "td")) if trs else 0
    except Exception:
        td_count = 0

    def find_col(pred):
        for idx, txt in header_texts.items():
            if pred(txt):
                return idx
        return None

    # Title
    title_idx = find_col(lambda t: any(k in t for k in ["title", "description", "name"]))  # flexible
    if title_idx is None:
        title_idx = 0 if td_count > 0 else -1

    # Status — prefer "award status", then generic "status"
    award_status_idx = find_col(lambda t: "status" in t and "award" in t)
    generic_status_idx = find_col(lambda t: t == "status" or "bid status" in t or ("status" in t and "award" not in t))

    if award_status_idx is not None:
        status_idx = award_status_idx
    elif generic_status_idx is not None:
        status_idx = generic_status_idx
    else:
        status_idx = max(0, td_count - 2)  # heuristic

    # End date
    end_idx = find_col(lambda t: any(k in t for k in ["end date", "closing date", "close date", "bid end date", "enddate"]))
    if end_idx is None:
        end_idx = max(0, td_count - 1)

    return {"title": title_idx, "status": status_idx, "end_date": end_idx}


def scrape_current_page_rows(driver) -> List[BidRow]:
    wait = WebDriverWait(driver, W)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.tabHome")))
    time.sleep(0.3)

    table_el = get_visible_tab_home(driver)
    if not table_el:
        logging.warning("No visible table.tabHome on page.")
        return []

    idx = resolve_indices(table_el)
    tbody = table_el.find_element(By.TAG_NAME, "tbody")
    trs = tbody.find_elements(By.TAG_NAME, "tr")

    out: List[BidRow] = []
    for tr in trs:
        try:
            tds = tr.find_elements(By.TAG_NAME, "td")
            if not tds:
                continue

            def cell_text(i: int) -> str:
                if i < 0 or i >= len(tds):
                    return ""
                td = tds[i]
                links = td.find_elements(By.TAG_NAME, "a")
                if links:
                    return smart_text(links[0]) or smart_text(td)
                return smart_text(td)

            # pull basic fields
            title = cell_text(idx["title"])
            status = cell_text(idx["status"])
            end_date = cell_text(idx["end_date"])

            # --- SAFETY NET: if any cell says FINALIZED, treat row as FINALIZED ---
            # This covers cases when header mapping is off or table variants exist.
            any_finalized = any(smart_text(td).strip().upper() == "FINALIZED" for td in tds)
            if any_finalized:
                status = "FINALIZED"

            out.append(BidRow(title=title, status=status, end_date=end_date))
        except StaleElementReferenceException:
            continue
    logging.info(f"Scraped {len(out)} row(s) from this page.")
    return out


def dismiss_cookies_if_present(driver):
    try:
        btns = driver.find_elements(By.XPATH, "//button[contains(., 'Accept') or contains(., 'I Accept')]")
        for b in btns:
            if b.is_displayed():
                b.click()
                time.sleep(0.2)
                break
    except Exception:
        pass


# -------------------- Apply “Last 3 months” filter + Search --------------------

def apply_last_3_months_filter_and_search(driver):
    """
    In table.searchParamTab, set endingTimeInterval -> '3m', click input#searchButton,
    and wait for the results to load.
    """
    logging.info("Applying filter: endingTimeInterval = '3m' and clicking Search…")

    # Controls are in top document
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    wait = WebDriverWait(driver, W)
    param_table = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.searchParamTab")))

    # Select the dropdown
    try:
        select_el = param_table.find_element(By.CSS_SELECTOR, "select#endingTimeInterval[name='endingTimeInterval']")
    except NoSuchElementException:
        select_el = driver.find_element(By.ID, "endingTimeInterval")

    Select(select_el).select_by_value("3m")

    # Click Search
    try:
        search_btn = param_table.find_element(By.CSS_SELECTOR, "input#searchButton.submit[type='submit'][name='search']")
    except NoSuchElementException:
        search_btn = driver.find_element(By.CSS_SELECTOR, "input#searchButton")

    # Capture old tbody if present to wait for staleness
    old_tbody = None
    try:
        switch_into_table_iframe(driver)
        table_el = get_visible_tab_home(driver)
        if table_el:
            old_tbody = table_el.find_element(By.TAG_NAME, "tbody")
    except Exception:
        pass

    # Back to top to click Search
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", search_btn)
    time.sleep(0.2)
    try:
        search_btn.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", search_btn)

    if old_tbody is not None:
        try:
            WebDriverWait(driver, W).until(EC.staleness_of(old_tbody))
        except TimeoutException:
            pass

    # Results may live in an iframe
    switch_into_table_iframe(driver)
    WebDriverWait(driver, W).until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.tabHome tbody tr")))
    time.sleep(0.3)


# -------------------- pagination --------------------

def pagination_containers(driver) -> List:
    """Containers likely holding pager links; near table first, then global."""
    containers = []
    containers += driver.find_elements(
        By.XPATH,
        "//table[contains(@class,'tabHome')]/following::*[self::div or self::nav or self::td or self::p][.//a or .//button or .//span][position()<=5]"
    )
    containers += driver.find_elements(By.CSS_SELECTOR, ".pagination, nav, .pager, .pages, .dataTables_paginate")
    uniq = []
    seen = set()
    for c in containers:
        try:
            key = c.id
        except Exception:
            key = None
        if key and key not in seen and c.is_displayed():
            uniq.append(c)
            seen.add(key)
    return uniq


def parse_srchpage_arg(anchor) -> Optional[int]:
    """Extract page number from onclick="srchPage('N');return false;"."""
    try:
        onclick = anchor.get_attribute("onclick") or ""
        import re
        m = re.search(r"srchPage\(['\"](\d+)['\"]\)", onclick)
        if m:
            return int(m.group(1))
    except Exception:
        pass
    return None


def find_numeric_anchor_by_target(containers: List, target_num: int):
    """Find <a> with onclick srchPage('target_num') in any container."""
    for c in containers:
        anchors = c.find_elements(By.XPATH, ".//a[contains(@onclick, 'srchPage(')]")
        for a in anchors:
            if not a.is_displayed():
                continue
            n = parse_srchpage_arg(a)
            if n == target_num:
                return a
    return None


def find_next_anchor_with_span(containers: List):
    """
    Prefer a Next anchor with inner span arrow; else any anchor whose text contains Next/arrows
    and has onclick srchPage('N').
    """
    for c in containers:
        anchors = c.find_elements(By.XPATH, ".//a[contains(@onclick, 'srchPage(')]")
        for a in anchors:
            if not a.is_displayed():
                continue
            spans = a.find_elements(By.XPATH, ".//span[contains(., '›') or contains(., '»') or contains(., '→')]")
            for s in spans:
                if s.is_displayed():
                    return a, s
    for c in containers:
        anchors = c.find_elements(By.XPATH, ".//a[contains(@onclick, 'srchPage(')]")
        for a in anchors:
            if not a.is_displayed():
                continue
            txt = smart_text(a).upper()
            if "NEXT" in txt or "›" in txt or "»" in txt or "→" in txt:
                return a, None
    return None


def looks_disabled_for_real(el) -> bool:
    """Treat very few things as disabled; do NOT treat 'return false' as disabled."""
    try:
        cls = (el.get_attribute("class") or "").lower()
        aria_disabled = (el.get_attribute("aria-disabled") or "").lower()
        return ("disabled" in cls) or (aria_disabled in ("true", "1"))
    except Exception:
        return False


def click_and_wait_for_change(driver, el) -> bool:
    """Click and verify page change (tbody staleness or first-row signature change)."""
    table_el = get_visible_tab_home(driver)
    if not table_el:
        return False
    try:
        tbody_before = table_el.find_element(By.TAG_NAME, "tbody")
    except Exception:
        tbody_before = None
    sig_before = first_row_signature(table_el)

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.2)
    try:
        el.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", el)

    wait = WebDriverWait(driver, W)
    changed = False
    if tbody_before is not None:
        try:
            wait.until(EC.staleness_of(tbody_before))
            changed = True
        except TimeoutException:
            pass
    if not changed:
        try:
            wait.until(lambda d: first_row_signature(get_visible_tab_home(d)) != sig_before)
            changed = True
        except TimeoutException:
            changed = False
    return changed


# -------------------- IO --------------------

def write_csv(rows: List[BidRow], path: str):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Title", "Status", "End Date"])
        for r in rows:
            w.writerow([r.title, r.status, r.end_date])


def write_xlsx(rows: List[BidRow], path: str):
    try:
        from openpyxl import Workbook
    except ImportError:
        logging.warning("openpyxl not installed—skipping Excel export.")
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "WY Closed Bids"
    ws.append(["Title", "Status", "End Date"])
    for r in rows:
        ws.append([r.title, r.status, r.end_date])
    wb.save(path)


# -------------------- main --------------------

def main():
    driver = build_driver()
    all_rows: List[BidRow] = []
    try:
        logging.info("Opening Wyoming closed bids…")
        driver.get(START_URL)

        # Cookie banner first (if any)
        dismiss_cookies_if_present(driver)

        # Apply filter 'Last 3 months' then click Search
        apply_last_3_months_filter_and_search(driver)

        # Ensure we’re in the results context (iframe if used)
        switch_into_table_iframe(driver)
        WebDriverWait(driver, PAGE_LOAD_TIMEOUT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.tabHome"))
        )

        page = 1
        seen_sigs = set()

        while True:
            logging.info(f"Scraping page {page}…")
            rows = scrape_current_page_rows(driver)
            all_rows.extend(rows)

            sig = first_row_signature(get_visible_tab_home(driver))
            if sig in seen_sigs and rows:
                logging.info("Duplicate page signature detected—stopping.")
                break
            seen_sigs.add(sig)

            containers = pagination_containers(driver)
            if not containers:
                logging.info("No pagination container found—done.")
                break

            moved = False

            # (A) Click numeric link srchPage('page+1') if present
            next_anchor = find_numeric_anchor_by_target(containers, page + 1)
            if next_anchor and not looks_disabled_for_real(next_anchor):
                if click_and_wait_for_change(driver, next_anchor):
                    page += 1
                    moved = True
            if moved:
                continue

            # (B) Click Next (prefer inner span › if present)
            combo = find_next_anchor_with_span(containers)
            if combo:
                a, s = combo
                target = s if s else a
                if not looks_disabled_for_real(a):
                    if click_and_wait_for_change(driver, target):
                        page += 1
                        moved = True
            if moved:
                continue

            logging.info("No usable pagination control found—done.")
            break

        logging.info(f"Total rows scraped (unfiltered): {len(all_rows)}")

        # -------- FINAL FILTER: keep only rows whose Status is FINALIZED --------
        finalized_rows = [r for r in all_rows if (r.status or "").strip().upper() == "FINALIZED"]
        logging.info(f"Rows after FINALIZED filter: {len(finalized_rows)}")

        write_csv(finalized_rows, CSV_PATH)
        write_xlsx(finalized_rows, XLSX_PATH)
        logging.info(f"Wrote filtered CSV: {CSV_PATH}")
        logging.info(f"Wrote filtered XLSX: {XLSX_PATH}")
        print(f"Done. Final (filtered) files created:\n - {CSV_PATH}\n - {XLSX_PATH}")

    finally:
        time.sleep(0.8)
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        driver.quit()


if __name__ == "__main__":
    main()