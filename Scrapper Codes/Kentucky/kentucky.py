# kentucky_awarded_rfp_to_excel.py
# Scrape Kentucky VSS "Published Solicitations" with filters:
# Show Me=All, Category=Any, Type=Request for Proposals, Status=Awarded
# Saves results to an Excel file. No DB/S3 used.

import os
import time
import logging
import tempfile
from typing import Optional, List, Dict, Tuple

import pandas as pd
from dateutil import parser

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

SOURCE_URL = "https://vss.ky.gov/vssprod-ext/Advantage4"

# ---------- small utils ----------

def make_driver(download_dir: str, headless: bool = False):
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1400,1000")
    opts.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    })
    return webdriver.Chrome(options=opts)

def parse_date_time(text: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
    if not text:
        return None, None
    s = text.strip()
    if not s or s == "-":
        return None, None
    try:
        dt = parser.parse(s, fuzzy=True)
        date_iso = dt.date().isoformat()
        t = dt.time()
        time_hhmm = f"{t.hour:02d}:{t.minute:02d}" if (":" in s.lower() or "am" in s.lower() or "pm" in s.lower()) else None
        return date_iso, time_hhmm
    except Exception:
        return None, None

def wait_for_grid(wait: WebDriverWait):
    # VSS pages often finish rendering grid cells with id="datacell"
    wait.until(EC.presence_of_all_elements_located((By.ID, "datacell")))
    time.sleep(0.5)

def set_dropdown_by_label(driver, wait, label_text: str, visible_text: str) -> bool:
    """
    Tries to set a <select> whose associated label reads `label_text` to `visible_text`.
    Returns True if it changed something, False otherwise.
    Works for the Search panel fields like Status, Type, Show Me, Category.
    """
    try:
        # Find an element that looks like a label/caption with the given text
        # Then locate a descendant/neighboring <select>.
        # Try several reasonable XPaths for robustness across skins.
        candidates = [
            f"//*[normalize-space(text())='{label_text}']",
            f"//label[normalize-space(.)='{label_text}']",
            f"//div[normalize-space(.)='{label_text}']",
        ]
        label_el = None
        for xp in candidates:
            try:
                label_el = driver.find_element(By.XPATH, xp)
                break
            except NoSuchElementException:
                continue
        if not label_el:
            return False

        # Nearest select (same container or a following sibling)
        select_xpath_variants = [
            ".//following::select[1]",
            "../following-sibling::*//select[1]",
            "../../following-sibling::*//select[1]",
            ".//ancestor::*[1]//select[1]",
        ]
        sel_el = None
        for sx in select_xpath_variants:
            try:
                sel_el = label_el.find_element(By.XPATH, sx)
                break
            except NoSuchElementException:
                continue
        if not sel_el:
            return False

        Select(sel_el).select_by_visible_text(visible_text)
        time.sleep(0.8)  # allow auto-filter refresh
        return True
    except Exception:
        return False

def expand_search_if_collapsed(driver, wait):
    # If "Search" section is collapsed, click it.
    try:
        # Look for a toggle or "Search" header
        hdr = driver.find_element(By.XPATH, "//*[contains(@class,'search') and (self::a or self::div) and contains(translate(., 'SEARCH', 'search'),'search')]")
        # If there is a "Show More" or collapsed state, click once
        try:
            # If Keyword input isn't visible, we assume it's collapsed and click the header
            driver.find_element(By.XPATH, "//input[@type='text' and @placeholder or //label[normalize-space(.)='Keyword Search']]")
        except NoSuchElementException:
            hdr.click()
            time.sleep(0.5)
    except Exception:
        # Best effort; continue either way
        pass

def get_cell_value(cells, idx) -> Optional[str]:
    # cells show "Label\nValue"—return Value part when available
    try:
        if idx < len(cells):
            lines = cells[idx].text.strip().split("\n")
            if len(lines) > 1:
                val = lines[1].strip()
                return val if val and val != "-" else None
    except Exception:
        pass
    return None

# ---------- main scrape ----------

def scrape_awarded_rfps(url: str, max_rows: Optional[int], headless: bool) -> pd.DataFrame:
    tmp_dir = tempfile.mkdtemp(prefix="ky_vss_")
    driver = make_driver(tmp_dir, headless=headless)
    wait = WebDriverWait(driver, 30)

    records: List[Dict] = []
    logging.info(f"Opening: {url}")
    print(f"Opening: {url}")

    try:
        driver.get(url)

        # Click the "View Published Solicitations" tile
        tile = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='View Published Solicitations']")))
        tile.click()

        # Ensure results grid is present
        wait_for_grid(wait)

        # Open/expand the Search panel if needed & set filters
        expand_search_if_collapsed(driver, wait)

        # Show Me = All (usually default, but set explicitly)
        set_dropdown_by_label(driver, wait, "Show Me", "All")

        # Category = (leave as Any / blank) — usually default; do nothing

        # Type = Request for Proposals
        ok_type = set_dropdown_by_label(driver, wait, "Type", "Request for Proposals")
        if not ok_type:
            logging.warning("Could not set Type to 'Request for Proposals' (continuing).")

        # Status = Awarded
        ok_status = set_dropdown_by_label(driver, wait, "Status", "Awarded")
        if not ok_status:
            logging.warning("Could not set Status to 'Awarded' (continuing).")

        # If there is an explicit Search/Apply button, try clicking it
        for xp in [
            "//button[normalize-space(.)='Search']",
            "//button[contains(.,'Apply')]",
            "//button[contains(.,'Filter')]",
        ]:
            try:
                btn = driver.find_element(By.XPATH, xp)
                if btn.is_enabled():
                    btn.click()
                    time.sleep(1.0)
                    break
            except NoSuchElementException:
                pass

        # Wait for the filtered grid to render
        wait_for_grid(wait)

        # Iterate visible rows. In this UI each row’s link is every 8th cell starting at 5th (index 4).
        collected = 0
        while True:
            try:
                cells = driver.find_elements(By.ID, "datacell")
                if not cells:
                    break
                start_index, step = 4, 8
                progressed = False

                for i in range(start_index, len(cells), step):
                    if max_rows and collected >= max_rows:
                        break

                    try:
                        link_cell = driver.find_elements(By.ID, "datacell")[i]  # re-fetch to avoid staleness
                        a = link_cell.find_element(By.TAG_NAME, "a")
                        notice_id = a.text.strip()
                        if not notice_id:
                            continue

                        a.click()
                        progressed = True

                        wait_for_grid(wait)
                        detail_cells = driver.find_elements(By.ID, "datacell")

                        # indexes are based on the common VSS layout used here
                        title = get_cell_value(detail_cells, 11)
                        office = get_cell_value(detail_cells, 0)
                        email = get_cell_value(detail_cells, 1)
                        phone = get_cell_value(detail_cells, 2)
                        raw_publish = get_cell_value(detail_cells, 3)
                        raw_deadline = get_cell_value(detail_cells, 4)
                        industry = get_cell_value(detail_cells, 7)

                        pub_date, _ = parse_date_time(raw_publish)
                        deadline_date, _ = parse_date_time(raw_deadline)

                        # attachments (names only; not downloading)
                        attach_names: List[str] = []
                        try:
                            # Try locate Attachments tab and collect anchor texts
                            # Many VSS skins have an info tab region
                            attach_tab = driver.find_element(By.XPATH, "//*[normalize-space(text())='Attachments']")
                            attach_tab.click()
                            time.sleep(0.6)
                            tables = driver.find_elements(By.TAG_NAME, "table")
                            if len(tables) > 1:
                                for a2 in tables[1].find_elements(By.TAG_NAME, "a"):
                                    t = a2.text.strip()
                                    if t:
                                        attach_names.append(t)
                        except Exception:
                            pass

                        rec = {
                            "notice_id": notice_id,
                            "title": title,
                            "publish_date": pub_date,
                            "proposal_deadline": deadline_date,
                            "office": office,
                            "email": email,
                            "phone_number": phone,
                            "industry": industry,
                            "attachments": "; ".join(attach_names) if attach_names else None,
                            "page_url": driver.current_url,
                            "source": "Kentucky",
                            "status": "Awarded",
                            "type": "Request for Proposals",
                        }
                        records.append(rec)
                        collected += 1
                        logging.info(f"{collected}: {notice_id} — {title}")

                        # Back to list
                        try:
                            back_btn = driver.find_element(By.XPATH, "//*[@id='page_header']//button")
                            back_btn.click()
                        except NoSuchElementException:
                            driver.back()

                        wait_for_grid(wait)

                    except (StaleElementReferenceException, NoSuchElementException, TimeoutException):
                        # try to recover to the list
                        try:
                            driver.back()
                            wait_for_grid(wait)
                        except Exception:
                            pass

                if max_rows and collected >= max_rows:
                    break
                if not progressed:
                    break

            except Exception as e:
                logging.warning(f"Loop issue: {e}")
                break

        return pd.DataFrame.from_records(records)

    finally:
        driver.quit()

# ---------- CLI ----------

def main():
    import argparse
    ap = argparse.ArgumentParser(description="Kentucky VSS — Awarded RFPs to Excel (no DB).")
    ap.add_argument("--url", default=SOURCE_URL, help="Base portal URL")
    ap.add_argument("--out", default=None, help="Excel output path (default: ./kentucky_awarded_rfps_YYYYmmdd_HHMM.xlsx)")
    ap.add_argument("--max-rows", type=int, default=None, help="Stop after N rows (optional)")
    ap.add_argument("--headless", action="store_true", help="Run Chrome headless")
    args = ap.parse_args()

    df = scrape_awarded_rfps(args.url, max_rows=args.max_rows, headless=args.headless)

    if df.empty:
        logging.info("No rows scraped for the selected filters.")
        return

    ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M")
    out_path = args.out or os.path.abspath(f"./kentucky_awarded_rfps_{ts}.xlsx")
    df.to_excel(out_path, index=False)
    logging.info(f"Excel written: {out_path} (rows: {len(df)})")
    print(out_path)

if __name__ == "__main__":
    main()
