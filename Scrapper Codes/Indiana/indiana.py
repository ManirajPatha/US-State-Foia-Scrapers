#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Indiana IDOA — Award Recommendations (Selenium)
Downloads all Award Date ZIP attachments across all pages
and writes a JSON of the two table columns + download info.

Usage:
  pip install selenium webdriver-manager requests
  python indiana_attachments.py \
    --out-json indiana_awards.json \
    --out-dir "Indiana Attachments" \
    [--headless]
"""

import argparse
import json
import os
import time
import re
import atexit
import signal
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional
from urllib.parse import urlparse, unquote

import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    NoSuchElementException,
    WebDriverException,
    ElementClickInterceptedException,
)

from webdriver_manager.chrome import ChromeDriverManager

BASE_URL = "https://www.in.gov/idoa/procurement/award-recommendations/"

# ==================== NEW: global state for checkpoint ====================
ALL_RECORDS: List[Dict] = []
OUT_JSON_PATH: Optional[str] = None

def save_json_progress():
    """Write whatever we have so far into the JSON file (atomic-ish)."""
    global OUT_JSON_PATH, ALL_RECORDS
    if not OUT_JSON_PATH:
        return
    try:
        tmp = OUT_JSON_PATH + ".partial"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(ALL_RECORDS, f, ensure_ascii=False, indent=2)
        # Replace target to ensure we always have the newest snapshot
        os.replace(tmp, OUT_JSON_PATH)
    except Exception:
        # We never raise here—this is a best-effort checkpoint.
        pass

def _graceful_shutdown_handler(signum, frame):
    # Make sure we persist whatever we have when user stops/close
    save_json_progress()
    # Exit immediately after saving
    raise SystemExit(0)

# Register autosave on normal interpreter exit and on common stop signals
atexit.register(save_json_progress)
signal.signal(signal.SIGINT, _graceful_shutdown_handler)
signal.signal(signal.SIGTERM, _graceful_shutdown_handler)
# ==========================================================================

def mk_driver(download_dir: str, headless: bool = False) -> webdriver.Chrome:
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    prefs = {
        "download.default_directory": os.path.abspath(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
    }
    opts.add_experimental_option("prefs", prefs)
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=opts)

def wait_table(driver, timeout=20):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located(
            (By.XPATH,
             "//table[.//th[normalize-space()='Award Date'] and .//th[normalize-space()='Event Information']]")
        )
    )

def current_page_id(driver) -> str:
    try:
        tbl = driver.find_element(
            By.XPATH,
            "//table[.//th[normalize-space()='Award Date'] and .//th[normalize-space()='Event Information']]"
        )
        first_text = tbl.text[:200]
        return str(hash(first_text))
    except Exception:
        return str(time.time())

def parse_filename_from_headers(resp: requests.Response, default: str) -> str:
    cd = resp.headers.get("Content-Disposition") or resp.headers.get("content-disposition")
    if cd:
        m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)"?', cd)
        if m:
            return unquote(m.group(1)).strip()
    path_name = os.path.basename(urlparse(resp.url).path)
    if path_name:
        return unquote(path_name)
    return default

def selenium_cookies_to_requests(driver, domain: Optional[str] = None) -> requests.Session:
    s = requests.Session()
    for c in driver.get_cookies():
        s.cookies.set(c["name"], c["value"], domain=c.get("domain"), path=c.get("path","/"))
    s.headers.update({
        "User-Agent": "Mozilla/5.0",
        "Referer": BASE_URL,
        "Accept": "*/*",
    })
    return s

def safe_click(driver, element):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
        time.sleep(0.1)
        element.click()
        return True
    except (ElementClickInterceptedException, WebDriverException):
        try:
            ActionChains(driver).move_to_element(element).click().perform()
            return True
        except Exception:
            return False

def collect_rows_on_page(driver) -> List[Dict]:
    table = wait_table(driver)
    rows_out: List[Dict] = []
    trs = table.find_elements(By.XPATH, ".//tbody/tr[td]")
    for tr in trs:
        try:
            date_cell = tr.find_element(By.XPATH, "./td[1]")
            try:
                date_a = date_cell.find_element(By.XPATH, ".//a[@href]")
            except NoSuchElementException:
                date_a = None
            award_text = date_cell.text.strip()
            award_href = date_a.get_attribute("href") if date_a else None

            info_cell = tr.find_element(By.XPATH, "./td[2]")
            try:
                ev_a = info_cell.find_element(By.XPATH, ".//a[@href]")
                event_title = ev_a.text.strip()
                event_url = ev_a.get_attribute("href")
            except NoSuchElementException:
                event_title = info_cell.text.strip()
                event_url = None

            rows_out.append({
                "row_el": tr,
                "award_date": award_text,
                "event_title": event_title,
                "event_url": event_url,
                "award_zip_url": award_href,
            })
        except StaleElementReferenceException:
            continue
    return rows_out

def download_zip_for_row(driver, row: Dict, out_dir: str) -> Dict:
    award_url = row.get("award_zip_url")
    result = {
        "download_ok": False,
        "zip_path": None,
        "zip_url": award_url,
        "error": None,
    }
    if award_url:
        try:
            sess = selenium_cookies_to_requests(driver, domain=urlparse(award_url).hostname)
            resp = sess.get(award_url, stream=True, timeout=60)
            resp.raise_for_status()
            guessed = parse_filename_from_headers(resp, default=f"award_{int(time.time()*1000)}.zip")
            if not guessed.lower().endswith(".zip"):
                ct = resp.headers.get("Content-Type","").lower()
                if "zip" in ct or guessed.find(".") == -1:
                    guessed = guessed.rsplit(".",1)[0] + ".zip"
            out_path = os.path.join(out_dir, guessed)
            base, ext = os.path.splitext(out_path)
            k = 1
            while os.path.exists(out_path):
                out_path = f"{base} ({k}){ext}"
                k += 1
            with open(out_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)
            result["download_ok"] = True
            result["zip_path"] = out_path
            return result
        except Exception as e:
            result["error"] = f"requests-fallback: {e}"

    try:
        tr = row["row_el"]
        date_link = tr.find_element(By.XPATH, "./td[1]//a[@href]")
        ok = safe_click(driver, date_link)
        if ok:
            time.sleep(2.0)
            result["download_ok"] = True
        else:
            result["error"] = "could not click award date link"
    except Exception as e:
        result["error"] = f"selenium-click: {e}"

    return result

def goto_next_page(driver, timeout=15) -> bool:
    sig_before = current_page_id(driver)
    try:
        next_link = driver.find_element(By.XPATH, "//a[normalize-space()='Next' and @href]")
    except NoSuchElementException:
        return False
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", next_link)
    time.sleep(0.1)
    if not safe_click(driver, next_link):
        return False
    try:
        WebDriverWait(driver, timeout).until(lambda d: current_page_id(d) != sig_before)
        wait_table(driver, timeout=timeout)
        return True
    except TimeoutException:
        return False

def scrape_all(args):
    global OUT_JSON_PATH, ALL_RECORDS
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    OUT_JSON_PATH = args.out_json  # NEW: set for checkpoint writer

    driver = mk_driver(str(out_dir), headless=args.headless)
    try:
        driver.get(args.base_url)
        wait_table(driver)

        seen_page_signatures = set()
        page_index = 1

        while True:
            sig = current_page_id(driver)
            if sig in seen_page_signatures:
                break
            seen_page_signatures.add(sig)

            rows = collect_rows_on_page(driver)
            for row in rows:
                # Try/catch each row so a single failure still saves progress
                try:
                    dl = download_zip_for_row(driver, row, str(out_dir))
                except WebDriverException as e:
                    # Browser closed or crashed mid-row
                    dl = {"download_ok": False, "zip_path": None, "zip_url": row.get("award_zip_url"), "error": f"webdriver: {e}"}

                rec = {
                    "page_index": page_index,
                    "award_date": row.get("award_date"),
                    "event_title": row.get("event_title"),
                    "event_url": row.get("event_url"),
                    "award_zip_url": dl.get("zip_url"),
                    "zip_path": dl.get("zip_path"),
                    "download_ok": dl.get("download_ok"),
                    "error": dl.get("error"),
                    "scraped_at": datetime.now().isoformat(),
                }
                ALL_RECORDS.append(rec)

                # ==================== NEW: checkpoint after every row ====================
                save_json_progress()
                # =======================================================================

            # Try to go to the next page. If not possible, break.
            try:
                advanced = goto_next_page(driver)
            except WebDriverException:
                # If window got closed, we still have JSON saved so just stop.
                break
            if not advanced:
                break
            page_index += 1

    except WebDriverException:
        # If the window was closed abruptly, we just drop out—progress is saved.
        pass
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    # Final write (also covered by atexit, but do it explicitly)
    save_json_progress()
    print(f"[OK] JSON written (progress-safe): {OUT_JSON_PATH}")
    print(f"[OK] ZIPs folder : {Path(args.out_dir).resolve()}")

def main():
    ap = argparse.ArgumentParser(description="Indiana IDOA Award Recommendations — download all Award Date ZIPs + JSON.")
    ap.add_argument("--base-url", default=BASE_URL, help="Starting URL (default: Indiana Award Recommendations).")
    ap.add_argument("--out-json", default=f"indiana_awards_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                    help="Output JSON filename.")
    ap.add_argument("--out-dir", default="Indiana Attachments", help='Folder to store all ZIP files (default: "Indiana Attachments").')
    ap.add_argument("--headless", action="store_true", help="Run Chrome in headless mode.")
    args = ap.parse_args()

    scrape_all(args)

if __name__ == "__main__":
    main()
