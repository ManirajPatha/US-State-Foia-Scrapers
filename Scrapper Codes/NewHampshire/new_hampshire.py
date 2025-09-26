#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
NH Statewide Bids — Export Awarded → Excel (headless)

Flow:
1) Go to https://apps.das.nh.gov/bidscontracts/bids.aspx
2) Set "Status/Bid Results" = "Awarded"
3) Click Search (applies filter)
4) Click "Export to Excel"
5) Wait for the download to finish in --out folder

Usage:
  pip install selenium webdriver-manager
  python nh_awarded_export.py --out "C:/path/to/downloads"

Notes:
- Works headless (background).
- We do not parse the page; we just download the Excel the site generates.
"""

import argparse
import os
import time
from datetime import datetime
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

LIST_URL = "https://apps.das.nh.gov/bidscontracts/bids.aspx"


def make_driver(download_dir: Path) -> webdriver.Chrome:
    download_dir.mkdir(parents=True, exist_ok=True)

    chrome_prefs = {
        "download.default_directory": str(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }

    opts = Options()
    opts.add_experimental_option("prefs", chrome_prefs)
    opts.add_argument("--headless=new")      # run in background
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1400,1400")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36")

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opts)
    driver.set_page_load_timeout(60)

    # Allow downloads in headless via CDP
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": str(download_dir)
        })
    except Exception:
        # Some builds already allow it; ignore failures here
        pass

    return driver


def wait_for_new_download(download_dir: Path, before_files: set[str], timeout: int = 120) -> Path:
    """
    Wait for a new file (.xls/.xlsx/.csv) to appear and finish downloading
    (no .crdownload). Returns the final file path.
    """
    end = time.time() + timeout
    seen_path: Path | None = None

    while time.time() < end:
        current = set(os.listdir(download_dir))
        new_files = [f for f in current - before_files if not f.endswith(".crdownload")]
        # prioritize spreadsheet extensions
        new_files_sorted = sorted(new_files, key=lambda f: (not (f.lower().endswith((".xlsx", ".xls", ".csv"))), f))
        if new_files_sorted:
            seen_path = download_dir / new_files_sorted[0]
            # ensure file size stabilizes
            last_size = -1
            for _ in range(40):  # up to ~10 seconds stabilization
                size = seen_path.stat().st_size if seen_path.exists() else -1
                if size == last_size and size > 0:
                    return seen_path
                last_size = size
                time.sleep(0.25)
        time.sleep(0.25)

    raise TimeoutError("Timed out waiting for the Excel download to complete.")


def run(download_dir: Path) -> Path:
    driver = make_driver(download_dir)
    try:
        driver.get(LIST_URL)

        # 1) Set Status/Bid Results = Awarded
        status_select = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//select[option[normalize-space()='Awarded']]"))
        )
        Select(status_select).select_by_visible_text("Awarded")

        # 2) Click Search to apply filter (site generally updates results)
        search_btn = driver.find_element(
            By.XPATH, "//input[@type='submit' and translate(@value,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='SEARCH']"
        )
        search_btn.click()

        # Wait until the results table is present (not strictly required for export, but safer)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, "//table[.//th[normalize-space()='Description'] and .//th[normalize-space()='Bid #']]")
            )
        )

        # 3) Snapshot existing files
        before = set(os.listdir(download_dir))

        # 4) Click "Export to Excel"
        export_link = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(normalize-space(.), 'Export to Excel')]")
            )
        )
        export_link.click()

        # 5) Wait for the new file to appear and finish
        downloaded = wait_for_new_download(download_dir, before, timeout=180)

        # Optionally, rename to something stable with timestamp if you like:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        if not downloaded.name.lower().endswith((".xlsx", ".xls", ".csv")):
            # if site gives odd extension, keep it
            final_path = downloaded
        else:
            final_path = downloaded.with_stem(f"nh_awarded_export_{ts}")
            try:
                downloaded.rename(final_path)
            except Exception:
                # If rename fails (locked by AV), just keep original name
                final_path = downloaded

        return final_path

    finally:
        driver.quit()


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Export NH 'Awarded' bids to Excel (headless).")
    ap.add_argument("--out", default=".", help="Download folder (default: current directory).")
    args = ap.parse_args()
    out_dir = Path(args.out)
    result = run(out_dir)
    print(f"[OK] Downloaded → {result}")
