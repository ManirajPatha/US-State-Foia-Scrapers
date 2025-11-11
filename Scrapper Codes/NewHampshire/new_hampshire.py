#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
NH Statewide Bids — Export Awarded → JSON (headless)

Flow:
1) Go to https://apps.das.nh.gov/bidscontracts/bids.aspx
2) Set "Status/Bid Results" = "Awarded"
3) Click Search (applies filter)
4) Click "Export to Excel" (site generates full-results spreadsheet)
5) Wait for the download to finish in --out folder
6) Convert the downloaded spreadsheet to JSON (records) and save alongside it

Usage:
  pip install selenium webdriver-manager pandas openpyxl
  python new_hampshire.py --out "C:/path/to/downloads"

Notes:
- Works headless (background).
- JSON output contains one object per row (list[dict]).
"""

import argparse
import json
import os
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
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
    opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1400,1400")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
    )

    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()), options=opts
    )
    driver.set_page_load_timeout(60)

    # Allow downloads in headless via CDP (best-effort)
    try:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": str(download_dir)},
        )
    except Exception:
        pass

    return driver


def wait_for_new_download(download_dir: Path, before_files: set[str], timeout: int = 180) -> Path:
    """
    Wait for a new file (.xls/.xlsx/.csv) to appear and finish downloading
    (no .crdownload). Returns the final file path.
    """
    end = time.time() + timeout
    seen_path: Path | None = None

    while time.time() < end:
        current = set(os.listdir(download_dir))
        new_files = [
            f for f in current - before_files
            if not f.endswith(".crdownload")
        ]
        # Prefer spreadsheet-like extensions first
        new_files_sorted = sorted(
            new_files,
            key=lambda f: (not (f.lower().endswith((".xlsx", ".xls", ".csv"))), f),
        )
        if new_files_sorted:
            seen_path = download_dir / new_files_sorted[0]
            # ensure file size stabilizes
            last_size = -1
            for _ in range(40):  # up to ~10 seconds
                size = seen_path.stat().st_size if seen_path.exists() else -1
                if size == last_size and size > 0:
                    return seen_path
                last_size = size
                time.sleep(0.25)
        time.sleep(0.25)

    raise TimeoutError("Timed out waiting for the export download to complete.")


def export_awarded_to_spreadsheet(download_dir: Path) -> Path:
    driver = make_driver(download_dir)
    try:
        driver.get(LIST_URL)

        # 1) Set Status/Bid Results = Awarded
        status_select = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//select[option[normalize-space()='Awarded']]"))
        )
        Select(status_select).select_by_visible_text("Awarded")

        # 2) Click Search to apply filter
        search_btn = driver.find_element(
            By.XPATH, "//input[@type='submit' and translate(@value,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='SEARCH']"
        )
        search_btn.click()

        # Wait until the results table is present (safer before exporting)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, "//table[.//th[normalize-space()='Description'] and .//th[normalize-space()='Bid #']]")
            )
        )

        # 3) Snapshot existing files so we can detect the new one
        before = set(os.listdir(download_dir))

        # 4) Click "Export to Excel"
        export_link = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(normalize-space(.), 'Export to Excel')]"))
        )
        export_link.click()

        # 5) Wait for the new file to appear and finish
        downloaded = wait_for_new_download(download_dir, before, timeout=180)
        return downloaded

    finally:
        driver.quit()


def spreadsheet_to_json(spreadsheet_path: Path, out_dir: Path) -> Path:
    """
    Convert the exported Excel/CSV to JSON (records).
    Returns the JSON file path.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_path = out_dir / f"nh_awarded_{ts}.json"

    # Load into DataFrame
    ext = spreadsheet_path.suffix.lower()
    if ext == ".csv":
        df = pd.read_csv(spreadsheet_path, dtype=str)
    else:
        # .xlsx / .xls
        # openpyxl handles .xlsx; .xls may need xlrd (rare here)
        df = pd.read_excel(spreadsheet_path, dtype=str, engine="openpyxl")

    # Normalize: strip whitespace, replace NaNs with ""
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df = df.fillna("")

    # Convert to list-of-dicts
    records = df.to_dict(orient="records")

    # Write JSON
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    return json_path


def run_to_json(out_dir: Path) -> Path:
    spreadsheet = export_awarded_to_spreadsheet(out_dir)
    json_file = spreadsheet_to_json(spreadsheet, out_dir)
    return json_file


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Export NH 'Awarded' bids to JSON (headless).")
    ap.add_argument("--out", default=".", help="Output/download folder (default: current directory).")
    args = ap.parse_args()

    out_dir = Path(args.out)
    result_json = run_to_json(out_dir)
    print(f"[OK] JSON written → {result_json}")
