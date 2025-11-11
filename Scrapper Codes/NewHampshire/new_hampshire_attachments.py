# new_hampshire.py
# -*- coding: utf-8 -*-
"""
New Hampshire Awarded Bids — Selenium-only click-to-download (robust & complete)

What it does
------------
1) Opens https://apps.das.nh.gov/bidscontracts/bids.aspx
2) Sets "Status/Bid Results:" = "Awarded" → clicks Search
3) For EVERY row on EVERY page:
   - Clicks ALL links in:
       • Bid # column
       • Addendum column
       • Status/Bid Results column
     Handles:
       • href links
       • onclick=window.open('...')
       • javascript:__doPostBack(...)
       • Same-tab navigations (returns back)
       • New-tab popups (closes and returns)
   - Waits for the downloads to finish (no .crdownload)
   - Moves only the NEW files from Chrome's download dir into:
       Bid <number>/<Description>/         (Bid # + Addendum)
       Bid <number>/Award Notice/          (Status/Award)
   - Appends a JSON record with all row fields + saved file paths
4) After all pages, writes JSON and a ZIP of the folder tree.

Hardened per request
--------------------
A) Pagination goes strictly 1→2→3→…; after 10 it clicks “…” to reveal next block and continues.
   Uses JS clicks, waits for real page change, handles stale elements, retries.
B) Checkpoint JSON written after each row and also on close/CTRL+C.

CLI
---
pip install selenium webdriver-manager
python new_hampshire.py --out "C:/path/out" [--headless] [--dl-timeout 25]
"""

from __future__ import annotations
import argparse
import json
import os
import re
import shutil
import time
import signal
import sys
from pathlib import Path
from datetime import datetime
from typing import Any, Dict, List, Optional

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    ElementClickInterceptedException,
)
from webdriver_manager.chrome import ChromeDriverManager

BASE_URL = "https://apps.das.nh.gov/bidscontracts/bids.aspx"

# Table & pagination xpaths (stable selectors by visible headers/text)
TABLE_XP = "//table[.//th[normalize-space()='Description'] and .//th[normalize-space()='Bid #']]"
ROWS_XP  = TABLE_XP + "//tr[td]"
PAGER_XP = "//div[.//a[normalize-space()='>>'] or .//a[normalize-space()='...'] or .//a[normalize-space()='1'] or .//input]"

# File detection
FILE_EXT_RE  = re.compile(r"\.(pdf|doc|docx|xls|xlsx|zip|rtf|txt)(?:\?.*)?$", re.I)

# ---------------- Helpers ----------------
def sanitize(s: str, maxlen: int = 120) -> str:
    s = re.sub(r"[\\/:*?\"<>|]+", " ", (s or "").strip())
    s = re.sub(r"\s+", " ", s).strip(" .")
    return s[:maxlen] if len(s) > maxlen else s

def extract_bid_number(text: str) -> str:
    m = re.search(r"Bid\s*(\d+)", text or "", re.I)
    if m: return m.group(1)
    m2 = re.search(r"(\d+)", text or "")
    return m2.group(1) if m2 else (text or "UNKNOWN")

def now_tag() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def wait_present(drv, xp: str, t: int = 25):
    return WebDriverWait(drv, t).until(EC.presence_of_element_located((By.XPATH, xp)))

def make_driver(download_dir: Path, headless: bool) -> webdriver.Chrome:
    download_dir.mkdir(parents=True, exist_ok=True)
    prefs = {
        "download.default_directory": str(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "plugins.always_open_pdf_externally": True,  # force PDF to download
        "safebrowsing.enabled": True,
    }
    opts = Options()
    opts.add_experimental_option("prefs", prefs)
    opts.page_load_strategy = "eager"
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1400,1400")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--ignore-certificate-errors")
    # speed up list pages
    opts.add_argument("--blink-settings=imagesEnabled=false")

    drv = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opts)
    drv.set_page_load_timeout(60)
    return drv

# ---------------- Page actions ----------------
def set_awarded_and_search(drv):
    drv.get(BASE_URL)
    status = wait_present(drv, "//select[option[normalize-space()='Awarded']]")
    Select(status).select_by_visible_text("Awarded")
    drv.find_element(
        By.XPATH,
        "//input[@type='submit' and translate(@value,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='SEARCH']",
    ).click()
    wait_present(drv, TABLE_XP)

def get_row_count(drv) -> int:
    try:
        return len(drv.find_elements(By.XPATH, ROWS_XP))
    except Exception:
        return 0

def cell_xpath(i: int, j: int) -> str:
    return ROWS_XP + f"[{i}]/td[{j}]"

def safe_cell_text(drv, i: int, j: int) -> str:
    xp = cell_xpath(i, j)
    for _ in range(3):
        try:
            td = drv.find_element(By.XPATH, xp)
            # innerText handles hidden line breaks better than .text in Selenium
            return (drv.execute_script("return (arguments[0].innerText||'').trim();", td) or "").strip()
        except (StaleElementReferenceException, NoSuchElementException):
            time.sleep(0.08)
    return ""

def links_in_cell(drv, i: int, j: int):
    try:
        return drv.find_element(By.XPATH, cell_xpath(i, j)).find_elements(By.TAG_NAME, "a")
    except Exception:
        return []

# ---------------- Download pipeline (Selenium) ----------------
def list_downloaded_files(dd: Path) -> List[Path]:
    return [p for p in dd.glob("*") if p.is_file() and not p.name.endswith(".crdownload")]

def wait_downloads_idle(dd: Path, timeout: int = 30, settle: int = 2):
    """
    Wait until there are no .crdownload files and the file list is stable for `settle` seconds.
    """
    start = time.time()
    last_count = -1
    stable_since = time.time()
    while True:
        if time.time() - start > timeout:
            return
        cr = list(dd.glob("*.crdownload"))
        files = list_downloaded_files(dd)
        count = len(files)
        if cr:
            stable_since = time.time()
        else:
            if count != last_count:
                last_count = count
                stable_since = time.time()
            else:
                if time.time() - stable_since >= settle:
                    return
        time.sleep(0.25)

def move_new_files(before: List[Path], dd: Path, dest: Path) -> List[str]:
    before_set = {p.resolve() for p in before}
    moved: List[str] = []
    for p in list_downloaded_files(dd):
        if p.resolve() not in before_set:
            dest.mkdir(parents=True, exist_ok=True)
            target = dest / p.name
            stem, ext = target.stem, target.suffix
            k = 1
            while target.exists():
                target = dest / f"{stem}_{k}{ext}"
                k += 1
            try:
                shutil.move(str(p), str(target))
                moved.append(str(target))
            except Exception:
                try:
                    shutil.copy2(str(p), str(target))
                    p.unlink(missing_ok=True)
                    moved.append(str(target))
                except Exception:
                    pass
    return moved

def click_link_and_download(drv, link_el, dd: Path, per_file_timeout: int) -> None:
    """
    Click one link robustly:
      - supports href, onclick window.open, __doPostBack
      - handles new-tab or same-tab navigation
      - clicks inner file anchors on the opened page (if any)
      - returns to list if navigated
    """
    handles_before = set(drv.window_handles)
    url_before = drv.current_url

    # Dispatch a real JS click to honor onclick/postbacks
    drv.execute_script("""
        const e = arguments[0];
        e.dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true, view: window}));
    """, link_el)

    # Detect new tab or same-tab nav
    new_handle = None
    for _ in range(40):  # ~2.4s
        diff = set(drv.window_handles) - handles_before
        if diff:
            new_handle = list(diff)[0]
            break
        if drv.current_url != url_before:
            break
        time.sleep(0.06)

    def click_file_anchors_on_current_tab():
        anchors = drv.find_elements(By.XPATH, "//a[@href]")
        for a in anchors:
            try:
                href = a.get_attribute("href") or ""
                if FILE_EXT_RE.search(href):
                    drv.execute_script("arguments[0].scrollIntoView({block:'center'});", a)
                    a.click()
                    time.sleep(0.1)
            except Exception:
                pass

    try:
        if new_handle:
            # New tab path
            root = drv.current_window_handle
            drv.switch_to.window(new_handle)
            time.sleep(0.2)
            click_file_anchors_on_current_tab()
            time.sleep(0.2)
            try:
                drv.close()
            except Exception:
                pass
            drv.switch_to.window(root)
        else:
            # Same-tab OR direct download
            if drv.current_url != url_before:
                time.sleep(0.2)
                click_file_anchors_on_current_tab()
                time.sleep(0.2)
                # Go back to list page
                try:
                    drv.back()
                    WebDriverWait(drv, 25).until(EC.presence_of_element_located((By.XPATH, TABLE_XP)))
                except Exception:
                    pass
            # If still same URL, it was a direct download → just wait
    finally:
        wait_downloads_idle(dd, per_file_timeout, settle=2)

def process_links_in_cell(drv, i: int, j: int, dd: Path, dest_dir: Path, per_file_timeout: int) -> List[str]:
    links = links_in_cell(drv, i, j)
    saved_all: List[str] = []
    for a in links:
        before = list_downloaded_files(dd)
        try:
            click_link_and_download(drv, a, dd, per_file_timeout)
        except Exception:
            # Ensure we’re still on a valid handle
            try:
                drv.switch_to.window(drv.window_handles[0])
            except Exception:
                pass
        moved = move_new_files(before, dd, dest_dir)
        saved_all.extend(moved)
    return saved_all

# ---------------- Pagination (HARDENED) ----------------
def _pager_signature(drv) -> str:
    """Compact pager signature for change detection."""
    try:
        el = drv.find_element(By.XPATH, PAGER_XP)
        return drv.execute_script("return (arguments[0].innerText||'').replace(/\\s+/g,' ').trim();", el)
    except Exception:
        return ""

def _table_signature(drv) -> str:
    """Compact signature of the first data row to detect page flip."""
    try:
        row = drv.find_element(By.XPATH, ROWS_XP + "[1]")
        return drv.execute_script("return (arguments[0].innerText||'').replace(/\\s+/g,' ').trim();", row)
    except Exception:
        return ""

def _current_page_num(drv) -> Optional[int]:
    """Read current page number from the pager <input> or active non-link element."""
    # ASP.NET pager often renders an <input> with the current page
    try:
        for inp in drv.find_elements(By.XPATH, PAGER_XP + "//input"):
            v = (inp.get_attribute("value") or "").strip()
            if v.isdigit():
                return int(v)
    except Exception:
        pass
    # Fallback: find an element that's NOT a link but looks like the active page (contains digits)
    try:
        els = drv.find_elements(By.XPATH, PAGER_XP + "//*[not(self::a)]")
        for e in els:
            t = (e.text or "").strip()
            if t.isdigit():
                return int(t)
    except Exception:
        pass
    return None

def _scroll_pager_into_view(drv):
    try:
        el = drv.find_element(By.XPATH, PAGER_XP)
        drv.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    except Exception:
        pass

def _click_page_number(drv, num: int) -> bool:
    """Find anchor with exact number and JS-click it."""
    try:
        links = drv.find_elements(By.XPATH, PAGER_XP + f"//a[normalize-space()='{num}']")
        if not links:
            return False
        a = links[0]
        drv.execute_script("arguments[0].scrollIntoView({block:'center'});", a)
        # JS click is more reliable against overlays
        drv.execute_script("arguments[0].click();", a)
        return True
    except (StaleElementReferenceException, NoSuchElementException, ElementClickInterceptedException):
        return False

def _click_ellipsis(drv) -> bool:
    try:
        ell = drv.find_element(By.XPATH, PAGER_XP + "//a[normalize-space()='...']")
        drv.execute_script("arguments[0].scrollIntoView({block:'center'});", ell)
        drv.execute_script("arguments[0].click();", ell)
        return True
    except Exception:
        return False

def _wait_page_changed(drv, old_page: Optional[int], target_page: int, old_sig: str, timeout: float = 20.0) -> bool:
    """Wait until pager input shows target_page AND first-row signature changes (or pager signature changes)."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            cur = _current_page_num(drv)
            if cur == target_page:
                # also ensure content changed (avoid misreads)
                sig = _table_signature(drv)
                if sig and sig != old_sig:
                    return True
        except Exception:
            pass
        time.sleep(0.2)
    return False

def paginate_next(drv) -> bool:
    """
    Strictly click (current+1). If it's not visible in current 10-box block,
    click "..." to reveal the next block, then click (current+1).
    Returns True if moved; False if no further page.
    """
    _scroll_pager_into_view(drv)
    cur = _current_page_num(drv)
    if cur is None:
        # Try to normalize: click 1 if available
        if not _click_page_number(drv, 1):
            return False
        try:
            wait_present(drv, TABLE_XP)
        except Exception:
            pass
        cur = _current_page_num(drv) or 1

    target = cur + 1
    old_row_sig = _table_signature(drv)

    # Try a few times: click target if present; else reveal more with "..."
    for _ in range(10):
        # If target link is visible, click it
        if _click_page_number(drv, target):
            if _wait_page_changed(drv, cur, target, old_row_sig, timeout=25.0):
                return True
            # if didn't visually update, small retry (stale/slow server)
            time.sleep(0.5)
            if _current_page_num(drv) == target:
                return True
        # Not visible → click "..." to reveal next block
        if not _click_ellipsis(drv):
            # No more blocks; we’re likely at the last page
            return False
        # Small wait for pager to rerender
        try:
            wait_present(drv, TABLE_XP)
        except Exception:
            pass
        time.sleep(0.3)
        _scroll_pager_into_view(drv)
    return False

# ---------------- Main ----------------
def main():
    ap = argparse.ArgumentParser(description="NH Awarded — Selenium click-to-download (robust)")
    ap.add_argument("--out", required=True, help="Base output directory")
    ap.add_argument("--headless", action="store_true", help="Run Chrome in headless mode")
    ap.add_argument("--dl-timeout", type=int, default=25, help="Per-click download wait seconds")
    args = ap.parse_args()

    out_base = Path(args.out).expanduser().resolve()
    out_base.mkdir(parents=True, exist_ok=True)

    # Chrome’s actual download location (temp)
    chrome_dl = out_base / "_chrome_downloads"
    chrome_dl.mkdir(parents=True, exist_ok=True)

    # Final structured output
    root_out = out_base / "NH_Awarded_Downloads"
    root_out.mkdir(parents=True, exist_ok=True)

    # --- Checkpoint JSON setup ---
    ts = now_tag()  # stable final filename per run
    json_path_final = out_base / f"nh_awarded_{ts}.json"
    json_path_ckpt  = out_base / "nh_awarded_checkpoint.json"

    def flush_checkpoint(rows: List[Dict[str, Any]]):
        try:
            with open(json_path_ckpt, "w", encoding="utf-8") as f:
                json.dump(rows, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    drv = make_driver(chrome_dl, args.headless)
    rows_json: List[Dict[str, Any]] = []

    # Always write JSON on close/Ctrl+C
    def _graceful_exit(signum, frame):
        try:
            with open(json_path_final, "w", encoding="utf-8") as f:
                json.dump(rows_json, f, ensure_ascii=False, indent=2)
            flush_checkpoint(rows_json)
        except Exception:
            pass
        try:
            drv.quit()
        except Exception:
            pass
        sys.exit(0)

    for sig in (signal.SIGINT, signal.SIGTERM, getattr(signal, "SIGHUP", None)):
        if sig:
            try:
                signal.signal(sig, _graceful_exit)
            except Exception:
                pass

    try:
        set_awarded_and_search(drv)

        page_idx = 1
        total_rows = 0
        while True:
            n = get_row_count(drv)
            for i in range(1, n + 1):
                # Re-read each cell freshly to avoid stale references
                meta = {
                    "Description":        safe_cell_text(drv, i, 1),
                    "Bid #":              safe_cell_text(drv, i, 2),
                    "Attachments":        safe_cell_text(drv, i, 3),
                    "Addendum":           safe_cell_text(drv, i, 4),
                    "Closing Date":       safe_cell_text(drv, i, 5),
                    "Closing Time":       safe_cell_text(drv, i, 6),
                    "Status/Bid Results": safe_cell_text(drv, i, 7),
                    "Contact":            safe_cell_text(drv, i, 8),
                    "Commodity Category": safe_cell_text(drv, i, 9),
                }
                bid_num = extract_bid_number(meta.get("Bid #", ""))
                desc    = sanitize(meta.get("Description") or "No Description")

                bid_root = root_out / f"Bid {bid_num}"
                desc_dir = bid_root / desc
                award_dir= bid_root / "Award Notice"
                desc_dir.mkdir(parents=True, exist_ok=True)
                award_dir.mkdir(parents=True, exist_ok=True)

                # BID # + ADDENDUM → Description folder (click ALL links)
                moved_desc: List[str] = []
                moved_desc += process_links_in_cell(drv, i, 2, chrome_dl, desc_dir, args.dl_timeout)
                moved_desc += process_links_in_cell(drv, i, 4, chrome_dl, desc_dir, args.dl_timeout)

                # STATUS/AWARD → Award Notice folder (click ALL links)
                moved_award = process_links_in_cell(drv, i, 7, chrome_dl, award_dir, args.dl_timeout)

                row_json = dict(meta)
                row_json["BidNumberParsed"] = bid_num
                row_json["DownloadPaths"] = {
                    "BidAndAddendum": moved_desc,
                    "AwardNotice": moved_award,
                }
                rows_json.append(row_json)
                total_rows += 1

                # Flush checkpoint after every row
                flush_checkpoint(rows_json)

                if total_rows % 25 == 0 or i == 1:
                    print(f"[Progress] page {page_idx} row {i}/{n} (total {total_rows})")

            # go to next page (current+1; uses "…" when needed)
            moved = paginate_next(drv)
            if not moved:
                break
            page_idx += 1

        # Save FINAL JSON + ZIP
        with open(json_path_final, "w", encoding="utf-8") as f:
            json.dump(rows_json, f, ensure_ascii=False, indent=2)
        flush_checkpoint(rows_json)  # keep checkpoint in sync

        zip_base = out_base / f"NH_Awarded_{ts}"
        archive_path = shutil.make_archive(str(zip_base), "zip", root_out)

        print(f"[OK] Rows processed: {len(rows_json)}")
        print(f"[OK] JSON: {json_path_final}")
        print(f"[OK] Checkpoint JSON: {json_path_ckpt}")
        print(f"[OK] ZIP : {archive_path}")

    finally:
        # Leave chrome_dl if partial .crdownload present; else try cleanup if empty
        try:
            if not list(chrome_dl.glob("*.crdownload")) and not list(chrome_dl.glob("*")):
                chrome_dl.rmdir()
        except Exception:
            pass
        try:
            drv.quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()
