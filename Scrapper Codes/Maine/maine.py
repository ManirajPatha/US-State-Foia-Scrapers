# maine.py — Selenium navigation + requests parsing for attachments
# Now with SAFE-PROGRESS SAVE: if Chrome is closed mid-run, a partial JSON is written.

import argparse
import json
import os
import re
import time
import zipfile
import shutil
import atexit
import signal
from datetime import datetime
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, ElementClickInterceptedException,
    WebDriverException
)

BASE_URL = "https://www.maine.gov/dafs/bbm/procurementservices/vendors/rfps/rfp-archives"

ATTACH_EXT = {
    ".pdf", ".doc", ".docx", ".rtf", ".txt",
    ".xls", ".xlsx", ".csv",
    ".ppt", ".pptx",
    ".zip",
    ".jpg", ".jpeg", ".png", ".gif", ".tif", ".tiff"
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
}

# -------------------- NEW: partial save machinery --------------------
ALL_ROWS = []          # accumulates rows across sections
OUT_JSON_PATH = None   # set in main()
SAVE_EVERY = 1         # save after each row (safe for your use-case)

def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

def save_json(rows, path):
    ensure_dir(os.path.dirname(os.path.abspath(path)) or ".")
    # Dedup by (RFP #, Title) before saving (keeps output clean on repeated saves)
    seen = set()
    dedup = []
    for r in rows:
        key = (r.get("RFP #", ""), r.get("Title", ""))
        if key in seen:
            continue
        seen.add(key)
        dedup.append(r)
    meta = {
        "_note": "Partial output may appear if the run was interrupted. Attachments already downloaded remain on disk.",
        "_saved_at": datetime.now().isoformat(timespec="seconds")
    }
    payload = {"meta": meta, "rows": dedup}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

def save_partial():
    if OUT_JSON_PATH:
        save_json(ALL_ROWS, OUT_JSON_PATH)

def register_exit_handlers():
    # Save on normal interpreter exit
    atexit.register(save_partial)
    # Save on common termination signals (Windows supports SIGINT & SIGBREAK; SIGTERM may also work)
    def _handler(signum, frame):
        save_partial()
        # re-raise default to actually stop program after saving
        raise SystemExit(0)
    for sig in [getattr(signal, "SIGINT", None),
                getattr(signal, "SIGTERM", None),
                getattr(signal, "SIGBREAK", None)]:
        if sig is not None:
            try:
                signal.signal(sig, _handler)
            except Exception:
                pass

def maybe_checkpoint():
    # Save every N rows (N=1 to capture progress ASAP)
    if OUT_JSON_PATH and (len(ALL_ROWS) % SAVE_EVERY == 0):
        save_partial()

# -------------------- utilities (unchanged logic) --------------------

def slugify(text: str, maxlen: int = 120) -> str:
    text = re.sub(r"[^\w\s-]", "", text, flags=re.UNICODE).strip()
    text = re.sub(r"[-\s]+", "-", text, flags=re.UNICODE)
    out = text[:maxlen] if text else "untitled"
    return out or "untitled"

def safe_filename_from_url(url: str) -> str:
    name = os.path.basename(urlparse(url).path) or "file"
    name = re.sub(r"[^\w\-. ]", "_", name)
    return name[:200] or "file"

def is_downloadable(href: str) -> bool:
    if not href:
        return False
    _, ext = os.path.splitext(urlparse(href).path.lower())
    return ext in ATTACH_EXT

def http_get(url: str, timeout: int = 90) -> requests.Response:
    for attempt in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=timeout, stream=True)
            r.raise_for_status()
            return r
        except Exception:
            if attempt == 2:
                raise
            time.sleep(1.2)

def zip_folder(src_dir: str, zip_path: str):
    ensure_dir(os.path.dirname(zip_path))
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(src_dir):
            for fn in files:
                fp = os.path.join(root, fn)
                arc = os.path.relpath(fp, src_dir)
                zf.write(fp, arc)

def new_driver(download_dir: str) -> webdriver.Chrome:
    ensure_dir(download_dir)
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": os.path.abspath(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()),
                              options=options)
    return driver

def wait_for(driver, by, locator, timeout=20):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, locator)))

def section_root(driver, section_text: str):
    # Try exact tag matches first
    heading = None
    for tag in ["h1", "h2", "h3", "h4", "strong", "span", "div"]:
        els = driver.find_elements(By.XPATH, f"//{tag}[normalize-space()='{section_text}']")
        if els:
            heading = els[0]
            break
    if heading is None:
        heading = driver.find_element(By.XPATH, f"//*[contains(normalize-space(), '{section_text}')]")
    table = heading.find_element(By.XPATH, "following::table[1]")
    container = table.find_element(By.XPATH, "./ancestor::*[self::div or self::section][1]")
    return container, table

def parse_row_cells(tr):
    tds = tr.find_elements(By.TAG_NAME, "td")
    if len(tds) < 9:
        return None

    # Title cell/link
    title_el = None
    try:
        title_el = tds[0].find_element(By.TAG_NAME, "a")
    except NoSuchElementException:
        pass
    title_text = (title_el.text.strip() if title_el else tds[0].text.strip())
    title_href = (title_el.get_attribute("href") if title_el else "")

    rfp_no = tds[1].text.strip()
    issuing_dept = tds[2].text.strip()
    date_posted = tds[3].text.strip()

    qa_links = []
    for a in tds[4].find_elements(By.TAG_NAME, "a"):
        href = a.get_attribute("href")
        qa_links.append({"text": a.text.strip(), "url": href})

    proposal_due = tds[5].text.strip()
    rfp_status = tds[6].text.strip()

    vendors_links = []
    as_ = tds[7].find_elements(By.TAG_NAME, "a")
    if as_:
        for a in as_:
            vendors_links.append({"text": a.text.strip(), "url": a.get_attribute("href")})
    else:
        txt = tds[7].text.strip()
        if txt:
            vendors_links.append({"text": txt, "url": ""})

    next_anticipated = tds[8].text.strip()

    return {
        "Title": title_text,
        "Title URL": title_href,
        "RFP #": rfp_no,
        "Issuing Department": issuing_dept,
        "Date Posted": date_posted,
        "Q&A/Amendments (JSON)": qa_links,
        "Proposal Due Date": proposal_due,
        "RFP Status": rfp_status,
        "Awarded Vendor(s) (JSON)": vendors_links,
        "Next Anticipated RFP Release": next_anticipated,
    }

# ----------- requests-based Title-page parsing (unchanged) -----------

def collect_downloads_from_title_requests(title_text: str, title_url: str, attachments_dir: str):
    """
    Fetch the Title page via requests, find all downloadable links, download them,
    and ZIP into Maine attachments/<slug>.zip
    Returns: (zip_path or "", downloaded_files[], source_urls[])
    """
    if not title_url:
        return "", [], []

    source_urls = []
    if is_downloadable(title_url):
        source_urls.append(title_url)
    else:
        try:
            resp = http_get(title_url)
            soup = BeautifulSoup(resp.text, "html.parser")
            for a in soup.select("a[href]"):
                href = urljoin(title_url, a.get("href"))
                if is_downloadable(href):
                    source_urls.append(href)
        except Exception:
            pass

    # Dedup
    seen = set()
    srcs = []
    for u in source_urls:
        if u not in seen:
            srcs.append(u)
            seen.add(u)

    if not srcs:
        return "", [], []

    slug = slugify(title_text)
    tmp_dir = os.path.join(attachments_dir, f"__tmp_{slug}")
    ensure_dir(tmp_dir)

    downloaded = []
    for u in srcs:
        try:
            r = http_get(u)
            fname = safe_filename_from_url(u)
            outp = os.path.join(tmp_dir, fname)
            with open(outp, "wb") as f:
                for chunk in r.iter_content(chunk_size=128 * 1024):
                    if chunk:
                        f.write(chunk)
            downloaded.append(outp)
        except Exception:
            continue

    if not downloaded:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return "", [], srcs

    zip_path = os.path.join(attachments_dir, f"{slug}.zip")
    zip_folder(tmp_dir, zip_path)
    shutil.rmtree(tmp_dir, ignore_errors=True)
    return zip_path, downloaded, srcs

# ----------- Section processing & pagination (minor: call checkpoint) -----------

def process_table_with_pagination(driver, section_name: str, attachments_dir: str, out_rows: list):
    print(f"[INFO] Processing section: {section_name}")
    container, table = section_root(driver, section_name)

    while True:
        # Current page rows
        tbody = table.find_element(By.TAG_NAME, "tbody") if table.find_elements(By.TAG_NAME, "tbody") else table
        rows = tbody.find_elements(By.TAG_NAME, "tr")
        first_row_marker = rows[0].text if rows else str(time.time())

        for r in rows:
            tds = r.find_elements(By.TAG_NAME, "td")
            if len(tds) < 2:
                continue
            data = parse_row_cells(r)
            if not data:
                continue
            if data["RFP Status"].strip().lower() != "awarded":
                continue

            title = data["Title"]
            title_url = data["Title URL"]

            # Download attachments via requests (no tab management)
            zip_path, downloaded, src_urls = collect_downloads_from_title_requests(
                title, title_url, attachments_dir
            )

            record = dict(data)
            record["attachments"] = {
                "zip_path": zip_path,
                "downloaded_files": downloaded,
                "source_urls": src_urls
            }
            out_rows.append(record)
            ALL_ROWS.append(record)    # <-- track globally for partial saves
            maybe_checkpoint()          # <-- NEW: checkpoint after each row

        # Try to click "Next" for this section
        next_btn = None
        for xpath in [
            ".//a[normalize-space()='Next' and not(contains(@class,'disabled'))]",
            ".//*[self::ul or self::div][contains(@class,'pagination')]//a[normalize-space()='Next' and not(contains(@class,'disabled'))]",
            "following::a[normalize-space()='Next' and not(contains(@class,'disabled'))][1]",
        ]:
            try:
                cand = container.find_element(By.XPATH, xpath)
                next_btn = cand
                break
            except NoSuchElementException:
                continue

        if next_btn:
            aria = next_btn.get_attribute("aria-disabled")
            if aria and aria.strip().lower() == "true":
                next_btn = None

        if not next_btn:
            break  # finished this section

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", next_btn)
        try:
            next_btn.click()
        except ElementClickInterceptedException:
            driver.execute_script("arguments[0].click();", next_btn)

        # Wait for page to change
        try:
            WebDriverWait(driver, 20).until(EC.staleness_of(rows[0]))
        except Exception:
            try:
                WebDriverWait(driver, 15).until(
                    lambda d: table.find_element(By.TAG_NAME, "tbody").find_elements(By.TAG_NAME, "tr")[0].text != first_row_marker
                )
            except Exception:
                pass

        # Re-locate container/table after DOM refresh
        container, table = section_root(driver, section_name)

# -------------------- main --------------------

def main():
    global OUT_JSON_PATH
    ap = argparse.ArgumentParser(description="Maine RFP Archives (Selenium UI + requests attachments) → JSON + ZIPs with partial-save")
    ap.add_argument("--url", default=BASE_URL)
    ap.add_argument("--out-json", default=f"maine_rfp_{datetime.now().strftime('%Y%m%d_%H%M')}.json")
    ap.add_argument("--attachments-dir", default="Maine attachments")
    args = ap.parse_args()

    OUT_JSON_PATH = os.path.abspath(args.out_json)
    attachments_dir = os.path.abspath(args.attachments_dir)
    ensure_dir(attachments_dir)

    # Register partial-save handlers
    register_exit_handlers()

    chrome_downloads = os.path.join(attachments_dir, "__chrome_dl")
    ensure_dir(chrome_downloads)

    driver = None
    try:
        driver = new_driver(download_dir=chrome_downloads)
        driver.get(args.url)

        section_rows = []

        # Section 1
        process_table_with_pagination(driver, "RFP Archives", attachments_dir, section_rows)

        # Reset & Section 2
        driver.get(args.url)
        process_table_with_pagination(driver, "Older Archives", attachments_dir, section_rows)

        # Final save (full)
        save_partial()
        print(f"[INFO] Rows saved: {len(ALL_ROWS)}")
        print(f"[INFO] JSON  → {OUT_JSON_PATH}")
        print(f'[INFO] ZIPs  → "{attachments_dir}"')

    except (WebDriverException, Exception) as e:
        # If the user closed the browser or any error occurs, write whatever we have
        print(f"[WARN] Interrupted: {e}")
        save_partial()
        print(f"[INFO] Partial JSON written to: {OUT_JSON_PATH}")
    finally:
        # Try to quit driver gracefully
        try:
            if driver is not None:
                driver.quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()
