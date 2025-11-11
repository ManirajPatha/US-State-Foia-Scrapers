#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Tennessee RFP Opportunities scraper (robust, per-row unique dirs; Selenium + requests)

- Scrapes the table once.
- For each row: downloads ALL file attachments discoverable from each link.
- Per-row temp dir is ALWAYS UNIQUE (row-index prefixed) to prevent collisions when
  different rows have the same first-link text.
- ZIP filename is also made unique if a duplicate exists.
"""

import os
import re
import json
import time
import shutil
import zipfile
import argparse
import mimetypes
from urllib.parse import urlparse, urljoin
from pathlib import Path
from typing import List, Dict, Tuple, Optional

import requests
from requests.adapters import HTTPAdapter, Retry

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

TN_URL = "https://www.tn.gov/generalservices/procurement/central-procurement-office--cpo-/supplier-information/request-for-proposals--rfp--opportunities1.html"

ATTACH_EXTS = (".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".ppt", ".pptx")
ONCLICK_URL_RE = re.compile(r"""window\.open\(['"]([^'"]+)['"]""", re.I)

REQ_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
}

# ---------------- small helpers ---------------- #

def sanitize_name(name: str) -> str:
    if not name:
        return "file"
    cleaned = re.sub(r"[^\w\-.]+", "_", name.strip())
    return cleaned.strip("_") or "file"

def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)

def looks_like_file(url: str) -> bool:
    u = (url or "").lower()
    base = u.split("?", 1)[0]
    return any(base.endswith(ext) for ext in ATTACH_EXTS)

def unique_path(base_dir: Path, name: str) -> Path:
    """Ensure filename uniqueness: name, name (1), name (2), ..."""
    p = base_dir / name
    if not p.exists():
        return p
    stem, ext = os.path.splitext(name)
    i = 1
    while True:
        cand = base_dir / f"{stem} ({i}){ext}"
        if not cand.exists():
            return cand
        i += 1

def unique_zip_path(base_dir: Path, base_name_no_ext: str) -> Path:
    """Return a unique .zip path in base_dir using base_name_no_ext."""
    p = base_dir / f"{base_name_no_ext}.zip"
    if not p.exists():
        return p
    i = 1
    while True:
        cand = base_dir / f"{base_name_no_ext} ({i}).zip"
        if not cand.exists():
            return cand
        i += 1

def filename_from_headers(url: str, headers: Dict[str, str]) -> str:
    cd = headers.get("content-disposition") or headers.get("Content-Disposition")
    if cd:
        m = re.search(r'filename\*=UTF-8\'\'([^;]+)', cd)
        if m:
            name = requests.utils.unquote(m.group(1))
            return sanitize_name(name)
        m = re.search(r'filename="([^"]+)"', cd)
        if m:
            return sanitize_name(m.group(1))
        m = re.search(r'filename=([^;]+)', cd)
        if m:
            return sanitize_name(m.group(1).strip().strip('"\''))
    name = sanitize_name(os.path.basename(urlparse(url).path))
    if not os.path.splitext(name)[1]:
        ctype = headers.get("content-type") or headers.get("Content-Type") or ""
        ext = mimetypes.guess_extension(ctype.split(";")[0].strip())
        if ext:
            name = f"{name}{ext}"
    return name or "download.bin"

def build_requests_session_from_driver(driver: webdriver.Chrome) -> requests.Session:
    sess = requests.Session()
    sess.headers.update(REQ_HEADERS)
    retries = Retry(total=3, backoff_factor=0.5, status_forcelist=(429, 500, 502, 503, 504))
    adapter = HTTPAdapter(max_retries=retries, pool_connections=20, pool_maxsize=20)
    sess.mount("http://", adapter)
    sess.mount("https://", adapter)
    try:
        for c in driver.get_cookies():
            sess.cookies.set(name=c.get("name"), value=c.get("value"),
                             domain=c.get("domain"), path=c.get("path", "/"))
    except Exception:
        pass
    return sess

# ---------------- selenium helpers ---------------- #

def build_driver() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.page_load_strategy = "eager"  # faster; we still wait for table
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(25)
    return driver

def safe_nav(driver: webdriver.Chrome, url: str, wait_css: Optional[str] = None, max_wait: int = 15) -> None:
    try:
        driver.get(url)
    except TimeoutException:
        pass
    if wait_css:
        try:
            WebDriverWait(driver, max_wait).until(EC.presence_of_element_located((By.CSS_SELECTOR, wait_css)))
        except TimeoutException:
            pass

def open_new_tab(driver: webdriver.Chrome, url: str) -> str:
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    time.sleep(0.2)
    handle = driver.window_handles[-1]
    driver.switch_to.window(handle)
    return handle

def close_current_tab(driver: webdriver.Chrome) -> None:
    try:
        driver.close()
    except Exception:
        pass
    try:
        driver.switch_to.window(driver.window_handles[0])
    except Exception:
        pass

# ---------------- discovery of file URLs ---------------- #

def gather_file_urls_on_page(driver: webdriver.Chrome) -> List[str]:
    urls = set()
    # anchors
    try:
        for a in driver.find_elements(By.XPATH, "//a[@href]"):
            href = (a.get_attribute("href") or "").strip()
            if looks_like_file(href):
                urls.add(href)
            onclick = " ".join([
                a.get_attribute("onclick") or "",
                a.get_attribute("data-onclick") or "",
                a.get_attribute("ng-click") or "",
                a.get_attribute("data-action") or ""
            ])
            m = ONCLICK_URL_RE.search(onclick)
            if m and looks_like_file(m.group(1)):
                urls.add(m.group(1))
    except Exception:
        pass
    # embeds/iframes/src
    try:
        for tag in ("iframe", "embed", "source"):
            for el in driver.find_elements(By.TAG_NAME, tag):
                src = (el.get_attribute("src") or "").strip()
                if looks_like_file(src):
                    urls.add(src)
    except Exception:
        pass
    # deep anchors inside iframes
    try:
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for i in range(len(iframes)):
            try:
                driver.switch_to.frame(i)
                for a in driver.find_elements(By.XPATH, "//a[@href]"):
                    href = (a.get_attribute("href") or "").strip()
                    if looks_like_file(href):
                        urls.add(href)
                try:
                    frame_url = driver.execute_script("return window.location.href;")
                    if isinstance(frame_url, str) and looks_like_file(frame_url):
                        urls.add(frame_url)
                except Exception:
                    pass
            finally:
                driver.switch_to.default_content()
    except Exception:
        pass
    return list(urls)

# ---------------- scraping + downloading ---------------- #

def scrape_table_blueprint(driver: webdriver.Chrome) -> List[Dict]:
    print("[STEP] Load table page…")
    safe_nav(driver, TN_URL)
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//table//tr[td]")))
    rows = driver.find_elements(By.XPATH, "//table//tr[td]")
    print(f"[INFO] Rows detected: {len(rows)}")

    blueprint = []
    for r in rows:
        tds = r.find_elements(By.TAG_NAME, "td")
        if len(tds) < 3:
            continue
        links = []
        for a in tds[0].find_elements(By.XPATH, ".//a[@href]"):
            text = (a.text or "").strip()
            href = (a.get_attribute("href") or "").strip()
            if href:
                links.append({"text": text, "url": href})
        event_start_due = (tds[1].text or "").strip()
        event_name = (tds[2].text or "").strip()
        last_updated = (tds[3].text or "").strip() if len(tds) > 3 else ""
        blueprint.append({
            "doc_links": links,
            "event_start_response_due": event_start_due,
            "event_name": event_name,
            "last_updated": last_updated,
        })
    return blueprint

def download_file(session: requests.Session, base_url: str, file_url: str, out_dir: Path, timeout: int = 90) -> Optional[Path]:
    try:
        url = urljoin(base_url, file_url)
        with session.get(url, stream=True, allow_redirects=True, timeout=timeout) as resp:
            resp.raise_for_status()
            name = filename_from_headers(url, resp.headers)
            out_path = unique_path(out_dir, name)
            with open(out_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=1024 * 512):
                    if chunk:
                        f.write(chunk)
        return out_path
    except Exception:
        return None

def download_for_row(driver: webdriver.Chrome, row_idx: int, row_info: Dict, zip_parent: Path) -> Tuple[str, List[str]]:
    # ZIP/dir base name from first link text; prefix with row index to avoid collisions
    first_text = row_info["doc_links"][0]["text"] if row_info["doc_links"] else f"row_{row_idx+1}"
    zip_base = f"{row_idx+1:03d}_{sanitize_name(first_text)}"

    # UNIQUE per-row temp dir (prefix ensures no reuse even if same first-link text appears)
    row_tmp = zip_parent / f"__tmp_{zip_base}"
    ensure_dir(row_tmp)

    files_saved: List[str] = []
    session = build_requests_session_from_driver(driver)

    for link_idx, link in enumerate(row_info["doc_links"], start=1):
        href = link["url"]
        print(f"[ROW {row_idx+1}] Link {link_idx}/{len(row_info['doc_links'])}: {href}")

        # If direct file, pull via requests immediately
        if looks_like_file(href):
            saved = download_file(session, href, href, row_tmp)
            if saved:
                files_saved.append(Path(saved).name)

        # Open link page and harvest any file URLs there
        tab_handle = None
        try:
            tab_handle = open_new_tab(driver, href)
            time.sleep(1.0)  # allow redirect
            base_for_rel = driver.current_url

            # If the final URL itself is a file, download it
            if looks_like_file(base_for_rel):
                saved = download_file(session, base_for_rel, base_for_rel, row_tmp)
                if saved:
                    files_saved.append(Path(saved).name)
            else:
                urls = gather_file_urls_on_page(driver)

                # Also include anchors with download attribute
                try:
                    for a in driver.find_elements(By.XPATH, "//a[@download and @href]"):
                        cand = (a.get_attribute("href") or "").strip()
                        if cand:
                            urls.append(cand)
                except Exception:
                    pass

                # Deduplicate
                seen = set()
                unique_urls = []
                for u in urls:
                    if u not in seen:
                        seen.add(u)
                        unique_urls.append(u)

                for u in unique_urls:
                    saved = download_file(session, base_for_rel, u, row_tmp)
                    if saved:
                        files_saved.append(Path(saved).name)

        except Exception as e:
            print(f"  [WARN] Tab error: {e}")
        finally:
            if tab_handle:
                close_current_tab(driver)
        session = build_requests_session_from_driver(driver)

    # Make ZIP name unique too (if same base reappears later)
    zip_path = unique_zip_path(zip_parent, zip_base)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for f in row_tmp.glob("*"):
            if f.is_file():
                zf.write(f, arcname=f.name)

    # Return list of files actually included
    files_in_zip = sorted([f.name for f in row_tmp.glob("*") if f.is_file()])

    # Cleanup temp
    try:
        shutil.rmtree(row_tmp, ignore_errors=True)
    except Exception:
        pass

    return str(zip_path), files_in_zip

def main():
    parser = argparse.ArgumentParser(description="Scrape Tennessee RFP table and download attachments per row.")
    parser.add_argument("--out-json", default="tn_opportunities.json", help="Output JSON path.")
    parser.add_argument("--attachments-dir", default="Tennessee attachments", help="Parent folder for all ZIPs.")
    args = parser.parse_args()

    out_json = Path(args.out_json).resolve()
    zip_parent = Path(args.attachments_dir).resolve()
    ensure_dir(zip_parent)

    driver = build_driver()
    results = []
    try:
        blueprint = scrape_table_blueprint(driver)
        print(f"[INFO] Rows to process: {len(blueprint)}")

        for idx, row in enumerate(blueprint):
            print(f"\n[PROCESS] Row {idx+1}/{len(blueprint)} — {row.get('event_name','')}")
            if not row["doc_links"]:
                results.append({
                    "row_index": idx + 1,
                    "document_links": row["doc_links"],
                    "event_start_response_due": row["event_start_response_due"],
                    "event_name": row["event_name"],
                    "last_updated": row["last_updated"],
                    "zip_file": None,
                    "files_in_zip": [],
                })
                continue

            try:
                zip_path, files_in_zip = download_for_row(driver, idx, row, zip_parent)
            except Exception as e:
                print(f"[ERROR] Row {idx+1} failed: {e}")
                zip_path, files_in_zip = (None, [])

            results.append({
                "row_index": idx + 1,
                "document_links": row["doc_links"],
                "event_start_response_due": row["event_start_response_due"],
                "event_name": row["event_name"],
                "last_updated": row["last_updated"],
                "zip_file": zip_path,
                "files_in_zip": files_in_zip,
            })
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    print("\n[DONE] JSON:", out_json)
    print("[DONE] ZIPs folder:", zip_parent)
    print(f"[INFO] Rows scraped: {len(results)}")

if __name__ == "__main__":
    main()
