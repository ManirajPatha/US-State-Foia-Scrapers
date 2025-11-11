# -*- coding: utf-8 -*-
"""
OHCA Procurement – Selenium downloader (column-ordered, in-tab navigation)

Changes in this version
-----------------------
- Strict column order: downloads Procurement Opportunity first, then Amendments.
- For top-level table links:
    * If link is a direct attachment → download directly (no new tab).
    * Else, open the link in the SAME TAB and crawl that page for attachments.
      (Crawler still opens new tabs only when sub-pages spawn them.)
- Everything else (zipping, JSON summary, skip-on-error) unchanged.

Usage
-----
pip install selenium webdriver-manager requests python-slugify
python oklahoma_selenium.py --out-json "C:/path/out.json" --out-dir "C:/path/Oklahoma"
"""

import os
import re
import json
import time
import zipfile
import argparse
from urllib.parse import urlparse

import requests
from slugify import slugify

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

BASE_URL = "https://oklahoma.gov/ohca/about/procurement.html"
OK_DOMAIN = "oklahoma.gov"

ATTACH_EXTS = (
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv",
    ".ppt", ".pptx", ".zip", ".txt", ".rtf", ".msg",
    ".jpg", ".jpeg", ".png"
)

WAIT_SHORT = 10
WAIT_LONG = 25


def log(msg): print(time.strftime("%H:%M:%S"), msg, flush=True)
def ensure_dir(p): os.makedirs(p, exist_ok=True); return p
def sanitize_filename(name: str) -> str:
    name = name.strip().replace(":", " -")
    return re.sub(r"[\\/*?\"<>|]", "_", name)


def init_driver(download_dir: str) -> webdriver.Chrome:
    prefs = {
        "download.default_directory": os.path.abspath(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }
    opts = webdriver.ChromeOptions()
    opts.add_experimental_option("prefs", prefs)
    opts.add_argument("--start-maximized")  # visible browser
    service = ChromeService(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=opts)


def wait_for_table(driver):
    WebDriverWait(driver, WAIT_LONG).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "table"))
    )


def find_procurement_table(driver):
    tables = driver.find_elements(By.CSS_SELECTOR, "table")
    for tbl in tables:
        header_cells = tbl.find_elements(By.CSS_SELECTOR, "thead th")
        if not header_cells:
            header_cells = tbl.find_elements(By.CSS_SELECTOR, "tr:first-child th, tr:first-child td")
        headers = " | ".join(h.text.strip().lower() for h in header_cells)
        if "requisition" in headers and "number" in headers:
            return tbl
    raise RuntimeError("Could not locate the procurement table on the page.")


def collect_rows(driver, table_el):
    header_map = []
    header_cells = table_el.find_elements(By.CSS_SELECTOR, "thead th")
    if not header_cells:
        header_cells = table_el.find_elements(By.CSS_SELECTOR, "tr:first-child th, tr:first-child td")
    for h in header_cells:
        header_map.append(h.text.strip().lower())

    expected = {
        "requisition number": "Requisition Number",
        "procurement opportunity": "Procurement Opportunity",
        "amendments": "Amendments",
        "status": "Status",
        "closing date": "Closing Date",
        "award date": "Award Date",
        "total annual contract value": "Total Annual Contract Value",
        "awardee(s)": "Awardee(s)",
        "awardee": "Awardee(s)",
    }

    body = table_el.find_element(By.TAG_NAME, "tbody") if table_el.find_elements(By.TAG_NAME, "tbody") else table_el
    out = []
    trs = body.find_elements(By.CSS_SELECTOR, "tr")
    log(f"Found {len(trs)} table rows.")
    for tr in trs:
        tds = tr.find_elements(By.CSS_SELECTOR, "td")
        if not tds: continue
        rec = {
            "Requisition Number": "",
            "Procurement Opportunity": {"text": "", "links": []},
            "Amendments": {"text": "", "links": []},
            "Status": "", "Closing Date": "", "Award Date": "",
            "Total Annual Contract Value": "", "Awardee(s)": "",
        }
        for idx, td in enumerate(tds):
            if idx >= len(header_map): continue
            norm = None
            col_raw = header_map[idx]
            if col_raw in expected: norm = expected[col_raw]
            else:
                for k, v in expected.items():
                    if k in col_raw: norm = v; break

            text = td.text.strip()
            links = [a.get_attribute("href") for a in td.find_elements(By.CSS_SELECTOR, "a[href]")]
            links = [u for u in links if u]

            if norm == "Procurement Opportunity":
                rec["Procurement Opportunity"]["text"] = text
                rec["Procurement Opportunity"]["links"] = links
            elif norm == "Amendments":
                rec["Amendments"]["text"] = text
                rec["Amendments"]["links"] = links
            elif norm:
                rec[norm] = text

        if any(v for v in rec.values()):
            out.append(rec)
    return out


def is_attachment_url(href: str) -> bool:
    if not href: return False
    path = urlparse(href).path.lower()
    return any(path.endswith(ext) for ext in ATTACH_EXTS)


def requests_download(url, dest_path, referer=None, timeout=60) -> bool:
    headers = {"User-Agent": "Mozilla/5.0"}
    if referer: headers["Referer"] = referer
    try:
        with requests.get(url, headers=headers, stream=True, timeout=timeout) as r:
            r.raise_for_status()
            cd = r.headers.get("content-disposition", "")
            if "filename=" in cd:
                fn = cd.split("filename=")[-1].strip('";\' ')
                dest_path = os.path.join(os.path.dirname(dest_path), sanitize_filename(fn))
            base, ext = os.path.splitext(dest_path)
            k = 1
            while os.path.exists(dest_path):
                dest_path = f"{base} ({k}){ext}"; k += 1
            with open(dest_path, "wb") as f:
                for chunk in r.iter_content(1 << 15):
                    if chunk: f.write(chunk)
        return True
    except Exception:
        return False


def wait_for_new_download(download_dir, before_files, timeout=90) -> str:
    start = time.time()
    while time.time() - start < timeout:
        now = set(os.listdir(download_dir))
        new_files = list(now - before_files)
        complete = [f for f in new_files if not f.endswith(".crdownload")]
        if complete: return complete[0]
        time.sleep(1.0)
    return ""


def click_download_in_browser(driver, url, download_dir) -> str:
    """Fallback when direct request fails but link might trigger a browser download."""
    before = set(os.listdir(download_dir))
    current = driver.current_url
    try:
        driver.execute_script("window.open(arguments[0], '_self');", url)
    except WebDriverException:
        driver.get(url)
    fname = wait_for_new_download(download_dir, before, timeout=120)
    # return to previous page if we navigated away
    try:
        if driver.current_url != current:
            driver.back()
    except Exception:
        pass
    return fname


def open_new_tab(driver, url):
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    driver.switch_to.window(driver.window_handles[-1])
    WebDriverWait(driver, WAIT_LONG).until(EC.presence_of_element_located((By.TAG_NAME, "body")))


def crawl_page_for_attachments(driver, page_url, save_dir, json_files, visited, depth=0, max_depth=2):
    """On the CURRENT PAGE, download any direct attachments; follow same-site listing pages up to max_depth."""
    if page_url in visited: return
    visited.add(page_url)

    links = driver.find_elements(By.CSS_SELECTOR, "a[href]")
    log(f"  [depth {depth}] Links found: {len(links)} @ {page_url}")

    # First: direct attachments
    for a in links:
        href = a.get_attribute("href") or ""
        if not href or not is_attachment_url(href): continue
        fn_base = sanitize_filename(os.path.basename(urlparse(href).path)) or "downloaded_file"
        dest_path = os.path.join(save_dir, fn_base)
        base, ext = os.path.splitext(dest_path)
        k = 1
        while os.path.exists(dest_path):
            dest_path = f"{base} ({k}){ext}"; k += 1

        ok = requests_download(href, dest_path, referer=page_url)
        if not ok:
            # fallback: try to make the browser download it
            fname = click_download_in_browser(driver, href, os.path.dirname(save_dir))
            if fname:
                fp = os.path.join(os.path.dirname(save_dir), fname)
                json_files.append(fp)
                log(f"    ✓ Downloaded by click: {fname}")
            else:
                log(f"    ✗ Failed to download: {href}")
        else:
            json_files.append(dest_path)
            log(f"    ✓ Downloaded via HTTP: {os.path.basename(dest_path)}")

    if depth >= max_depth: return

    # Second: recurse into internal pages that might list more attachments
    for a in links:
        href = a.get_attribute("href") or ""
        if not href or is_attachment_url(href): continue
        netloc = urlparse(href).netloc
        if netloc and OK_DOMAIN not in netloc: continue

        current = driver.current_window_handle
        try:
            # Some listing links may open a NEW TAB automatically; handle both cases.
            prev_handles = set(driver.window_handles)
            a.click()
            time.sleep(1.0)
            new_handles = set(driver.window_handles)
            opened_new = list(new_handles - prev_handles)

            if opened_new:
                driver.switch_to.window(opened_new[0])
                WebDriverWait(driver, WAIT_SHORT).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                sub_url = driver.current_url
                crawl_page_for_attachments(driver, sub_url, save_dir, json_files, visited, depth + 1, max_depth)
                driver.close()
                driver.switch_to.window(current)
            else:
                # navigated in same tab
                WebDriverWait(driver, WAIT_SHORT).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                sub_url = driver.current_url
                crawl_page_for_attachments(driver, sub_url, save_dir, json_files, visited, depth + 1, max_depth)
                driver.back()
                WebDriverWait(driver, WAIT_SHORT).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        except Exception as e:
            log(f"    (skip subpage) {href} → {e}")
            try:
                # try to return if navigation happened
                driver.switch_to.window(current)
            except Exception:
                pass


def zip_dir(src_dir, zip_path):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(src_dir):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, os.path.dirname(src_dir))
                z.write(full, rel)


def process_link_list(driver, links, row_dir, downloaded_files, errors, depth):
    """Helper that enforces the click/open rules you requested."""
    for link in links:
        try:
            if is_attachment_url(link):
                # Direct file → download without opening a tab
                fn = sanitize_filename(os.path.basename(urlparse(link).path)) or "file"
                dest = os.path.join(row_dir, fn)
                ok = requests_download(link, dest, referer=BASE_URL)
                if not ok:
                    # last resort: let Chrome try
                    fname = click_download_in_browser(driver, link, os.path.dirname(row_dir))
                    if fname:
                        downloaded_files.append(os.path.join(os.path.dirname(row_dir), fname))
                        log(f"    ✓ Direct-by-click: {fname}")
                    else:
                        log(f"    ✗ Direct download failed: {link}")
                else:
                    downloaded_files.append(dest)
                    log(f"    ✓ Direct HTTP: {os.path.basename(dest)}")
            else:
                # Not a file → open IN THE SAME TAB and crawl
                driver.get(link)
                time.sleep(1.2)
                visited = set()
                crawl_page_for_attachments(driver, driver.current_url, row_dir, downloaded_files, visited, depth=0, max_depth=depth)
                # Return to procurement table for next link
                driver.get(BASE_URL)
                wait_for_table(driver)
        except Exception as e:
            err = f"Link failed: {link} ({e})"
            log("  " + err)
            errors.append(err)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-json", required=True, help="Path to write JSON summary.")
    ap.add_argument("--out-dir", required=True, help="Working/output root directory.")
    ap.add_argument("--depth", type=int, default=2, help="Max crawl depth for nested attachment pages.")
    args = ap.parse_args()

    root = os.path.abspath(args.out_dir)
    work_dir = ensure_dir(os.path.join(root, "work"))
    zips_dir = ensure_dir(os.path.join(root, "Oklahoma Attachments"))
    downloads_dir = ensure_dir(os.path.join(root, "chrome_downloads"))

    driver = init_driver(downloads_dir)

    summary = {
        "source_url": BASE_URL,
        "scraped_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "rows": [],
    }

    try:
        log(f"Opening {BASE_URL}")
        driver.get(BASE_URL)
        wait_for_table(driver)
        table = find_procurement_table(driver)
        rows = collect_rows(driver, table)

        log(f"Processing {len(rows)} rows...")
        for idx, row in enumerate(rows, 1):
            req_no = row.get("Requisition Number", "").strip() or f"row_{idx}"
            folder_name = sanitize_filename(req_no)
            row_dir = ensure_dir(os.path.join(work_dir, folder_name))
            downloaded_files, errors = [], []

            # ==== STRICT ORDER ====
            # 1) Procurement Opportunity links
            po_links = list(dict.fromkeys(row["Procurement Opportunity"]["links"]))  # de-dupe keep order
            if po_links:
                log(f"[{req_no}] Procurement Opportunity: {len(po_links)} link(s)")
                process_link_list(driver, po_links, row_dir, downloaded_files, errors, depth=args.depth)

            # 2) Amendments links
            am_links = list(dict.fromkeys(row["Amendments"]["links"]))
            if am_links:
                log(f"[{req_no}] Amendments: {len(am_links)} link(s)")
                process_link_list(driver, am_links, row_dir, downloaded_files, errors, depth=args.depth)

            # Zip the row folder (even if empty, for consistency)
            zip_path = os.path.join(zips_dir, f"{folder_name}.zip")
            if os.path.exists(zip_path): os.remove(zip_path)
            zip_dir(row_dir, zip_path)
            log(f"[{req_no}] Zipped → {zip_path}")

            summary["rows"].append({
                "Requisition Number": row.get("Requisition Number", ""),
                "Procurement Opportunity": row.get("Procurement Opportunity", {}),
                "Amendments": row.get("Amendments", {}),
                "Status": row.get("Status", ""),
                "Closing Date": row.get("Closing Date", ""),
                "Award Date": row.get("Award Date", ""),
                "Total Annual Contract Value": row.get("Total Annual Contract Value", ""),
                "Awardee(s)": row.get("Awardee(s)", ""),
                "download_folder": row_dir,
                "zip_path": zip_path,
                "attachments_downloaded": len(downloaded_files),
                "downloaded_files": downloaded_files,
                "errors": errors,
            })

        with open(args.out_json, "w", encoding="utf-8") as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)
        log(f"✅ JSON written → {args.out_json}")
        log(f"✅ ZIPs folder → {zips_dir}")

    finally:
        try: driver.quit()
        except Exception: pass


if __name__ == "__main__":
    main()
