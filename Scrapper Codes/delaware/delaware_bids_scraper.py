#!/usr/bin/env python3
"""
Delaware MMP 'Recently Closed' scraper (visible Firefox + downloads)

- Opens https://mmp.delaware.gov/Bids/
- Clicks the 'Recently Closed' button (id=btnClosed)
- Iterates all pages (20 rows/page)
- For each data row:
  * Reads Contract Number, Title, Agency Code
  * Opens the modal and:
      - expands all 'Load More…' rows in the attachments table
      - downloads all attachments into:
          ./Delaware Scrapped attachments/<sanitized contract title>/
- Writes ./output_delaware.xlsx with columns:
  Contract Number | Contract Title | Agency Code

Requires: Python 3.10+, Selenium 4.x, Firefox + geckodriver
"""

import os
import re
import time
import logging
from pathlib import Path
from urllib.parse import urljoin, urlparse, unquote

import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver import FirefoxOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

URL = "https://mmp.delaware.gov/Bids/"

# ── Locators ───────────────────────────────────────────────────────────────────
BTN_CLOSED = (By.ID, "btnClosed")  # relaxed: id is stable; 'active' class is transient
TABLE = (By.CSS_SELECTOR, "table#jqGridBids.table.table-condensed.table-hover.table-bordered.ui-jqgrid-btable")
ROWS_CSS = "#jqGridBids > tbody > tr[role='row']"
TITLE_ANCHOR = (By.CSS_SELECTOR, "td.jqGridCellWrap a[href]")
CELL_CONTRACT = (By.CSS_SELECTOR, "td[role='gridcell'][aria-describedby='jqGridBids_ContractNumber']")
CELL_AGENCY = (By.CSS_SELECTOR, "td[role='gridcell'][aria-describedby='jqGridBids_AgencyCode']")
AGENCY_SPAN = (By.CSS_SELECTOR, "span.ui-jqgrid-cell-wrapper")
NEXT_BTN = (By.CSS_SELECTOR, "td#next_jqg1.btn.btn-xs.ui-pg-button[role='button']")

# Modal + attachments
MODAL = (By.CSS_SELECTOR, "div#dynamicDialogInnerHtml.modal-content.panel-info")
ATTACH_TABLE = (By.CSS_SELECTOR, "div#dynamicDialogInnerHtml table.table.table-bordered.table-sm")
ATTACH_TBODY_ROWS = (By.CSS_SELECTOR, "div#dynamicDialogInnerHtml table.table.table-bordered.table-sm tbody tr")
ATTACH_LINKS = (By.CSS_SELECTOR, "div#dynamicDialogInnerHtml table.table.table-bordered.table-sm tbody a[href]")

# ── Output ─────────────────────────────────────────────────────────────────────
ROOT_ATTACH_DIR = Path.cwd() / "Delaware Scrapped attachments"
EXCEL_PATH = Path.cwd() / "output_delaware.xlsx"

# ── Behavior knobs ─────────────────────────────────────────────────────────────
PAGE_LOAD_TIMEOUT = 60
WAIT_TIMEOUT = 25
SCROLL_PAUSE = 0.15
REQUEST_TIMEOUT = 60
LOAD_MORE_WAIT = 10  # seconds to wait for rows to increase after clicking "Load More..."


def setup_logger():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


def sanitize_name(name: str, max_len: int = 120) -> str:
    name = (name or "").strip()
    name = re.sub(r'[\\/:*?"<>|]', " ", name)
    name = re.sub(r"\s+", " ", name)
    name = name[:max_len].rstrip()
    return name or "Untitled"


def visible_scroll(driver, pixels=800):
    driver.execute_script(f"window.scrollBy(0,{int(pixels)});")
    time.sleep(SCROLL_PAUSE)


def build_requests_session_from_selenium(driver) -> requests.Session:
    s = requests.Session()
    for c in driver.get_cookies():
        s.cookies.set(c["name"], c["value"], domain=c.get("domain"), path=c.get("path", "/"))
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/124 Safari/537.36"
    })
    return s


def best_filename_from_response(resp: requests.Response, fallback_url: str) -> str:
    cd = resp.headers.get("Content-Disposition", "")
    if "filename=" in cd:
        fname = cd.split("filename=")[-1].strip().strip(";").strip('"')
        if fname:
            return sanitize_name(unquote(fname))
    path = unquote(urlparse(fallback_url).path)
    base = os.path.basename(path) or "download"
    return sanitize_name(base)


def element_text_safe(el):
    try:
        return el.text.strip()
    except Exception:
        return ""


def safe_scroll_into_view(driver, locator_or_element) -> bool:
    """Scrolls element into view; withstands one staleness hiccup."""
    try:
        el = driver.find_element(*locator_or_element) if isinstance(locator_or_element, tuple) else locator_or_element
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        return True
    except StaleElementReferenceException:
        try:
            if isinstance(locator_or_element, tuple):
                el = driver.find_element(*locator_or_element)
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                return True
        except Exception:
            pass
    except Exception:
        pass
    return False


def close_modal(driver, wait):
    candidates = [
        (By.CSS_SELECTOR, "div#dynamicDialogInnerHtml button.close"),
        (By.CSS_SELECTOR, "div#dynamicDialogInnerHtml button[data-dismiss='modal']"),
        (By.XPATH, "//div[@id='dynamicDialogInnerHtml']//button[normalize-space()='Close']"),
        (By.XPATH, "//div[@id='dynamicDialogInnerHtml']//button[contains(., 'Close')]"),
    ]
    for by, sel in candidates:
        try:
            btn = driver.find_element(by, sel)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            btn.click()
            wait.until(EC.invisibility_of_element_located(MODAL))
            return
        except Exception:
            pass
    try:
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        WebDriverWait(driver, 5).until(EC.invisibility_of_element_located(MODAL))
    except Exception:
        try:
            driver.execute_script("document.querySelector('.modal').click();")
            WebDriverWait(driver, 5).until(EC.invisibility_of_element_located(MODAL))
        except Exception:
            logging.warning("Modal did not close gracefully; continuing.")


def next_enabled(driver) -> bool:
    try:
        next_el = driver.find_element(*NEXT_BTN)
    except Exception:
        return False
    cls = (next_el.get_attribute("class") or "").lower()
    aria = (next_el.get_attribute("aria-disabled") or "").lower()
    return ("ui-state-disabled" not in cls) and (aria != "true")


def collect_rows(driver):
    """
    Return only real data rows: rows that contain a clickable title anchor.
    This avoids jqGrid utility/placeholder rows that cause 'no title link' logs.
    """
    all_rows = driver.find_elements(By.CSS_SELECTOR, ROWS_CSS)
    data_rows = []
    for r in all_rows:
        try:
            r.find_element(*TITLE_ANCHOR)
            data_rows.append(r)
        except Exception:
            pass
    return data_rows


def expand_all_modal_attachments(driver, wait) -> None:
    """
    Inside the open modal, repeatedly click the 'Load More…' row (if present)
    until no more rows are added. The 'Load More…' is typically the LAST <tr>,
    with a <td> (or <a> inside it) whose text is 'Load More...' and/or whose
    onclick is ReloadBidDetailBidDocumentsList(5).
    """
    while True:
        try:
            # Re-find table & rows each loop to avoid staleness
            table = driver.find_element(*ATTACH_TABLE)
            rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
            if not rows:
                return

            last_row = rows[-1]
            # Gather text across all cells to check for "Load More..."
            cells = last_row.find_elements(By.CSS_SELECTOR, "td")
            last_text = " ".join(element_text_safe(td) for td in cells).lower()
            onclick_hit = False
            click_target = None

            # Prefer a td with the known onclick
            for td in cells:
                onclick = (td.get_attribute("onclick") or "")
                if "ReloadBidDetailBidDocumentsList" in onclick:
                    onclick_hit = True
                    click_target = td
                    break

            # If not found via onclick, fall back to "Load More" text, try <a> first, then row
            if not onclick_hit and "load more" in last_text:
                try:
                    click_target = last_row.find_element(By.CSS_SELECTOR, "a[href]")
                except Exception:
                    click_target = last_row

            if not click_target:
                # No 'Load More…' at the end → fully expanded
                return

            # Click and wait for row count to increase
            prev_count = len(rows)
            safe_scroll_into_view(driver, click_target)
            driver.execute_script("arguments[0].click();", click_target)

            # Wait for more rows to load (or timeout if none come)
            try:
                WebDriverWait(driver, LOAD_MORE_WAIT).until(
                    lambda d: len(d.find_element(*ATTACH_TABLE)
                                 .find_elements(By.CSS_SELECTOR, "tbody tr")) > prev_count
                )
            except TimeoutException:
                # If it didn't increase, assume we're done to avoid an infinite loop
                return

            # Tiny buffer for DOM settle
            time.sleep(0.2)

        except Exception:
            # Any unexpected issue: exit the expander gracefully
            return


def main():
    setup_logger()
    ROOT_ATTACH_DIR.mkdir(exist_ok=True)

    opts = FirefoxOptions()
    # Visible browser (not headless)
    opts.set_preference("pdfjs.disabled", True)

    driver = webdriver.Firefox(options=opts)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    driver.maximize_window()
    wait = WebDriverWait(driver, WAIT_TIMEOUT)

    data_rows_out = []
    s = None

    try:
        logging.info("Opening page…")
        driver.get(URL)
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        time.sleep(0.5)

        # Little motion so you can see it work
        visible_scroll(driver, 400)
        visible_scroll(driver, -300)

        logging.info("Clicking 'Recently Closed'…")
        closed_btn = wait.until(EC.presence_of_element_located(BTN_CLOSED))
        safe_scroll_into_view(driver, closed_btn)
        time.sleep(0.2)
        driver.execute_script("arguments[0].click();", closed_btn)

        logging.info("Waiting for bids table to render…")
        wait.until(EC.presence_of_element_located(TABLE))
        wait.until(lambda d: len(collect_rows(d)) > 0)

        if s is None:
            s = build_requests_session_from_selenium(driver)

        page_number = 1
        while True:
            rows = collect_rows(driver)
            logging.info(f"Page {page_number}: found {len(rows)} data rows")

            try:
                first_cn = element_text_safe(rows[0].find_element(*CELL_CONTRACT)) if rows else ""
            except Exception:
                first_cn = ""

            # Iterate by index; re-fetch rows each loop to avoid stale references
            for idx in range(len(rows)):
                try:
                    rows = collect_rows(driver)
                    row = rows[idx]
                except IndexError:
                    break

                safe_scroll_into_view(driver, row)
                time.sleep(0.05)

                # Freshly locate per-row elements to avoid staleness
                try:
                    title_a = row.find_element(*TITLE_ANCHOR)
                except StaleElementReferenceException:
                    rows = collect_rows(driver)
                    row = rows[idx]
                    title_a = row.find_element(*TITLE_ANCHOR)

                title = element_text_safe(title_a)

                try:
                    contract_num = element_text_safe(row.find_element(*CELL_CONTRACT))
                except Exception:
                    contract_num = ""

                try:
                    agency_cell = row.find_element(*CELL_AGENCY)
                    try:
                        agency_code = element_text_safe(agency_cell.find_element(*AGENCY_SPAN)) or element_text_safe(agency_cell)
                    except Exception:
                        agency_code = element_text_safe(agency_cell)
                except Exception:
                    agency_code = ""

                logging.info(f"  Row {idx+1}: {contract_num} | {title} | {agency_code}")

                # Open modal (JS click is resilient)
                try:
                    safe_scroll_into_view(driver, title_a)
                    driver.execute_script("arguments[0].click();", title_a)
                except StaleElementReferenceException:
                    rows = collect_rows(driver)
                    row = rows[idx]
                    title_a = row.find_element(*TITLE_ANCHOR)
                    driver.execute_script("arguments[0].click();", title_a)

                # Wait for modal
                wait.until(EC.visibility_of_element_located(MODAL))
                time.sleep(0.2)

                # ---- NEW: expand all 'Load More…' rows before collecting links ----
                try:
                    driver.find_element(*ATTACH_TABLE)  # ensure table exists
                    expand_all_modal_attachments(driver, wait)
                except Exception:
                    pass  # no attachment table is okay

                # Download attachments (filter out any 'Load More…' pseudo-links)
                attach_dir = ROOT_ATTACH_DIR / sanitize_name(title or contract_num or f"row_{idx+1}")
                attach_dir.mkdir(parents=True, exist_ok=True)

                try:
                    links = driver.find_elements(*ATTACH_LINKS)
                    # Filter: ignore javascript pseudo-links or ones with 'load more' text
                    real_links = []
                    for a in links:
                        txt = element_text_safe(a).lower()
                        href = a.get_attribute("href") or ""
                        onclick = (a.get_attribute("onclick") or "")
                        if "load more" in txt:
                            continue
                        if "ReloadBidDetailBidDocumentsList" in onclick:
                            continue
                        if href.startswith("javascript:"):
                            continue
                        real_links.append(a)

                    logging.info(f"    Attachments found: {len(real_links)}")
                    for k, a in enumerate(real_links, start=1):
                        href = a.get_attribute("href")
                        if not href:
                            continue
                        file_url = urljoin(driver.current_url, href)
                        try:
                            r = s.get(file_url, stream=True, timeout=REQUEST_TIMEOUT)
                            r.raise_for_status()
                            fname = best_filename_from_response(r, file_url)
                            target = attach_dir / fname
                            base, ext = os.path.splitext(fname)
                            counter = 2
                            while target.exists():
                                target = attach_dir / f"{base} ({counter}){ext}"
                                counter += 1
                            with open(target, "wb") as f:
                                for chunk in r.iter_content(chunk_size=1 << 14):
                                    if chunk:
                                        f.write(chunk)
                            logging.info(f"      ✓ Saved: {target.name}")
                        except Exception as e:
                            logging.warning(f"      ✗ Failed to download {file_url}: {e}")
                except Exception:
                    logging.info("    No attachment table detected in this modal.")

                # Close modal and record row
                close_modal(driver, wait)

                data_rows_out.append({
                    "Contract Number": contract_num,
                    "Contract Title": title,
                    "Agency Code": agency_code,
                })

            # Pagination
            if not next_enabled(driver):
                logging.info("No further pages (Next disabled).")
                break

            old_first = first_cn
            next_el = driver.find_element(*NEXT_BTN)
            safe_scroll_into_view(driver, next_el)
            time.sleep(0.1)
            next_el.click()

            # Wait until first row changes (confirm page flip)
            def first_row_changed(drv):
                try:
                    new_rows = collect_rows(drv)
                    if not new_rows:
                        return False
                    new_cn = element_text_safe(new_rows[0].find_element(*CELL_CONTRACT))
                    return new_cn != old_first
                except Exception:
                    return False

            try:
                WebDriverWait(driver, WAIT_TIMEOUT).until(first_row_changed)
            except Exception:
                time.sleep(1.0)

            page_number += 1

        # Excel output
        logging.info(f"Writing Excel → {EXCEL_PATH.name}")
        df = pd.DataFrame(data_rows_out, columns=["Contract Number", "Contract Title", "Agency Code"])
        df.to_excel(EXCEL_PATH, index=False)
        logging.info("Done.")
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(0.6)

    finally:
        time.sleep(1.2)  # keep the browser visible briefly at the end
        driver.quit()


if __name__ == "__main__":
    main()