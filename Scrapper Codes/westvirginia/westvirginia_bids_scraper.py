# wv_vss_step3_clean.py
import time
import csv
import re
import logging

from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
    NoSuchElementException,
    WebDriverException,
)

# --- CONFIG ---
START_URL = "https://prd311.wvoasis.gov/PRDVSS1X1ERP/Advantage4"

CSV_PATH_RAW = "wv_vss_results.csv"
CSV_PATH_CLEAN = "wv_vss_results_clean.csv"

# Table/pagination locators (broader & more resilient across Advantage4 skins)
TABLE_GUESS_LOCATORS = [
    # Alaska-style (exact)
    (By.CSS_SELECTOR, 'div.css-jtdyb table#vsspageVVSSX10019gridView1group1cardGridgrid1'),
    # Generic Advantage4 grid tables
    (By.CSS_SELECTOR, 'table[role="grid"]'),
    (By.CSS_SELECTOR, 'div[role="region"] table'),
]
TBODY_LOCATOR = (By.CSS_SELECTOR, "tbody")
ROW_LOCATOR = (By.CSS_SELECTOR, "tbody > tr")
NEXT_BUTTON_LOCATOR = (By.CSS_SELECTOR, 'button[aria-label="Next"]')

# Search button: try name/aria, then generic "Search" aria-label
SEARCH_BUTTON_LOCATORS = [
    (By.CSS_SELECTOR, 'button[name*="gridView1.Search"][aria-label="Search"]'),
    (By.CSS_SELECTOR, 'button[aria-label="Search"]'),
    (By.XPATH, '//button[@aria-label="Search" or normalize-space()="Search"]'),
]

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


def build_driver() -> webdriver.Firefox:
    service = Service()  # relies on geckodriver on PATH
    opts = Options()
    # opts.add_argument("-headless")  # uncomment to run headless
    driver = webdriver.Firefox(service=service, options=opts)
    driver.maximize_window()
    driver.set_page_load_timeout(60)
    return driver


def wait_for_page_ready(driver, timeout=40):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )


def js_scroll_center(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)


def safe_click(driver, locator, timeout=25):
    wait = WebDriverWait(driver, timeout)
    el = wait.until(EC.presence_of_element_located(locator))
    js_scroll_center(driver, el)
    wait.until(EC.element_to_be_clickable(locator))
    try:
        el.click()
    except (ElementClickInterceptedException, StaleElementReferenceException, WebDriverException):
        el = driver.find_element(*locator)
        driver.execute_script("arguments[0].click();", el)
    return el


def click_view_published_solicitations(driver, timeout=35):
    # The WV site still uses this label per state guidance.
    locator = (
        By.CSS_SELECTOR,
        'div[title="View Published Solicitations"][aria-label="View Published Solicitations"]'
    )
    safe_click(driver, locator, timeout=timeout)
    logging.info("Clicked 'View Published Solicitations'.")


def expand_show_more(driver, timeout=20):
    wait = WebDriverWait(driver, timeout)
    btn_locator = (By.CSS_SELECTOR, 'button[aria-label="Please Enter or Space to expand Show More"]')
    try:
        btn = wait.until(EC.presence_of_element_located(btn_locator))
        js_scroll_center(driver, btn)
        wait.until(EC.element_to_be_clickable(btn_locator))
        try:
            btn.click()
        except (ElementClickInterceptedException, StaleElementReferenceException):
            btn = driver.find_element(*btn_locator)
            driver.execute_script("arguments[0].click();", btn)
        logging.info("Expanded 'Show More' filters.")
        time.sleep(0.4)
    except TimeoutException:
        logging.info("No 'Show More' button (already expanded or different skin).")


def _select_by_visible_text(select_el, wanted_texts):
    sel = Select(select_el)
    options = [o.text.strip() for o in sel.options]
    for wanted in wanted_texts:
        for opt in options:
            if opt.lower() == wanted.lower():
                sel.select_by_visible_text(opt)
                return True
    return False


def set_status_awarded(driver, timeout=25):
    wait = WebDriverWait(driver, timeout)
    # Try Alaska-style name first, then fallback by aria-label/title
    candidates = [
        (By.CSS_SELECTOR, 'select[name*=".SO_STA"][aria-label="Status"]'),
        (By.CSS_SELECTOR, 'select[aria-label="Status"]'),
        (By.XPATH, '//label[contains(.,"Status")]/following::select[1]'),
    ]
    select_el = None
    for loc in candidates:
        try:
            select_el = wait.until(EC.presence_of_element_located(loc))
            break
        except TimeoutException:
            continue
    if not select_el:
        logging.info("Status selector not found; skipping (site may default to All/Open).")
        return

    js_scroll_center(driver, select_el)
    # Prefer exact value when present; else fall back to visible text
    try:
        Select(select_el).select_by_value("A")  # Awarded = 'A' in many Advantage4 sites
    except Exception:
        ok = _select_by_visible_text(select_el, ["Awarded", "Closed", "Award", "Active - Awarded"])
        if not ok:
            logging.info("Could not set Status to Awarded by text; leaving default.")
            return
    logging.info("Set Status = Awarded.")


def set_show_me_all(driver, timeout=25):
    wait = WebDriverWait(driver, timeout)
    candidates = [
        (By.CSS_SELECTOR, 'select[name*=".SHOW_TXT"][aria-label="Show Me"]'),
        (By.CSS_SELECTOR, 'select[aria-label="Show Me"]'),
        (By.XPATH, '//label[contains(.,"Show Me")]/following::select[1]'),
    ]
    select_el = None
    for loc in candidates:
        try:
            select_el = wait.until(EC.presence_of_element_located(loc))
            break
        except TimeoutException:
            continue
    if not select_el:
        logging.info("'Show Me' selector not found; skipping.")
        return

    js_scroll_center(driver, select_el)
    try:
        # Alaska used value "1" for All
        Select(select_el).select_by_value("1")
    except Exception:
        ok = _select_by_visible_text(select_el, ["All", "Show All"])
        if not ok:
            logging.info("Could not set 'Show Me' to All; leaving default.")
            return
    logging.info("Set Show Me = All.")


def click_search(driver, timeout=30):
    last_err = None
    for loc in SEARCH_BUTTON_LOCATORS:
        try:
            safe_click(driver, loc, timeout=timeout)
            logging.info("Clicked Search.")
            return
        except Exception as e:
            last_err = e
            continue
    raise TimeoutException(f"Search button not found/clickable: {last_err}")


def find_results_table(driver, timeout=35):
    wait = WebDriverWait(driver, timeout)
    table = None
    for loc in TABLE_GUESS_LOCATORS:
        try:
            table = wait.until(EC.presence_of_element_located(loc))
            # ensure it has a TBODY and at least 1 row (or we keep trying)
            tbody = table.find_element(*TBODY_LOCATOR)
            rows = tbody.find_elements(*ROW_LOCATOR)
            if rows:
                js_scroll_center(driver, table)
                logging.info("Results table located.")
                return table
        except Exception:
            continue
    # last attempt: any table with rows
    tables = driver.find_elements(By.TAG_NAME, 'table')
    for t in tables:
        try:
            tbody = t.find_element(*TBODY_LOCATOR)
            rows = tbody.find_elements(*ROW_LOCATOR)
            if rows:
                js_scroll_center(driver, t)
                logging.info("Results table located (fallback).")
                return t
        except Exception:
            continue
    raise TimeoutException("Could not locate a populated results table.")


def extract_text_safe(el):
    try:
        txt = el.get_attribute("aria-label") or el.text
        return (txt or "").strip()
    except Exception:
        return ""


def sanitize_type(value: str) -> str:
    if value is None:
        return ""
    return re.sub(r"^\s*null[\s:,-]*", "", str(value), flags=re.IGNORECASE).strip()


def scrape_rows_from_current_page(driver, table):
    results = []
    tbody = table.find_element(*TBODY_LOCATOR)
    rows = tbody.find_elements(*ROW_LOCATOR)
    logging.info(f"Found {len(rows)} row elements on this page.")

    for i, row in enumerate(rows, start=1):
        try:
            js_scroll_center(driver, row)
            time.sleep(0.03)

            # Preferred selectors (Advantage4 data-qa)
            def q(sel):
                try:
                    return extract_text_safe(row.find_element(By.CSS_SELECTOR, sel))
                except NoSuchElementException:
                    return ""

            description = q('[data-qa*=".DOC_DSCR"]')
            department  = q('[data-qa*=".DeptBuyr.DEPT_NM"]')
            solnum      = q('a[data-qa*=".DOC_REF"], a.css-xv6zqn[data-qa*=".DOC_REF"]')
            bid_type    = sanitize_type(q('[data-qa*=".DOC_CD_CONCAT"]'))

            # If those aren't present, fallback to column order
            if not any([description, department, solnum, bid_type]):
                tds = row.find_elements(By.CSS_SELECTOR, "td")
                if tds:
                    # best guess: common order -> [sol#, description, dept, type] or similar
                    cand = [extract_text_safe(td) for td in tds]
                    # try to pick plausible fields
                    solnum = solnum or (cand[0] if cand else "")
                    description = description or (cand[1] if len(cand) > 1 else "")
                    department  = department or (cand[2] if len(cand) > 2 else "")
                    bid_type    = bid_type or sanitize_type(cand[3] if len(cand) > 3 else "")

            if any([description, department, solnum, bid_type]):
                results.append({
                    "Description": description,
                    "Department": department,
                    "Solicitation Number": solnum,
                    "Type": bid_type,
                })

        except StaleElementReferenceException:
            logging.warning(f"Row {i} went stale; skipping.")
        except Exception as e:
            logging.warning(f"Unexpected error scraping row {i}: {e}")

    return results


def get_first_row_marker(driver, table):
    try:
        tbody = table.find_element(*TBODY_LOCATOR)
        rows = tbody.find_elements(*ROW_LOCATOR)
        if not rows:
            return f"no-rows-{time.time()}"
        first = rows[0]
        try:
            sol = first.find_element(By.CSS_SELECTOR, 'a[data-qa*=".DOC_REF"], a.css-xv6zqn[data-qa*=".DOC_REF"]')
            return extract_text_safe(sol) or f"empty-sol-{time.time()}"
        except NoSuchElementException:
            return f"no-sol-{time.time()}"
    except Exception:
        return f"err-{time.time()}"


def is_next_clickable(driver):
    try:
        btn = driver.find_element(*NEXT_BUTTON_LOCATOR)
    except NoSuchElementException:
        return False, None
    disabled = btn.get_attribute("disabled")
    aria_disabled = btn.get_attribute("aria-disabled")
    if disabled is not None or (aria_disabled and aria_disabled.lower() == "true"):
        return False, btn
    try:
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable(NEXT_BUTTON_LOCATOR))
        return True, btn
    except TimeoutException:
        return False, btn


def click_next_and_wait_for_change(driver, table, timeout=20):
    can_click, btn = is_next_clickable(driver)
    if not can_click or btn is None:
        return False
    before_marker = get_first_row_marker(driver, table)
    js_scroll_center(driver, btn)
    time.sleep(0.08)
    try:
        btn.click()
    except (ElementClickInterceptedException, StaleElementReferenceException, WebDriverException):
        try:
            btn = driver.find_element(*NEXT_BUTTON_LOCATOR)
            driver.execute_script("arguments[0].click();", btn)
        except Exception:
            return False

    end = time.time() + timeout
    while time.time() < end:
        time.sleep(0.4)
        try:
            new_table = find_results_table(driver, timeout=5)
        except Exception:
            continue
        after_marker = get_first_row_marker(driver, new_table)
        if after_marker != before_marker:
            logging.info("Pagination advanced.")
            return True
    logging.info("Pagination did not change (likely last page).")
    return False


def write_csv(path, rows):
    fieldnames = ["Description", "Department", "Solicitation Number", "Type"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in rows:
            writer.writerow(r)


def clean_type_column(input_path: str, output_path: str) -> int:
    changed = 0
    out_rows = []
    with open(input_path, "r", encoding="utf-8", newline="") as fin:
        reader = csv.DictReader(fin)
        fieldnames = reader.fieldnames or []
        if "Type" not in fieldnames:
            fieldnames.append("Type")
        for row in reader:
            original = row.get("Type", "")
            cleaned = sanitize_type(original)
            if cleaned != original:
                changed += 1
            row["Type"] = cleaned
            out_rows.append(row)

    with open(output_path, "w", encoding="utf-8", newline="") as fout:
        writer = csv.DictWriter(fout, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(out_rows)
    return changed


def main():
    driver = build_driver()
    all_rows = []
    try:
        logging.info("Opening WV VSS…")
        driver.get(START_URL)
        wait_for_page_ready(driver)
        time.sleep(1)

        logging.info("Navigating to 'View Published Solicitations'…")
        click_view_published_solicitations(driver)
        wait_for_page_ready(driver)
        time.sleep(1)

        logging.info("Expanding filters and applying selections…")
        expand_show_more(driver)
        set_status_awarded(driver)
        set_show_me_all(driver)

        logging.info("Submitting the search…")
        click_search(driver)

        logging.info("Waiting for results table…")
        table = find_results_table(driver)
        time.sleep(0.6)

        page_num = 1
        while True:
            logging.info(f"Scraping page {page_num}…")
            # re-resolve table each loop in case DOM refreshes
            table = find_results_table(driver, timeout=20)
            page_rows = scrape_rows_from_current_page(driver, table)
            logging.info(f"Collected {len(page_rows)} rows on page {page_num}.")
            all_rows.extend(page_rows)

            time.sleep(0.4)
            advanced = click_next_and_wait_for_change(driver, table)
            if not advanced:
                logging.info("Next not clickable or no change — done paginating.")
                break
            page_num += 1

        if not all_rows:
            print("\nNo records")
            return

        write_csv(CSV_PATH_RAW, all_rows)
        logging.info(f"Wrote {len(all_rows)} rows to {CSV_PATH_RAW}.")

        changed = clean_type_column(CSV_PATH_RAW, CSV_PATH_CLEAN)
        logging.info(f"Cleaned 'Type' column in {changed} rows -> {CSV_PATH_CLEAN}")

        print(f"\n✅ Scrape complete. {len(all_rows)} rows saved.")
        print(f"   Raw CSV:   {CSV_PATH_RAW}")
        print(f"   Clean CSV: {CSV_PATH_CLEAN} (Type column fixed)")
        input("Press ENTER to close the browser.")
    except TimeoutException as e:
        logging.error("Timed out waiting for an element: %s", e)
        print("\n❌ Timeout—an expected control didn’t appear. We can tweak waits/selectors if needed.")
        input("Press ENTER to close the browser.")
    finally:
        driver.quit()


if __name__ == "__main__":
    main()