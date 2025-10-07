# michigan_vss_step3_clean.py
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

# === Michigan START URL ===
START_URL = "https://sigma.michigan.gov/PRDVSS1X1/Advantage4"

CSV_PATH_RAW = "michigan_vss_results.csv"
CSV_PATH_CLEAN = "michigan_vss_results_clean.csv"

# Results table (same structure as Alaska tenant; adjust if needed)
TABLE_LOCATOR = (
    By.CSS_SELECTOR,
    'div.css-jtdyb table#vsspageVVSSX10019gridView1group1cardGridgrid1.css-1uq2am7.css-if7ucn'
)
TBODY_LOCATOR = (By.CSS_SELECTOR, "tbody")
ROW_LOCATOR = (By.CSS_SELECTOR, "tbody > tr")

# Michigan pager's Next button (button wrapping the <i> icon)
NEXT_BUTTON_LOCATOR = (By.CSS_SELECTOR, 'button.css-1yn6b58[aria-label="Next"]')

# Search button (unchanged)
SEARCH_BUTTON_LOCATOR = (By.CSS_SELECTOR, 'button[name="vss.page.VVSSX10019.gridView1.Search"][aria-label="Search"]')

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


def build_driver() -> webdriver.Firefox:
    service = Service()  # set executable_path if geckodriver isn't on PATH
    opts = Options()
    # opts.add_argument("-headless")  # uncomment to run headless
    driver = webdriver.Firefox(service=service, options=opts)
    driver.maximize_window()
    driver.set_page_load_timeout(60)
    return driver


def wait_for_page_ready(driver, timeout=30):
    WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")


def js_scroll_center(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)


def click_initial_next_to_reveal_filters(driver, timeout=15):
    """
    Michigan flow: click the 'Next' on the landing carousel to expose 'View Published Solicitations'.
    <div class="ng-tns-c151-0 css-1h8g0ba">
      <a class="ng-tns-c151-0 css-1ptusf6 css-kny0og" title="Next" role="button">...</a>
    """
    wait = WebDriverWait(driver, timeout)
    locator = (
        By.CSS_SELECTOR,
        'div.ng-tns-c151-0.css-1h8g0ba a.ng-tns-c151-0.css-1ptusf6.css-kny0og[title="Next"][role="button"]'
    )
    try:
        el = wait.until(EC.presence_of_element_located(locator))
        js_scroll_center(driver, el)
        try:
            wait.until(EC.element_to_be_clickable(locator))
        except TimeoutException:
            pass
        try:
            el.click()
        except (ElementClickInterceptedException, StaleElementReferenceException, WebDriverException):
            el = driver.find_element(*locator)
            driver.execute_script("arguments[0].click();", el)
        logging.info("Clicked initial landing 'Next' to reveal options.")
        time.sleep(0.7)
    except TimeoutException:
        logging.info("Landing 'Next' not found; proceeding (maybe already on the right panel).")


def click_view_published_solicitations(driver, timeout=30):
    wait = WebDriverWait(driver, timeout)
    locator = (
        By.CSS_SELECTOR,
        'div[title="View Published Solicitations"][aria-label="View Published Solicitations"]'
    )
    el = wait.until(EC.presence_of_element_located(locator))
    js_scroll_center(driver, el)
    try:
        wait.until(EC.element_to_be_clickable(locator))
    except TimeoutException:
        pass
    try:
        el.click()
    except (ElementClickInterceptedException, StaleElementReferenceException, WebDriverException):
        el = driver.find_element(*locator)
        driver.execute_script("arguments[0].click();", el)
    logging.info("Opened 'View Published Solicitations'.")


def expand_show_more(driver, timeout=20):
    """
    Robustly expand filters if a 'Show More' exists; skip if already expanded.
    """
    wait = WebDriverWait(driver, timeout)

    # Already visible?
    try:
        wait.until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            'select[name="vss.page.VVSSX10019.gridView1.group1.cardSearch.search1.SO_STA"]'
        )))
        logging.info("Filters already visible; skipping 'Show More'.")
        return
    except TimeoutException:
        pass

    # Try common variants of Show More
    candidates = [
        (By.CSS_SELECTOR, 'button[aria-label*="Show More" i]'),
        (By.CSS_SELECTOR, 'button[title*="Show More" i]'),
        (By.XPATH, "//button[contains(., 'Show More')]"),
        (By.CSS_SELECTOR, 'a[role="button"][aria-label*="Show More" i]'),
        (By.CSS_SELECTOR, 'a[role="button"][title*="Show More" i]'),
        (By.XPATH, "//a[@role='button' and contains(., 'Show More')]"),
        (By.CSS_SELECTOR, '[aria-expanded="false"][aria-label*="Show" i]'),
        (By.CSS_SELECTOR, '[aria-expanded="false"][title*="Show" i]'),
    ]

    for how, sel in candidates:
        try:
            ctrl = wait.until(EC.presence_of_element_located((how, sel)))
            js_scroll_center(driver, ctrl)
            try:
                wait.until(EC.element_to_be_clickable((how, sel)))
            except TimeoutException:
                pass
            try:
                ctrl.click()
            except (ElementClickInterceptedException, StaleElementReferenceException, WebDriverException):
                try:
                    ctrl = driver.find_element(how, sel)
                    driver.execute_script("arguments[0].click();", ctrl)
                except Exception:
                    continue

            # Verify filters appeared
            try:
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((
                    By.CSS_SELECTOR,
                    'select[name="vss.page.VVSSX10019.gridView1.group1.cardSearch.search1.SO_STA"]'
                )))
                logging.info("Expanded filters via 'Show More'.")
                return
            except TimeoutException:
                continue
        except TimeoutException:
            continue

    # Final check
    try:
        wait.until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            'select[name="vss.page.VVSSX10019.gridView1.group1.cardSearch.search1.SO_STA"]'
        )))
        logging.info("Filters visible; no 'Show More' needed.")
    except TimeoutException:
        logging.info("Could not find 'Show More' and filters still hidden.")


def set_status_awarded(driver, timeout=20):
    wait = WebDriverWait(driver, timeout)
    sel_locator = (
        By.CSS_SELECTOR,
        'select[name="vss.page.VVSSX10019.gridView1.group1.cardSearch.search1.SO_STA"][aria-label="Status"]'
    )
    sel = wait.until(EC.presence_of_element_located(sel_locator))
    js_scroll_center(driver, sel)
    Select(sel).select_by_value("A")  # Awarded
    logging.info("Set Status = Awarded (A).")
    time.sleep(0.2)


def set_show_me_all(driver, timeout=20):
    wait = WebDriverWait(driver, timeout)
    sel_locator = (
        By.CSS_SELECTOR,
        'select[name="vss.page.VVSSX10019.gridView1.group1.cardSearch.search1.SHOW_TXT"][aria-label="Show Me"]'
    )
    sel = wait.until(EC.presence_of_element_located(sel_locator))
    js_scroll_center(driver, sel)
    Select(sel).select_by_value("1")  # All
    logging.info("Set Show Me = All (1).")
    time.sleep(0.2)


def set_type_rfp(driver, timeout=20):
    """
    Type = Request for Proposals (value 'RFP').
    """
    wait = WebDriverWait(driver, timeout)
    try:
        sel = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//select[option/@value='RFP' or option[normalize-space(text())='Request for Proposals']]")
        ))
        js_scroll_center(driver, sel)
        Select(sel).select_by_value("RFP")
        logging.info("Set Type = Request for Proposals (RFP).")
        time.sleep(0.2)
    except TimeoutException:
        logging.warning("Type select with option 'RFP' not found; continuing without it.")


def click_search(driver, timeout=20):
    wait = WebDriverWait(driver, timeout)
    btn = wait.until(EC.presence_of_element_located(SEARCH_BUTTON_LOCATOR))
    js_scroll_center(driver, btn)
    try:
        wait.until(EC.element_to_be_clickable(SEARCH_BUTTON_LOCATOR))
    except TimeoutException:
        pass
    try:
        btn.click()
    except (ElementClickInterceptedException, StaleElementReferenceException, WebDriverException):
        btn = driver.find_element(*SEARCH_BUTTON_LOCATOR)
        driver.execute_script("arguments[0].click();", btn)
    logging.info("Clicked Search.")


def wait_for_table(driver, timeout=30):
    wait = WebDriverWait(driver, timeout)
    table = wait.until(EC.presence_of_element_located(TABLE_LOCATOR))
    js_scroll_center(driver, table)
    wait.until(EC.presence_of_element_located(TBODY_LOCATOR))
    wait.until(EC.presence_of_all_elements_located(ROW_LOCATOR))
    logging.info("Results table present.")
    return table


def extract_text_safe(el):
    try:
        return (el.get_attribute("aria-label") or el.text or "").strip()
    except Exception:
        return ""


def sanitize_type(value: str) -> str:
    """Remove leading 'null' (any case) and punctuation following it."""
    if value is None:
        return ""
    cleaned = re.sub(r"^\s*null[\s:,-]*", "", str(value), flags=re.IGNORECASE).strip()
    return cleaned


def scrape_rows_from_current_page(driver):
    results = []
    table = driver.find_element(*TABLE_LOCATOR)
    tbody = table.find_element(*TBODY_LOCATOR)
    rows = tbody.find_elements(*ROW_LOCATOR)

    logging.info(f"Found {len(rows)} rows on this page.")
    for i, row in enumerate(rows, start=1):
        try:
            js_scroll_center(driver, row)
            time.sleep(0.05)

            # Description
            try:
                desc_el = row.find_element(By.CSS_SELECTOR, '[data-qa*=".DOC_DSCR"]')
                description = extract_text_safe(desc_el)
            except NoSuchElementException:
                description = ""

            # Department
            try:
                dept_el = row.find_element(By.CSS_SELECTOR, '[data-qa*=".DeptBuyr.DEPT_NM"]')
                department = extract_text_safe(dept_el)
            except NoSuchElementException:
                department = ""

            # Solicitation Number
            try:
                sol_el = row.find_element(By.CSS_SELECTOR, 'a.css-xv6zqn[data-qa*=".DOC_REF"]')
                solicitation_number = extract_text_safe(sol_el)
            except NoSuchElementException:
                solicitation_number = ""

            # Type (sanitize)
            try:
                type_el = row.find_element(By.CSS_SELECTOR, '[data-qa*=".DOC_CD_CONCAT"]')
                bid_type = sanitize_type(extract_text_safe(type_el))
            except NoSuchElementException:
                bid_type = ""

            if any([description, department, solicitation_number, bid_type]):
                results.append({
                    "Description": description,
                    "Department": department,
                    "Solicitation Number": solicitation_number,
                    "Type": bid_type,
                })

        except StaleElementReferenceException:
            logging.warning(f"Row {i} went stale; skipping.")
        except Exception as e:
            logging.warning(f"Unexpected error scraping row {i}: {e}")

    return results


def get_first_row_solnum_or_marker(driver):
    """A marker to detect page change."""
    try:
        table = driver.find_element(*TABLE_LOCATOR)
        tbody = table.find_element(*TBODY_LOCATOR)
        rows = tbody.find_elements(*ROW_LOCATOR)
        if not rows:
            return f"no-rows-{time.time()}"
        first = rows[0]
        try:
            sol = first.find_element(By.CSS_SELECTOR, 'a.css-xv6zqn[data-qa*=".DOC_REF"]')
            return extract_text_safe(sol) or f"empty-sol-{time.time()}"
        except NoSuchElementException:
            return f"no-sol-{time.time()}"
    except Exception:
        return f"err-{time.time()}"


def find_next_button(driver, timeout=8):
    """
    Find the Michigan pager's Next button:
      <button class="css-1yn6b58" aria-label="Next"> <i class="adv-svg-Single-Arrow-Right-icon" ...> </button>
    """
    wait = WebDriverWait(driver, timeout)
    try:
        btn = wait.until(EC.presence_of_element_located(NEXT_BUTTON_LOCATOR))
        js_scroll_center(driver, btn)
        return btn
    except TimeoutException:
        # Fallback: find by aria-label without class (just in case)
        try:
            btn = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'button[aria-label="Next"]')
            ))
            js_scroll_center(driver, btn)
            return btn
        except TimeoutException:
            # Last resort: click the icon, then its ancestor button
            try:
                icon = wait.until(EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'i.adv-svg-Single-Arrow-Right-icon[title="Next"]')
                ))
                try:
                    btn = icon.find_element(By.XPATH, ".//ancestor::button[1]")
                except NoSuchElementException:
                    btn = icon
                js_scroll_center(driver, btn)
                return btn
            except TimeoutException:
                return None


def is_disabled(el):
    try:
        if el.get_attribute("disabled") is not None:
            return True
        if (el.get_attribute("aria-disabled") or "").lower() == "true":
            return True
        cls = (el.get_attribute("class") or "").lower()
        return "disabled" in cls
    except Exception:
        return False


def click_next_and_wait_for_change(driver, timeout=20):
    btn = find_next_button(driver, timeout=8)
    if not btn:
        logging.info("Next button not found.")
        return False
    if is_disabled(btn):
        logging.info("Next button is disabled (likely last page).")
        return False

    before_marker = get_first_row_solnum_or_marker(driver)

    # Some tenants require results container scrolled fully for pager to work
    try:
        container = driver.find_element(By.CSS_SELECTOR, "div.css-jtdyb")
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", container)
    except Exception:
        pass

    try:
        btn.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", btn)
        except Exception:
            logging.info("Could not click Next button.")
            return False

    end = time.time() + timeout
    while time.time() < end:
        time.sleep(0.4)
        after_marker = get_first_row_solnum_or_marker(driver)
        if after_marker != before_marker:
            logging.info("Pagination advanced to next page.")
            return True

    logging.info("Pagination did not change (last page or virtualized rows not updated yet).")
    return False


def write_csv(path, rows):
    fieldnames = ["Description", "Department", "Solicitation Number", "Type"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in rows:
            writer.writerow(r)


def clean_type_column(input_path: str, output_path: str) -> int:
    """
    Reads input CSV and writes output CSV with 'Type' column cleaned.
    Returns number of rows that changed.
    """
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
        logging.info("Opening Michigan VSS (Advantage4)…")
        driver.get(START_URL)
        wait_for_page_ready(driver)
        time.sleep(1)

        logging.info("Revealing options via landing Next…")
        click_initial_next_to_reveal_filters(driver)

        logging.info("Opening 'View Published Solicitations'…")
        click_view_published_solicitations(driver)
        wait_for_page_ready(driver)
        time.sleep(1)

        logging.info("Expanding filters and applying selections…")
        expand_show_more(driver)
        set_status_awarded(driver)
        set_show_me_all(driver)
        set_type_rfp(driver)

        logging.info("Submitting the search…")
        click_search(driver)

        logging.info("Waiting for results table…")
        wait_for_table(driver)
        time.sleep(1.0)

        page_num = 1
        while True:
            logging.info(f"Scraping page {page_num}…")
            WebDriverWait(driver, 30).until(EC.presence_of_element_located(TABLE_LOCATOR))
            WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located(ROW_LOCATOR))

            page_rows = scrape_rows_from_current_page(driver)
            logging.info(f"Collected {len(page_rows)} rows on page {page_num}.")
            all_rows.extend(page_rows)

            time.sleep(0.6)  # small visual pause

            advanced = click_next_and_wait_for_change(driver)
            if not advanced:
                logging.info("Next not clickable or no change — done paginating.")
                break
            page_num += 1

        # Write RAW CSV
        write_csv(CSV_PATH_RAW, all_rows)
        logging.info(f"Wrote {len(all_rows)} rows to {CSV_PATH_RAW}.")

        # Clean 'Type' column -> CLEAN CSV
        changed = clean_type_column(CSV_PATH_RAW, CSV_PATH_CLEAN)
        logging.info(f"Cleaned 'Type' in {changed} rows -> {CSV_PATH_CLEAN}")

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