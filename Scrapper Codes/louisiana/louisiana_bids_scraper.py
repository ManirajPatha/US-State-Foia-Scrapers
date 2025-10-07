# filename: la_lapac_awarded_scraper.py
import csv
import logging
import time
from pathlib import Path
from typing import List, Dict

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException, ElementClickInterceptedException

URL = "https://wwwcfprd.doa.louisiana.gov/osp/lapac/altlist.cfm"
OUTPUT_CSV = "la_awarded_bids.csv"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

def norm_text(s: str) -> str:
    return " ".join((s or "").split())

def create_driver(headless: bool = False) -> webdriver.Firefox:
    opts = Options()
    if headless:
        opts.add_argument("-headless")
    # Helpful for some sites that care about UA (optional)
    opts.set_preference("general.useragent.override",
                        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:128.0) Gecko/20100101 Firefox/128.0")
    driver = webdriver.Firefox(options=opts)
    driver.maximize_window()
    return driver

def wait_for_table(driver: webdriver.Firefox, timeout: int = 20):
    wait = WebDriverWait(driver, timeout)
    # Wait for the table body of class 'bid'
    return wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.bid tbody")))

def parse_page_rows(driver: webdriver.Firefox) -> List[Dict[str, str]]:
    """
    Scrape all rows on the current page and return awarded rows with:
    - Bid Number (from the first <a> in the row)
    - Description (from the last-but-one <td>)
    Only rows where a <td class="txt"> has text exactly 'Awarded' (case-insensitive) are returned.
    """
    results: List[Dict[str, str]] = []
    tbody = driver.find_element(By.CSS_SELECTOR, "table.bid tbody")
    rows = tbody.find_elements(By.CSS_SELECTOR, "tr")

    for tr in rows:
        try:
            tds = tr.find_elements(By.TAG_NAME, "td")
            if not tds:
                continue

            # Find any td with class 'txt' that equals 'Awarded'
            awarded = False
            for td in tds:
                cls = td.get_attribute("class") or ""
                if "txt" in cls.split():
                    if norm_text(td.text).strip().lower() == "awarded":
                        awarded = True
                        break
            if not awarded:
                continue

            # Bid Number: first <a> in the row
            bid_number = ""
            try:
                a = tr.find_element(By.CSS_SELECTOR, "a")
                bid_number = norm_text(a.text)
            except NoSuchElementException:
                # Fallback: look in tds for an <a>
                for td in tds:
                    try:
                        a2 = td.find_element(By.TAG_NAME, "a")
                        bid_number = norm_text(a2.text)
                        if bid_number:
                            break
                    except NoSuchElementException:
                        continue

            # Description: last but one td
            description = ""
            if len(tds) >= 2:
                description = norm_text(tds[-2].text)

            if bid_number and description:
                results.append({
                    "Bid Number": bid_number,
                    "Description": description
                })
        except StaleElementReferenceException:
            # If the table re-rendered mid-iteration, skip this row
            continue

    return results

def click_next_if_available(driver: webdriver.Firefox) -> bool:
    """
    Attempts to click a 'Next' pagination control if present.
    Returns True if clicked (and navigation started), False otherwise.
    Tries a few common patterns used on LaPAC pages.
    """
    # Try anchor with text 'Next'
    selectors = [
        ("xpath", "//a[normalize-space(.)='Next']"),
        # Sometimes there might be an input or button with value 'Next'
        ("xpath", "//input[@type='button' or @type='submit'][@value='Next']"),
        # Any anchor/button containing 'Next'
        ("xpath", "//*[self::a or self::button][contains(normalize-space(.), 'Next')]")
    ]

    for how, sel in selectors:
        try:
            if how == "xpath":
                el = driver.find_element(By.XPATH, sel)
            else:
                el = driver.find_element(By.CSS_SELECTOR, sel)

            # Check disabled state heuristically
            disabled = el.get_attribute("disabled")
            classes = (el.get_attribute("class") or "").lower()
            if (disabled and disabled != "false") or "disabled" in classes:
                continue

            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            time.sleep(0.2)
            el.click()
            # Brief wait for page to load/refresh table
            WebDriverWait(driver, 10).until(EC.staleness_of(el))
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.bid tbody")))
            return True
        except (NoSuchElementException, ElementClickInterceptedException, TimeoutException, StaleElementReferenceException):
            continue

    return False

def write_csv(records: List[Dict[str, str]], path: str):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["Bid Number", "Description"])
        writer.writeheader()
        for row in records:
            writer.writerow(row)

def main():
    driver = create_driver(headless=False)  # set True if you want headless
    all_results: List[Dict[str, str]] = []
    try:
        logging.info("Opening LaPAC page…")
        driver.get(URL)
        wait_for_table(driver, timeout=30)

        page_idx = 1
        while True:
            logging.info(f"Scraping page {page_idx}…")
            page_results = parse_page_rows(driver)
            logging.info(f"Found {len(page_results)} awarded rows on this page.")
            all_results.extend(page_results)

            # Try to go to next page; break if none
            if not click_next_if_available(driver):
                break
            page_idx += 1

        if not all_results:
            logging.info("No records")
            return

        write_csv(all_results, OUTPUT_CSV)
        logging.info(f"Done. Wrote {len(all_results)} awarded rows to {Path(OUTPUT_CSV).resolve()}")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()