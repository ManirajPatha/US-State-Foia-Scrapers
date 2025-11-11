import time
import json
import logging
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ---------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------
BASE_URL = "https://wwwcfprd.doa.louisiana.gov/osp/lapac/altlist.cfm"
OUTPUT_FILE = "louisiana_awarded_results.json"
INCREMENTAL_SAVE_EVERY = 3
SLEEP_BETWEEN_RECORDS = 1.0

# ---------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("louisiana_scraper.log", mode="w", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger()

# ---------------------------------------------------------------------
# CHROME DRIVER (visible mode)
# ---------------------------------------------------------------------
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 15)

# ---------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------
def save_json_incremental(records):
    if not records:
        return
    if os.path.exists(OUTPUT_FILE):
        with open(OUTPUT_FILE, "r", encoding="utf-8") as f:
            existing = json.load(f)
    else:
        existing = []
    existing.extend(records)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(existing, f, indent=2, ensure_ascii=False)
    log.info(f"Incremental save: +{len(records)} records")

def safe_open_new_tab(url):
    current_tabs = set(driver.window_handles)
    driver.execute_script("window.open(arguments[0]);", url)
    WebDriverWait(driver, 10).until(
        lambda d: len(set(d.window_handles) - current_tabs) == 1
    )
    new_tab = list(set(driver.window_handles) - current_tabs)[0]
    time.sleep(0.5)
    driver.switch_to.window(new_tab)
    return new_tab

def safe_close_and_return(main_handle):
    if len(driver.window_handles) > 1:
        driver.close()
        time.sleep(0.5)
        driver.switch_to.window(main_handle)

# ---------------------------------------------------------------------
# MAIN SCRAPER
# ---------------------------------------------------------------------
def main():
    driver.get(BASE_URL)
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="mainbg"]/div[3]/table')))
    main_tab = driver.current_window_handle
    rows = driver.find_elements(By.XPATH, '//*[@id="mainbg"]/div[3]/table/tbody/tr')[1:]
    log.info(f"Found {len(rows)} rows")

    batch = []
    processed = 0

    for i, row in enumerate(rows, start=1):
        try:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) < 3:
                continue

            status = cols[1].text.strip().lower()
            if status != "awarded":
                continue

            bid_link = cols[0].find_element(By.TAG_NAME, "a")
            bid_number = bid_link.text.strip()
            description = cols[2].text.strip()
            bid_url = bid_link.get_attribute("href")

            log.info(f"[{i}] Awarded bid: {bid_number} – {description}")

            new_tab = safe_open_new_tab(bid_url)
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="mainbg"]/div[3]/table')))

                # --- Award Date ---
                try:
                    award_date_raw = driver.find_element(
                        By.XPATH, "//*[contains(text(),'Date of Award')]"
                    ).get_attribute("textContent")
                    award_date = award_date_raw.replace("Date of Award:", "").strip()
                except:
                    award_date = ""

                # --- Awardee Name (dynamic lookup) ---
                try:
                    contractor_td = driver.find_element(By.XPATH, "//td[contains(.,'Contractor:')]")
                    contractor_text = contractor_td.get_attribute("textContent")
                    awardee_name = contractor_text.replace("Contractor:", "").strip()
                except:
                    awardee_name = ""

                # --- Award Amount (dynamic lookup) ---
                try:
                    amount_td = driver.find_element(By.XPATH, "//td[contains(.,'Amount')]")
                    amount_text = amount_td.get_attribute("textContent")
                    award_amount = amount_text.replace("Amount:", "").strip()
                except:
                    award_amount = ""

                # --- Attachments ---
                attachments = []
                try:
                    attach_links = driver.find_elements(
                        By.XPATH, "//a[contains(@href,'/osp/lapac/agency/pdf/')]"
                    )
                    for a in attach_links:
                        href = a.get_attribute("href")
                        text = a.text.strip()
                        attachments.append({"name": text, "url": href})
                except:
                    pass

                record = {
                    "bid_number": bid_number,
                    "description": description,
                    "award_date": award_date,
                    "awardee_name": awardee_name,
                    "award_amount": award_amount,
                    "attachments": attachments,
                    "detail_url": bid_url,
                }

                batch.append(record)
                processed += 1
                log.info(f"→ Saved {bid_number}")

                safe_close_and_return(main_tab)

                if processed % INCREMENTAL_SAVE_EVERY == 0:
                    save_json_incremental(batch)
                    batch.clear()

                time.sleep(SLEEP_BETWEEN_RECORDS)

            except Exception as e:
                log.error(f"Error on detail page {bid_number}: {e}")
                safe_close_and_return(main_tab)
                continue

        except Exception as e:
            log.error(f"Row {i} error: {e}")
            continue

    if batch:
        save_json_incremental(batch)
    log.info("✅ Scraping complete.")

# ---------------------------------------------------------------------
if __name__ == "__main__":
    try:
        main()
    finally:
        driver.quit()
