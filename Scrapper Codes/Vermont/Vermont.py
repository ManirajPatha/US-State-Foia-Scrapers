import time
import json
import logging
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# --------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------
BASE_URL = "https://www.vermontbusinessregistry.com/BidSearch.aspx"
OUTPUT_FILE = "vermont_awarded_rfp_results.json"
LOG_FILE = "vermont_scraper.log"
INCREMENTAL_SAVE_EVERY = 5

os.makedirs("output", exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8"),
              logging.StreamHandler()]
)
log = logging.getLogger()

# --------------------------------------------------------------------
# SETUP CHROME
# --------------------------------------------------------------------
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-notifications")
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# --------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------
def safe_find(by, value):
    """Return element text if present, else empty string."""
    try:
        return driver.find_element(by, value).text.strip()
    except:
        return ""

def save_json(records):
    """Save JSON results incrementally."""
    path = os.path.join("output", OUTPUT_FILE)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2, ensure_ascii=False)
    log.info(f"Saved {len(records)} records → {OUTPUT_FILE}")

# --------------------------------------------------------------------
# NAVIGATION
# --------------------------------------------------------------------
def open_power_search():
    driver.get(BASE_URL)
    log.info("Opened Vermont Business Registry main page")
    power_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='Table1']/tbody/tr[3]/td[2]/a[2]")))
    driver.execute_script("arguments[0].click();", power_link)
    log.info("Clicked Power Search link")

def set_filters():
    status_dropdown = Select(wait.until(EC.presence_of_element_located((By.ID, "ddlPowerBidStatus"))))
    status_dropdown.select_by_value("CLOSE/AWAR")
    log.info("Selected bid status: Closed/Awarded")

    type_dropdown = Select(wait.until(EC.presence_of_element_located((By.ID, "ddlPowerBidTypes"))))
    type_dropdown.select_by_value("1")
    log.info("Selected bid type: RFP")

    search_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnPowerSearch")))
    driver.execute_script("arguments[0].click();", search_btn)
    log.info("Clicked Search button")

# --------------------------------------------------------------------
# SCRAPE DETAIL PAGE
# --------------------------------------------------------------------
def scrape_awardees():
    awardees = []
    rows = driver.find_elements(By.XPATH, "//table[@border='0' and @width='700']//tr")
    current = {"name": "", "date": "", "addr1": "", "addr2": "", "amount": ""}

    for row in rows:
        spans = row.find_elements(By.TAG_NAME, "span")
        if not spans:
            continue
        for sp in spans:
            sid = sp.get_attribute("id")
            text = sp.text.strip()
            if not text:
                continue
            if "lblAwardeeName" in sid:
                current["name"] = text
            elif "lblAwareeDate" in sid:
                current["date"] = text
            elif "lblAwardeeAddress1" in sid:
                current["addr1"] = text
            elif "lblAwardeeAddress2" in sid:
                current["addr2"] = text
            elif "lblAwardeeAmount" in sid:
                current["amount"] = text

        # When both address1 or address2 appear, treat as a complete record
        if current["name"] and (current["addr1"] or current["addr2"]):
            city, state = "", ""
            if current["addr2"] and "," in current["addr2"]:
                parts = current["addr2"].split(",")
                city = parts[0].strip()
                state = parts[1].strip() if len(parts) > 1 else ""
            awardees.append({
                "awardee_name": current["name"],
                "awardee_address": current["addr1"],
                "awardee_city": city,
                "awardee_state": state,
                "award_date": current["date"],
                "award_amount": current["amount"]
            })
            current = {"name": "", "date": "", "addr1": "", "addr2": "", "amount": ""}

    return awardees

def scrape_detail():
    bid_title = safe_find(By.ID, "lblBidTitle")
    department = safe_find(By.ID, "lblAuthorName")
    bid_description = safe_find(By.ID, "lblBidDescription")
    buyer_address = safe_find(By.ID, "lblAddressLine1")
    buyer_city = safe_find(By.ID, "lblCityTownRegion")
    buyer_state = safe_find(By.ID, "lblStateCode")
    buyer_zip = safe_find(By.ID, "lblPostalCode")
    contract_value = safe_find(By.ID, "lblEstDollarValue")

    awardees = scrape_awardees()
    records = []
    for a in awardees:
        rec = {
            "bid_title": bid_title,
            "department": department,
            "bid_description": bid_description,
            "buyer_address": buyer_address,
            "buyer_city": buyer_city,
            "buyer_state": buyer_state,
            "buyer_zip": buyer_zip,
            "contract_value": contract_value,
            **a
        }
        records.append(rec)
    return records

# --------------------------------------------------------------------
# MAIN LOOP WITH POPUP HANDLING
# --------------------------------------------------------------------
def scrape_all():
    open_power_search()
    set_filters()
    time.sleep(3)

    all_records = []
    page = 1

    while True:
        log.info(f"Processing page {page} …")
        links = driver.find_elements(By.XPATH, "//a[contains(@href,'BidPreview.aspx?BidID=')]")
        log.info(f"Found {len(links)} bids on page {page}")

        for idx, link in enumerate(links, start=1):
            main_window = driver.current_window_handle
            current_windows = driver.window_handles

            try:
                driver.execute_script("arguments[0].click();", link)

                # Wait for new popup window
                WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(current_windows))
                new_window = [w for w in driver.window_handles if w not in current_windows][0]
                driver.switch_to.window(new_window)

                wait.until(EC.presence_of_element_located((By.ID, "lblBidTitle")))
                recs = scrape_detail()
                if recs:
                    all_records.extend(recs)
                    log.info(f" Added {len(recs)} awardee records from bid {idx}")
                else:
                    log.info(f"Skipped bid {idx} (no awardee found)")

            except TimeoutException:
                log.warning(f"No popup detected for record {idx}, skipping.")
            except Exception as e:
                log.warning(f"Error scraping record {idx}: {e}")
            finally:
                # Close popup safely if it still exists
                try:
                    for w in driver.window_handles:
                        if w != main_window:
                            driver.switch_to.window(w)
                            driver.close()
                    driver.switch_to.window(main_window)
                except Exception as e2:
                    log.warning(f"Cleanup issue: {e2}")

            if len(all_records) % INCREMENTAL_SAVE_EVERY == 0:
                save_json(all_records)

        # Pagination
        next_links = driver.find_elements(By.XPATH, "//a[starts-with(@href,\"javascript:__doPostBack('gvResults','Page$\")]")
        if not next_links:
            log.info("Reached last page.")
            break

        clicked = False
        for l in next_links:
            if l.text.strip() == str(page + 1):
                driver.execute_script("arguments[0].click();", l)
                time.sleep(3)
                clicked = True
                page += 1
                break
        if not clicked:
            log.info("No more pages.")
            break

    save_json(all_records)
    log.info(" Scraping completed successfully with %d records.", len(all_records))

# --------------------------------------------------------------------
# RUN
# --------------------------------------------------------------------
if __name__ == "__main__":
    try:
        scrape_all()
    finally:
        driver.quit()
