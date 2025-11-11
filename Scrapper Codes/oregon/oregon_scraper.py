# oregon_buys_bid_to_po_scraper_v2.py
import os, time, json, logging, re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# --------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------
BASE_URL = "https://oregonbuys.gov/bso/view/search/external/advancedSearchBid.xhtml?openBids=true"
OUTPUT_FILE = "oregon_buys_bid_to_po_results.json"
DOWNLOAD_DIR = os.path.abspath("attachments_oregon")
TIMEOUT = 25
LONG_TIMEOUT = 60
SAVE_EVERY = 25

os.makedirs(DOWNLOAD_DIR, exist_ok=True)
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# --------------------------------------------------------------------
# BROWSER SETUP
# --------------------------------------------------------------------
def open_browser(headless=False):
    opts = webdriver.ChromeOptions()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1600,1000")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    opts.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True
    })
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": DOWNLOAD_DIR
        })
    except Exception:
        pass
    driver.get(BASE_URL)
    return driver

# --------------------------------------------------------------------
# FILTER ACTIONS
# --------------------------------------------------------------------
def open_advanced_search(driver):
    try:
        legend = WebDriverWait(driver, TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='advancedSearchMainPanelContainer']/legend"))
        )
        driver.execute_script("arguments[0].click();", legend)
        logging.info("Expanded Advanced Search section.")
        time.sleep(1)
    except Exception as e:
        logging.warning(f"Advanced search expand failed: {e}")

def set_status_bid_to_po(driver):
    sel = WebDriverWait(driver, TIMEOUT).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='bidSearchForm:status']"))
    )
    Select(sel).select_by_value("2BPO")
    logging.info("Selected Status = Bid to PO")

def click_search(driver):
    btn = WebDriverWait(driver, TIMEOUT).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='bidSearchForm:btnBidSearch']"))
    )
    driver.execute_script("arguments[0].click();", btn)
    logging.info("Clicked Search button")

# --------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------
def get_text_safe(row, index):
    try:
        return row.find_element(By.XPATH, f"./td[{index}]").text.strip()
    except Exception:
        return ""

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name)

# --------------------------------------------------------------------
# ATTACHMENT SCRAPER
# --------------------------------------------------------------------
def scrape_attachments(driver, solicitation_url):
    attachments = []
    if not solicitation_url:
        return attachments

    driver.execute_script("window.open(arguments[0]);", solicitation_url)
    driver.switch_to.window(driver.window_handles[-1])
    logging.info(f"   → Opened solicitation detail page: {solicitation_url}")

    try:
        WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href,'javascript:downloadFile')]"))
        )
        links = driver.find_elements(By.XPATH, "//a[contains(@href,'javascript:downloadFile')]")
        logging.info(f"   → Found {len(links)} attachment links")
        for idx, link in enumerate(links, 1):
            name = link.text.strip() or f"Attachment_{idx}"
            safe_name = sanitize_filename(name)
            attachments.append({"name": safe_name})
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
                driver.execute_script("arguments[0].click();", link)
                logging.info(f"      ↳ Download triggered: {safe_name}")
                time.sleep(3)
            except Exception as e:
                logging.warning(f"      ↳ Could not download {safe_name}: {e}")
    except Exception as e:
        logging.warning(f"   → No attachments found or error: {e}")
    finally:
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

    return attachments

# --------------------------------------------------------------------
# DATA EXTRACTION
# --------------------------------------------------------------------
def get_row_data(row):
    try:
        sol_el = row.find_element(By.XPATH, "./td[1]/a")
        sol_num = sol_el.text.strip()
        sol_url = sol_el.get_attribute("href")
    except Exception:
        sol_num = sol_url = ""
    return {
        "Solicitation Number": sol_num,
        "Solicitation URL": sol_url,
        "Organization Name": get_text_safe(row, 3),
        "Buyer Name": get_text_safe(row, 6),
        "Description": get_text_safe(row, 7),
        "Awarded Vendor Name": get_text_safe(row, 10)
    }

# --------------------------------------------------------------------
# PAGE SCRAPER + PAGINATION
# --------------------------------------------------------------------
def scrape_page(driver, all_data):
    rows = driver.find_elements(By.XPATH, "//*[@id='bidSearchResultsForm:bidResultId_data']/tr")
    if not rows:
        logging.warning("   → No rows found on this page.")
    for i, row in enumerate(rows, 1):
        data = get_row_data(row)
        sol = data["Solicitation Number"]
        if not sol:
            continue
        logging.info(f"[Page Row {i}] → {sol}")
        data["Attachments"] = scrape_attachments(driver, data["Solicitation URL"])
        all_data.append(data)

        if len(all_data) % SAVE_EVERY == 0:
            with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
                json.dump(all_data, f, indent=2, ensure_ascii=False)
            logging.info(f"✅ Incremental save: {len(all_data)} records → {OUTPUT_FILE}")

def paginate_and_scrape(driver):
    all_data = []
    page_num = 1
    while True:
        logging.info(f"Processing page {page_num} …")

        # wait for first row to load
        WebDriverWait(driver, LONG_TIMEOUT).until(
            EC.presence_of_all_elements_located((By.XPATH, "//*[@id='bidSearchResultsForm:bidResultId_data']/tr"))
        )
        time.sleep(2)

        scrape_page(driver, all_data)

        # check pagination
        next_btn = driver.find_elements(
            By.XPATH,
            "//*[@id='bidSearchResultsForm:bidResultId_paginator_top']/a[contains(@class,'ui-paginator-next') and not(contains(@class,'ui-state-disabled'))]"
        )
        if next_btn:
            old_html = driver.find_element(By.ID, "bidSearchResultsForm:bidResultId_data").get_attribute("innerHTML")
            driver.execute_script("arguments[0].click();", next_btn[0])
            logging.info("Clicked next page … waiting for update.")
            WebDriverWait(driver, LONG_TIMEOUT).until(
                lambda d: d.find_element(By.ID, "bidSearchResultsForm:bidResultId_data").get_attribute("innerHTML") != old_html
            )
            page_num += 1
            time.sleep(2)
        else:
            logging.info("Reached last page.")
            break
    return all_data

# --------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------
def main():
    driver = open_browser(headless=False)
    try:
        open_advanced_search(driver)
        set_status_bid_to_po(driver)
        click_search(driver)

        # Wait for at least one result
        WebDriverWait(driver, LONG_TIMEOUT).until(
            EC.presence_of_all_elements_located((By.XPATH, "//*[@id='bidSearchResultsForm:bidResultId_data']/tr"))
        )
        time.sleep(2)

        logging.info("Scraping all pages …")
        results = paginate_and_scrape(driver)

        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        logging.info(f"Final save: {len(results)} records → {OUTPUT_FILE}")

        time.sleep(5)
    finally:
        driver.quit()

# --------------------------------------------------------------------
if __name__ == "__main__":
    main()
