from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

chrome_options = webdriver.ChromeOptions()
download_dir = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(download_dir):
    os.makedirs(download_dir)
    os.chmod(download_dir, 0o755)

chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

try:
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    logging.info("WebDriver initialized successfully")
except Exception as e:
    logging.error(f"Failed to initialize WebDriver: {e}")
    raise

try:
    logging.info("Navigating to website")
    driver.get("https://evp.nc.gov/solicitations/")

    logging.info("Clicking filter header")
    filter_header = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "filter-header"))
    )
    filter_header.click()

    logging.info("Selecting 'Last Year' in Posted Date")
    posted_date_dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "dropdownfilter_2"))
    )
    posted_date_dropdown.click()
    last_year_option = driver.find_element(By.XPATH, "//select[@id='dropdownfilter_2']/option[@value='5']")
    last_year_option.click()

    logging.info("Unchecking 'Open' in Solicitation Status")
    open_checkbox = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@name='3' and @value='0']"))
    )
    if open_checkbox.is_selected():
        open_checkbox.click()

    logging.info("Checking 'Awarded' in Solicitation Status")
    awarded_checkbox = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@name='3' and @value='2']"))
    )
    if not awarded_checkbox.is_selected():
        awarded_checkbox.click()

    logging.info("Clicking Apply button")
    apply_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn-entitylist-filter-submit"))
    )
    apply_button.click()

    time.sleep(3)

    logging.info("Clicking Export Results button")
    export_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.entitylist-download.btn.btn-info.pull-right.action"))
    )
    export_button.click()

    logging.info("Waiting for file download")
    time.sleep(10)

    downloaded_files = [f for f in os.listdir(download_dir) if f.endswith('.xlsx') or f.endswith('.xls')]
    if downloaded_files:
        logging.info(f"Excel file downloaded: {downloaded_files[0]} to {download_dir}")
    else:
        logging.warning("No Excel file found in download directory")

except Exception as e:
    logging.error(f"An error occurred: {e}")
finally:
    logging.info("Closing browser")
    driver.quit()