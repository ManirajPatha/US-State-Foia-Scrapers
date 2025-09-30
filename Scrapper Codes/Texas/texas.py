import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


def scrape_and_convert(download_dir="downloads"):
    os.makedirs(download_dir, exist_ok=True)

    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": os.path.abspath(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(driver, 20)

    try:
        url = "https://www.txsmartbuy.gov/esbd?status=2&dateRange=lastFiscalYear&startDate=09%2F01%2F2024&endDate=08%2F31%2F2025&page=160"
        driver.get(url)

        status_dropdown = Select(wait.until(EC.presence_of_element_located((By.NAME, "status"))))
        status_dropdown.select_by_value("2")

        date_range_dropdown = Select(wait.until(EC.presence_of_element_located((By.NAME, "dateRange"))))
        date_range_dropdown.select_by_value("lastFiscalYear")

        search_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))
        search_btn.click()
        time.sleep(5)  

        export_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-action='export-csv']")))
        export_btn.click()

        time.sleep(10)

        csv_file = None
        for file in os.listdir(download_dir):
            if file.endswith(".csv"):
                csv_file = os.path.join(download_dir, file)
                break

        if not csv_file:
            raise Exception("CSV file not downloaded")

        excel_file = csv_file.replace(".csv", ".xlsx")
        df = pd.read_csv(csv_file)
        df.to_excel(excel_file, index=False)

        print(f"Excel file saved: {excel_file}")

    finally:
        driver.quit()


if __name__ == "__main__":
    scrape_and_convert()
