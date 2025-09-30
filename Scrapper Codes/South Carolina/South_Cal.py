from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import os
import logging
import re

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
    driver.get("https://apps.sceis.sc.gov/SCSolicitationWeb/solicitationSearch.do")

    logging.info("Selecting 'Closed' status")
    closed_radio = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='searchStatus'][value='C']"))
    )
    closed_radio.click()

    logging.info("Selecting '200' search limit")
    search_limit_dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "searchLimit"))
    )
    search_limit_dropdown.click()
    limit_option = driver.find_element(By.CSS_SELECTOR, "select[name='searchLimit'] option[value='200']")
    limit_option.click()

    logging.info("Clicking Search button")
    search_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='btnSearch']"))
    )
    search_button.click()

    time.sleep(3)

    all_data = []

    while True:
        logging.info("Scraping current page")
        rows = driver.find_elements(By.CSS_SELECTOR, "table tr")
        
        for row in rows:
            try:
                header_cells = row.find_elements(By.TAG_NAME, "th")
                cells = row.find_elements(By.TAG_NAME, "td")
                if header_cells or len(cells) < 4:
                    continue
                
                first_cell_text = cells[0].text.strip()
                
                if not re.match(r'^\d{8,}$', first_cell_text):
                    continue
                solicitation_number = first_cell_text
                description = cells[1].text.strip()
                agency = cells[2].text.strip()
                submission_date = cells[3].text.strip()
                
                all_data.append({
                    'Solicitation Number': solicitation_number,
                    'Solicitation Description': description,
                    'Purchasing Agency': agency,
                    'Submission Ending Date/Time': submission_date
                })
            except Exception as e:
                logging.warning(f"Error parsing row: {e}")
                continue


        try:
            next_button = driver.find_element(By.LINK_TEXT, "Next")
            if "disabled" not in next_button.get_attribute("class") and next_button.is_enabled():
                logging.info("Clicking Next button")
                next_button.click()
                time.sleep(3)
            else:
                logging.info("No more pages available")
                break
        except:
            logging.info("Next button not found - end of pages")
            break


    if all_data:
        df = pd.DataFrame(all_data)
        output_file = os.path.join(os.getcwd(), "SC_Closed_Solicitations.xlsx")
        df.to_excel(output_file, index=False, sheet_name='Solicitations')
        logging.info(f"Data scraped successfully. Excel file saved: {output_file}")
        logging.info(f"Total records scraped: {len(all_data)}")
    else:
        logging.warning("No data scraped")

except Exception as e:
    logging.error(f"An error occurred: {e}")
finally:
    
    logging.info("Closing browser")
    driver.quit()