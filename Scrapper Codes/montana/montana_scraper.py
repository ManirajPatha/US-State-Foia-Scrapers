from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import time
import os

CHROME_BINARY_PATH = "/usr/bin/google-chrome"

if not os.path.exists(CHROME_BINARY_PATH):
    print(f"Chrome binary not found at {CHROME_BINARY_PATH}. Please install Google Chrome or update the binary path.")
    exit()

chrome_options = Options()
# chrome_options.add_argument("--headless")  # uncomment if you want headless
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.binary_location = CHROME_BINARY_PATH

try:
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    print("ChromeDriver initialized successfully.")
except Exception as e:
    print(f"Error setting up ChromeDriver: {e}")
    exit()

initial_url = "https://bids.sciquest.com/apps/Router/PublicEvent?CustomerOrg=StateOfMontana"
try:
    driver.get(initial_url)
    print("Opened the initial URL.")
    time.sleep(3)
except Exception as e:
    print(f"Error loading initial URL: {e}")
    driver.quit()
    exit()

# Click the "Awarded" tab
try:
    awarded_tab = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "PhoenixNavLink_PHX_NAV_SourcingAward"))
    )
    awarded_tab.click()
    print("Clicked the 'Awarded' tab.")
    time.sleep(5)
except Exception as e:
    print(f"Error clicking the 'Awarded' tab: {e}")
    driver.quit()
    exit()

all_data = []
page_number = 1

while True:
    print(f"Scraping page {page_number}...")
    page_source = driver.page_source
    soup = BeautifulSoup(page_source, 'html.parser')

    table = soup.find('table')
    if not table:
        print("No table found on this page. Skipping...")
        break

    rows = table.find_all('tr')
    for row in rows[1:]:
        tds = row.find_all('td')
        if len(tds) >= 2:
            status = tds[0].get_text().strip()
            details_td = tds[1]
            details_text = details_td.get_text(separator='\n').strip()
            lines = [line.strip() for line in details_text.split('\n') if line.strip()]

            i = 0
            name_parts = []
            while i < len(lines) and lines[i] not in ['Open', 'Close', 'Type', 'Number', 'Contact']:
                name_parts.append(lines[i])
                i += 1

            name = ' '.join(name_parts).strip()

            entry = {
                'Name': name,
                'Open': '',
                'Close': '',
                'Type': '',
                'Number': '',
                'Contact': '',
                'Status': status,
                'Details': ''
            }

            extra_details = []
            while i < len(lines):
                key = lines[i]
                i += 1
                value = ''
                if i < len(lines) and lines[i] not in ['Open', 'Close', 'Type', 'Number', 'Contact']:
                    value = lines[i]
                    i += 1

                if key in entry:
                    entry[key] = value
                else:
                    extra_details.append(f"{key}: {value}")

            entry['Details'] = '\n'.join(extra_details)
            all_data.append(entry)

    # Check for Next button
    try:
        next_button = driver.find_element(By.XPATH, "//button[@aria-label='Next page' and not(@disabled)]")
        # Avoid duplicate scraping by checking first row
        first_row_name = rows[1].find_all('td')[1].get_text().strip() if len(rows) > 1 else ''
        driver.execute_script("arguments[0].click();", next_button)
        # Wait for the table to refresh (first row changes)
        WebDriverWait(driver, 10).until(
            lambda d: d.find_element(By.XPATH, "//table//tr[2]/td[2]").text.strip() != first_row_name
        )
        page_number += 1
        time.sleep(2)
    except:
        print("No more pages or Next button disabled.")
        break

driver.quit()

# Save all data to Excel
if all_data:
    df = pd.DataFrame(all_data)
    columns_order = ['Name', 'Open', 'Close', 'Type', 'Number', 'Contact', 'Status', 'Details']
    df = df[columns_order]
    df.to_excel('awarded_bids_all_pages.xlsx', index=False)
    print(f"Data scraped from {page_number} pages and saved to 'awarded_bids_all_pages.xlsx'.")
else:
    print("No data found.")
