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
chrome_options.add_argument("--headless")
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

initial_url = "https://bids.sciquest.com/apps/Router/PublicEvent?CustomerOrg=StateOfNewMexico&tab=PHX_NAV_SourcingOpenForBid&tmstmp=1467214109161"
try:
    driver.get(initial_url)
    print("Opened the initial URL.")
    time.sleep(3)
except Exception as e:
    print(f"Error loading initial URL: {e}")
    driver.quit()
    exit()

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

try:
    page_source = driver.page_source
    print("Retrieved page source successfully.")
except Exception as e:
    print(f"Error retrieving page source: {e}")
    driver.quit()
    exit()

driver.quit()

soup = BeautifulSoup(page_source, 'html.parser')

table = soup.find('table')
if not table:
    print("No table found on the page. Check the page structure or URL.")
    exit()


rows = table.find_all('tr')

data = []
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
        data.append(entry)

if data:
    df = pd.DataFrame(data)
    columns_order = ['Name', 'Open', 'Close', 'Type', 'Number', 'Contact', 'Status', 'Details']
    df = df[columns_order]
    df.to_excel('awarded_bids.xlsx', index=False)
    print("Data scraped and saved to 'awarded_bids.xlsx'.")
else:
    print("No data rows found in the table.")