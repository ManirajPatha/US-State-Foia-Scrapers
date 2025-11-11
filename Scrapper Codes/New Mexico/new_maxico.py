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
import requests
import zipfile
from pathlib import Path

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

# Create downloads directory if it doesn't exist
downloads_dir = Path('downloads')
downloads_dir.mkdir(exist_ok=True)

data = []
total_rows = len(rows[1:])
successful_downloads = 0

print(f"Starting to process {total_rows} opportunities...")

for row_idx, row in enumerate(rows[1:], 1):
    print(f"[{row_idx}/{total_rows}] Processing opportunity...")
    
    tds = row.find_all('td')
    if len(tds) >= 2:
        status = tds[0].get_text().strip()
        details_td = tds[1]
        details_text = details_td.get_text(separator='\n').strip()
        lines = [line.strip() for line in details_text.split('\n') if line.strip()]

        # Extract PDF link - search in the entire row for the PDF link
        pdf_link = ''
        pdf_anchor = row.find('a', string=lambda text: text and 'View as PDF' in text)
        if not pdf_anchor:
            pdf_anchor = row.find('a', {'id': lambda x: x and 'BUTTON_PDF_VIEW' in x})
        if not pdf_anchor:
            pdf_anchor = row.find('a', href=lambda x: x and '.pdf' in x.lower())
        
        if pdf_anchor and pdf_anchor.has_attr('href'):
            pdf_link = pdf_anchor['href']
            print(f" PDF link found")
        else:
            print(f" No PDF link found")

        i = 0
        name_parts = []
        while i < len(lines) and lines[i] not in ['Open', 'Close', 'Type', 'Number', 'Contact']:
            name_parts.append(lines[i])
            i += 1

        name = ' '.join(name_parts).strip()
        print(f"  Name: {name[:50]}..." if len(name) > 50 else f"  Name: {name}")

        entry = {
            'Name': name,
            'Open': '',
            'Close': '',
            'Type': '',
            'Number': '',
            'Contact': '',
            'Status': status,
            'Details': '',
            'PDF_Link': pdf_link,
            'Zip_File': ''
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
        
        # Download PDF and create zip immediately for this opportunity
        if pdf_link:
            try:
                # Create a safe filename from the Name or use index
                safe_name = "".join(c for c in name if c.isalnum() or c in (' ', '-', '_')).strip()
                if not safe_name or len(safe_name) < 3:
                    safe_name = f"opportunity_{row_idx}"
                
                if len(safe_name) > 100:
                    safe_name = safe_name[:100]
                    
                pdf_filename = f"{safe_name}.pdf"
                pdf_filepath = downloads_dir / pdf_filename
                
                print(f"  Downloading PDF...", end=" ")
                response = requests.get(pdf_link, timeout=30)
                response.raise_for_status()
                
                with open(pdf_filepath, 'wb') as f:
                    f.write(response.content)
                print(f"Downloaded!")
                
                # Create individual zip file for this opportunity
                zip_filename = f"{safe_name}.zip"
                zip_filepath = downloads_dir / zip_filename
                
                print(f"  Creating zip file: {zip_filename}...", end=" ")
                with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    zipf.write(pdf_filepath, pdf_filename)
                
                
                # Remove the individual PDF file after zipping
                pdf_filepath.unlink()
                
                entry['Zip_File'] = str(zip_filepath)
                successful_downloads += 1
                print(f"Successfully processed and zipped!")
                
                time.sleep(1)
                
            except Exception as e:
                print(f"Error: {str(e)}")
                entry['Zip_File'] = f"Failed: {str(e)}"
        else:
            print(f"Skipping download (no PDF link)")
            entry['Zip_File'] = 'No PDF link available'
        
        data.append(entry)
        print()

# Save all data to Excel at the end
if data:
    df = pd.DataFrame(data)
    columns_order = ['Name', 'Open', 'Close', 'Type', 'Number', 'Contact', 'Status', 'Details', 'PDF_Link', 'Zip_File']
    df = df[columns_order]
    df.to_excel('awarded_bids.xlsx', index=False)
    
    print(f"✓ Data saved to 'awarded_bids.xlsx'")
    print(f"Total opportunities processed: {len(df)}")
    print(f"Opportunities with PDF links: {df['PDF_Link'].astype(bool).sum()}")
    print(f"Successfully downloaded and zipped: {successful_downloads}")
    print(f"All zip files saved in: {downloads_dir}")
        
else:
    print("No data rows found in the table.")