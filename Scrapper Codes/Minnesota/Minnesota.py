from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time


options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

url = "https://qcpi.questcdn.com/cdn/results/?group=6506969&provider=6506969&projType=all"
driver.get(url)

wait = WebDriverWait(driver, 10)
bid_award_input = wait.until(EC.presence_of_element_located((By.NAME, "col9_search")))
bid_award_input.clear()
bid_award_input.send_keys("Final")
bid_award_input.send_keys(Keys.RETURN)
time.sleep(3)

all_data = []

def scrape_page():
    rows = driver.find_elements(By.XPATH, "//table[@id='table_id']/tbody/tr")
    for row in rows:
        cols = row.find_elements(By.TAG_NAME, "td")
        row_data = [col.text.strip() for col in cols]
        all_data.append(row_data)

page_count = 0
max_pages = 20

while page_count < max_pages:
    scrape_page()
    page_count += 1
    
    try:
        next_button = driver.find_element(By.XPATH, "//a[contains(@class,'page-link') and text()='Next']")
        driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", next_button)
        time.sleep(3)
    except NoSuchElementException:
        print(f"No more pages to scrape. Stopped at page {page_count}.")
        break

driver.quit()

columns = ["Quest Number", "Bid/Request Name", "Bid Closing Date", "City", "County", 
           "State", "Owner", "Solicitor", "Posting Type", "Bid Award Type"]

df = pd.DataFrame(all_data, columns=columns)
df.to_excel("QuestCDN_Final_Bids_20Pages.xlsx", index=False)

print(f"Scraping complete! Total rows: {len(all_data)}")
print(f"Scraped {page_count} pages. Data saved to 'QuestCDN_Final_Bids_20Pages.xlsx'")