from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

MAX_ROWS = 100  


options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
driver.get("https://www.njstart.gov/bso/view/search/external/advancedSearchBid.xhtml")

wait = WebDriverWait(driver, 30)


status_dropdown = wait.until(EC.presence_of_element_located((By.ID, "bidSearchForm:status")))
select = Select(status_dropdown)
select.select_by_value("2BPO")  


search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Search']")))
search_btn.click()


wait.until(EC.presence_of_element_located((By.ID, "bidSearchResultsForm:bidResultId_head")))
time.sleep(2)  


headers = driver.find_elements(By.XPATH, "//thead[@id='bidSearchResultsForm:bidResultId_head']//th[not(@style='display:none')]/span[@class='ui-column-title']")
columns = [h.text.strip() for h in headers if h.text.strip()]

all_data = []

while True:
    
    if len(all_data) >= MAX_ROWS:
        break

    
    rows = driver.find_elements(By.XPATH, "//tbody[@id='bidSearchResultsForm:bidResultId_data']/tr")
    for row in rows:
        if len(all_data) >= MAX_ROWS:
            break
        cells = row.find_elements(By.XPATH, "./td[not(@style='display:none')]")
        row_data = [cell.text.strip() for cell in cells]
        all_data.append(row_data)

    
    try:
        next_btn = driver.find_element(By.XPATH, "//a[contains(@class,'ui-paginator-next')]")
        if "ui-state-disabled" in next_btn.get_attribute("class"):
            break  
        next_btn.click()
        time.sleep(2)  
    except:
        break  


df = pd.DataFrame(all_data[:MAX_ROWS], columns=columns)  


df.to_excel("njstart_bid_to_po_all.xlsx", index=False)

print(f"Scraped {len(df)} rows (capped at {MAX_ROWS}) and saved to njstart_bid_to_po_all.xlsx")


driver.quit()
