from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time


driver = webdriver.Chrome()
driver.get("https://mvendor.cgieva.com/Vendor/public/AllOpportunities.jsp")

wait = WebDriverWait(driver, 30)


awarded_filter = wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//li[contains(@class,'facet-item-type-status')][contains(text(),'Awarded')]")
))
driver.execute_script("arguments[0].click();", awarded_filter)
time.sleep(3)  

cards = driver.find_elements(By.CSS_SELECTOR, "div.card")
scroll_pause = 5
while len(cards) < 500:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(scroll_pause)
    cards = driver.find_elements(By.CSS_SELECTOR, "div.card")


data = []
for card in cards[:100]:
    try:
        
        title_elem = card.find_element(By.CSS_SELECTOR, "h5.card-title")
        title = title_elem.text.strip()

        
        try:
            sol_id = card.find_element(By.CSS_SELECTOR, "h6.card-title").text.strip()
        except:
            sol_id = ""

        
        try:
            status = card.find_element(By.CSS_SELECTOR, "div.statusdisplayAllOps span b").text.strip()
        except:
            status = ""

        
        try:
            close_date_elem = card.find_element(By.XPATH, ".//p[contains(text(),'Closed On')]")
            closing_date = close_date_elem.text.replace("Closed On:", "").strip()
        except:
            closing_date = ""

        data.append([title, sol_id, status, closing_date])
    except Exception as e:
        print("Error parsing card:", e)


df = pd.DataFrame(data, columns=["Title", "Solicitation ID", "Status", "Closing Date"])
df.to_excel("eva_awarded_opportunities_100.xlsx", index=False)

print(f" Scraped {len(df)} records. Data saved to eva_awarded_opportunities_100.xlsx")

driver.quit()
