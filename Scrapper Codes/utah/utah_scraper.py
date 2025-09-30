from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time


options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)

url = "https://utah.bonfirehub.com/portal/?tab=openOpportunities"
driver.get(url)

wait = WebDriverWait(driver, 20)


past_tab = wait.until(EC.element_to_be_clickable((By.ID, "pastOpportunitiesTab")))
past_tab.click()


wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#DataTables_Table_1 tbody tr")))
time.sleep(3)  

rows = driver.find_elements(By.CSS_SELECTOR, "#DataTables_Table_1 tbody tr")

data = []
for row in rows:
    try:
        cols = row.find_elements(By.TAG_NAME, "td")
        if not cols:
            continue


        status = cols[0].text.strip()

        if status.lower() == "awarded":
            ref_no = cols[1].text.strip()
            project = cols[2].text.strip()
            department = cols[3].text.strip()
            close_date = cols[4].text.strip()

            data.append({
                "Ref #": ref_no,
                "Project": project,
                "Department": department,
                "Close Date": close_date,
                "Status": status
            })
    except Exception as e:
        print("Error parsing row:", e)


df = pd.DataFrame(data)
df.to_excel("utah_awarded.xlsx", index=False)

print(f" Scraped {len(data)} awarded opportunities. Saved to utah_awarded.xlsx")


driver.quit()
