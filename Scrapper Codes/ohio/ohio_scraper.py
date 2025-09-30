from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

driver = webdriver.Chrome()
wait = WebDriverWait(driver, 60)

driver.get("https://ohiobuys.ohio.gov/page.aspx/en/rfp/request_browse_public")


begin_date_input = wait.until(EC.element_to_be_clickable((By.ID, "body_x_txtRfpBeginDate")))
begin_date_input.clear()
begin_date_input.send_keys("9/2/2025")

time.sleep(1)  


date_picker_date = wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//div[contains(@class, 'ui-datepicker')]//a[text()='2']")
))
ActionChains(driver).move_to_element(date_picker_date).click().perform()

wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "ui-datepicker")))


dropdown = wait.until(EC.element_to_be_clickable((By.ID, "body_x_cbRfpPubAward_search")))
dropdown.click()

yes_option = wait.until(EC.element_to_be_clickable((By.ID, "body_x__ctl1_True")))
yes_option.click()


search_button = wait.until(EC.element_to_be_clickable((By.ID, "body_x_prxFilterBar_x_cmdSearchBtn")))
search_button.click()


rows = wait.until(EC.presence_of_all_elements_located(
    (By.XPATH, "//tr[starts-with(@id,'body_x_grid_grd_tr_')]")
))

data = []
for row in rows:
    cols = row.find_elements(By.XPATH, ".//td")
    row_data = [col.text.strip() for col in cols]
    data.append(row_data)

columns = [
    "Edit", "Solicitation ID", "Solicitation Name",
    "Original Begin Date (ET.)", "Begin Date (ET.)", "End Date (ET.)",
    "Inquiry End Date (ET.)", "Commodity", "MBE Set Aside",
    "Agency", "Hidden Col1", "Hidden Col2", "Solicitation Status",
    "Awarded", "Solicitation Type"
]

df = pd.DataFrame(data, columns=columns[:len(data[0])])
df.to_excel("ohio_buys_awarded.xlsx", index=False)
print("Data saved to ohio_buys_awarded.xlsx")

driver.quit()
