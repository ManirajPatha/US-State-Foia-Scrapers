import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# ====== Load Excel Data ======
input_file = "/home/developer/Desktop/Scraper/Minnesota_Final_Bids_20Pages.xlsx"
df = pd.read_excel(input_file)

# Add column for results
df["Success Msg"] = ""

# ====== Setup WebDriver ======
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 20)

# ====== Constants ======
URL = "https://mn.gov/admin/osp/contact-us/data-requests/?utm_source=chatgpt.com"
NAME = "Raaj Thipparthy"
EMAIL = "raajnrao@gmail.com"
PHONE = "8325197135"

# ====== Loop over each row in Excel ======
for idx, row in df.iterrows():
    quest_number = str(row["Quest Number"]).strip()
    bid_name = str(row["Bid/Request Name"]).strip()

    try:
        driver.get(URL)

        # Fill "Your Name"
        wait.until(EC.presence_of_element_located((By.ID, "field130399080"))).send_keys(NAME)

        # Fill "Email address"
        driver.find_element(By.ID, "field130399096").send_keys(EMAIL)

        # Fill "Confirm Email"
        driver.find_element(By.ID, "field130399111").send_keys(EMAIL)

        # Fill "Telephone Number"
        driver.find_element(By.ID, "field130399154").send_keys(PHONE)

        # Fill textarea with dynamic text
        text_area = driver.find_element(By.ID, "field130399587")
        request_text = (
            f"I am requesting a copy of the winning and shortlisted proposals for {bid_name}. "
            f"I am requesting a copy of the winning and shortlisted proposals for the referenced award. "
            f"The solicitation/contract number is {quest_number}."
        )
        text_area.send_keys(request_text)

        # Submit the form
        submit_btn = driver.find_element(By.ID, "fsSubmitButton4957435")
        driver.execute_script("arguments[0].click();", submit_btn)

        # Wait for confirmation (success message)
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(),'Thank you')]")))
            df.at[idx, "Success Msg"] = "Submitted"
        except TimeoutException:
            df.at[idx, "Success Msg"] = "Form submission uncertain"

    except Exception as e:
        print(f"Error on Quest Number {quest_number}: {e}")
        df.at[idx, "Success Msg"] = f"Failed: {e}"

# ====== Save Results ======
output_file = "Minnesota_FOIA_Submission_Results.xlsx"
df.to_excel(output_file, index=False)

print(f"✅ Automation complete! Results saved to {output_file}")

driver.quit()
