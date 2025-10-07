#!/usr/bin/env python3
"""
South Carolina DPS FOIA Auto-Fill Script
----------------------------------------
- Reads Excel file with columns: 'Solicitation Description', 'Solicitation Number'
- Auto-fills FOIA request form
- Logs success/failure into output Excel

Requirements:
    pip install selenium pandas openpyxl
"""

import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ---------- CONFIG ----------
INPUT_EXCEL = r"C:\Users\ADMIN\Desktop\Scraper\South_Cal_Closed_Solicitations.xlsx"   # <-- change this to your Excel file
OUTPUT_EXCEL = r"C:\Users\ADMIN\Desktop\Scraper\South_Cal_dps_foia_results.xlsx"
BASE_URL = ("https://sceisimage.sc.gov/appnetdps/UnityForm.aspx?d1="
            "AYoB1mXTJm8uTasLsQij%2b13zuxByM6Lxjec9WupAvHF0GmJXyF984H9bLubh0gQtkxtnhg2tG4n9Er9pwijxnbd0AzyCe%2bDAZrQPECYk3MDisDyhrccklWX5RSjnnFii652yy40%2f%2bJu0RVLysVkG5XkDTF4I7mz9AVJyYr3QcXKX%2bB9c3VxwTBWq%2bN91axRzSR0ucVhptbzgXV1IscMtIOA%3d")

# Fixed personal info fields
FIXED_FIELDS = {
    "name14_input": "Raaj Thipparthy",
    "address15_input": "8181 Fannin St",
    "city17_input": "Houston",
    "zipcode19_input": "77054",
    "phone20_input": "8325197135",
    "email22_input": "raajnrao@gmail.com",
    "prrequestshortdescription76_input": "Requesting a copy of the winning and shortlisted proposals",
}
STATE_VALUE = "TX"
MOTOR_VEHICLE_VALUE = "No"

# ---------- SELENIUM SETUP ----------
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

# Load Excel
df = pd.read_excel(INPUT_EXCEL)
df["Success Msg"] = ""


def switch_to_form_iframe():
    """Switch into the form iframe."""
    driver.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe")))


def fill_static_fields():
    """Fill all fixed fields once per form load."""
    for field_id, value in FIXED_FIELDS.items():
        box = wait.until(EC.presence_of_element_located((By.ID, field_id)))
        box.clear()
        box.send_keys(value)

    # Select State
    state_box = driver.find_element(By.ID, "state18_input")
    state_box.clear()
    state_box.send_keys(STATE_VALUE)
    time.sleep(0.5)

    # Motor vehicle dropdown
    mv_box = driver.find_element(By.ID, "motorvehicle78_input")
    mv_box.clear()
    mv_box.send_keys(MOTOR_VEHICLE_VALUE)
    time.sleep(0.5)


# ---------- MAIN LOOP ----------
for idx, row in df.iterrows():
    try:
        # Open form fresh each time
        driver.get(BASE_URL)
        switch_to_form_iframe()
        fill_static_fields()

        solicitation_desc = str(row["Solicitation Description"])
        solicitation_num = str(row["Solicitation Number"])

        # Fill long description field
        textarea = wait.until(EC.presence_of_element_located((By.ID, "multilinetextbox76_input")))
        textarea.clear()
        long_text = (
            f"I am requesting a copy of the winning and shortlisted proposals for {solicitation_desc}. "
            f"I am requesting a copy of the winning and shortlisted proposals for the referenced award. "
            f"The solicitation/contract number is {solicitation_num}."
        )
        textarea.send_keys(long_text)

        # Submit
        submit_btn = driver.find_element(By.XPATH, "//div[@id='submit']//input[@type='submit']")
        submit_btn.click()

        # Wait a little for confirmation (could improve if page shows message)
        time.sleep(5)

        if "UnityForm.aspx" not in driver.current_url:
            df.at[idx, "Success Msg"] = "Submitted Successfully"
        else:
            df.at[idx, "Success Msg"] = "Submission May Have Failed"

    except TimeoutException:
        df.at[idx, "Success Msg"] = "Timeout waiting for field"
    except Exception as e:
        df.at[idx, "Success Msg"] = f"Error: {e}"

# Save output Excel
df.to_excel(OUTPUT_EXCEL, index=False)
print(f"✅ Done. Results saved to {OUTPUT_EXCEL}")

driver.quit()
 