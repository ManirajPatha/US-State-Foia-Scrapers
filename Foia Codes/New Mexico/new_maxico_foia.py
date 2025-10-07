import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

INPUT_FILE = "/home/developer/Desktop/Scraper/New_Maxico_awarded_bids.xlsx"
OUTPUT_FILE = "NM_RLD_Form_Submissions.xlsx"

def fill_form(name, number, driver, wait):
    """Fill the form for one Name & Number pair"""

    try:
        driver.get("https://www.rld.nm.gov/about-us/public-information-hub/inspection-of-public-records/")
        wait.until(EC.presence_of_element_located((By.ID, "input_8_1")))

        today = datetime.today().strftime("%m/%d/%Y")
        date_input = driver.find_element(By.ID, "input_8_1")
        date_input.clear()
        date_input.send_keys(today)

        driver.find_element(By.ID, "input_8_4_3").send_keys("Raaj")
        driver.find_element(By.ID, "input_8_4_6").send_keys("Thipparthy")

        driver.find_element(By.ID, "input_8_5_1").send_keys("8181 Fannin St")
        driver.find_element(By.ID, "input_8_5_3").send_keys("Houston")
        state_dropdown = driver.find_element(By.ID, "input_8_5_4")
        for option in state_dropdown.find_elements(By.TAG_NAME, "option"):
            if option.text.strip() == "Texas":
                option.click()
                break
        driver.find_element(By.ID, "input_8_5_5").send_keys("77054")

        driver.find_element(By.ID, "input_8_6").send_keys("8325197135")

        driver.find_element(By.ID, "input_8_7").send_keys("raajnrao@gmail.com")

        driver.find_element(By.ID, "choice_8_8_1").click()

        textarea = driver.find_element(By.ID, "input_8_9")
        text_value = f"I am requesting a copy of the winning and shortlisted proposals for {name}. " \
                     f"I am requesting a copy of the winning and shortlisted proposals for the referenced award. " \
                     f"The solicitation/contract number is {number}."
        textarea.send_keys(text_value)

        driver.find_element(By.ID, "choice_8_10_0").click()

        print(f"⚠️ Please solve the CAPTCHA manually for {name}, {number}...")
        input("Press ENTER after solving CAPTCHA...")

        submit_btn = driver.find_element(By.ID, "gform_submit_button_8")
        submit_btn.click()

        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "gform_confirmation_message")))
        return "Form Submitted Successfully"

    except Exception as e:
        return f"Failed: {str(e)}"


def main():
    df = pd.read_excel(INPUT_FILE)

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 40)

    results = []

    for _, row in df.iterrows():
        name = str(row.get("Name", "")).strip()
        number = str(row.get("Number", "")).strip()

        if not name or not number:
            continue

        print(f"Processing: Name={name}, Number={number}")
        msg = fill_form(name, number, driver, wait)
        results.append({"Name": name, "Number": number, "Success Msg": msg})
        time.sleep(3)

    driver.quit()

    results_df = pd.DataFrame(results)
    results_df.to_excel(OUTPUT_FILE, index=False)
    print(f"Process completed. Results saved in {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
