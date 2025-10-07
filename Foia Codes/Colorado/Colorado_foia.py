import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


# ==================================================
# Setup Selenium Driver
# ==================================================
def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


# ==================================================
# Submit one FOIA request
# ==================================================
def submit_colorado_foia(driver, solicitation_number, description):
    try:
        url = "https://dhsem.colorado.gov/form/cora-request-form"
        driver.get(url)

        wait = WebDriverWait(driver, 20)

        # Fill Name
        name_input = wait.until(EC.presence_of_element_located((By.ID, "edit-name")))
        name_input.clear()
        name_input.send_keys("Raaj Thipparthy")

        # Fill Email
        email_input = driver.find_element(By.ID, "edit-email-address")
        email_input.clear()
        email_input.send_keys("raajnrao@gmail.com")

        # Fill Information Requested
        info_input = driver.find_element(By.ID, "edit-request")
        request_text = (
            f"I am requesting a copy of the winning and shortlisted proposals "
            f"for {solicitation_number}. I am requesting a copy of the winning "
            f"and shortlisted proposals for the referenced award. "
            f"The solicitation/contract number is {description}."
        )
        info_input.clear()
        info_input.send_keys(request_text)

        # Always choose "No" for Personal Gain Statement
        no_label = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "label[for='edit-personal-gain-statement-crs-24-72-305-5-no']")
        ))
        driver.execute_script("arguments[0].click();", no_label)

        print("⚠️ Please solve the CAPTCHA manually...")
        input("Press ENTER after completing the CAPTCHA...")

        # Submit form
        submit_button = driver.find_element(By.ID, "edit-actions-submit")
        driver.execute_script("arguments[0].click();", submit_button)

        # Wait briefly for success page
        time.sleep(3)

        return "Submitted successfully ✅"

    except Exception as e:
        return f"Error: {str(e)}"


# ==================================================
# Main Script
# ==================================================
if __name__ == "__main__":
    print("Starting Colorado CORA Form Automation")
    print("==================================================")

    # Load Excel file
    input_file = "/home/developer/Desktop/Scraper/colorado_vss_recent_awards.xlsx"   # <-- Replace with your file path
    df = pd.read_excel(input_file)

    results = []

    driver = setup_driver()

    for i, row in df.iterrows():
        print(f"\nProcessing row {i+1}/{len(df)}")
        solicitation = row["Solicitation Number"]
        desc = row["Description"]
        print(f"Solicitation Number: {solicitation}")
        print(f"Description: {desc}")

        result = submit_colorado_foia(driver, solicitation, desc)
        print(f"Result: {result}")

        results.append({
            "Solicitation Number": solicitation,
            "Description": desc,
            "Success Msg": result
        })

    driver.quit()

    # Save results to Excel
    output_file = "Colorado_FOIA_Results.xlsx"
    pd.DataFrame(results).to_excel(output_file, index=False)
    print(f"\n✅ Results saved to {output_file}")
