import time
from datetime import date
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Path to your input Excel file
INPUT_XLSX = "/home/developer/Desktop/Scraper/vermont_awards_data.xlsx"
# Path for the output Excel file
OUTPUT_XLSX = "/home/developer/Desktop/Scraper/output_submission_results.xlsx"
# URL of the public-records request form
FORM_URL = "https://ago.vermont.gov/public-records-request-form"

def read_input_data(path):
    """Read Project Title and Award Date from Excel, parse and format dates."""
    df = pd.read_excel(path, usecols=["Project Title", "Award Date"])
    # Parse Award Date column as datetime, coerce errors
    df["Award Date"] = pd.to_datetime(df["Award Date"], errors="coerce")
    # Format dates back to MM/DD/YYYY strings, fill missing
    df["Award Date"] = df["Award Date"].dt.strftime("%m/%d/%Y").fillna("")
    return df.to_dict(orient="records")

def setup_driver():
    """Create and return a Chrome WebDriver."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # options.add_argument("--headless=new")  # Uncomment to run headless
    return webdriver.Chrome(options=options)

def fill_and_submit(driver, wait, project_title, award_date):
    """Fill the form fields for one record and submit. Returns success message or error."""
    driver.get(FORM_URL)
    wait.until(EC.presence_of_element_located((By.ID, "edit-first-name")))

    # Fill text fields
    driver.find_element(By.ID, "edit-first-name").send_keys("Raaj")
    driver.find_element(By.ID, "edit-last-name").send_keys("Thipparthy")
    driver.find_element(By.ID, "edit-email-mail-1").send_keys("raajnrao@gmail.com")
    driver.find_element(By.ID, "edit-email-mail-2").send_keys("raajnrao@gmail.com")
    driver.find_element(By.ID, "edit-phone-number").send_keys("8325197135")

    # Fill description
    textarea = driver.find_element(
        By.ID,
        "edit-please-describe-the-records-you-are-requesting-and-provide-as-mu"
    )
    text = (
        f"I am requesting a copy of the winning and shortlisted proposals for {project_title}. "
        f"I am requesting a copy of the winning and shortlisted proposals for the referenced award. "
        f"The solicitation/contract number is {award_date}."
    )
    textarea.send_keys(text)

    # Check Declaration
    driver.find_element(By.XPATH, "//label[@for='edit-declaration-required-']").click()

    # Fill today's date
    today = date.today().strftime("%m/%d/%Y")  # MM/DD/YYYY
    driver.find_element(By.ID, "edit-date-submitted").send_keys(today)

    # Wait for manual CAPTCHA
    print("Complete CAPTCHA then press Enter to continue...")
    input()

    # Submit
    driver.find_element(By.ID, "edit-submit").click()

    # Capture confirmation
    try:
        success_elem = wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, ".webform-confirmation-message, .messages--status")
        ))
        return success_elem.text.strip()
    except:
        return "No confirmation detected"

def automate_requests(input_path, output_path):
    records = read_input_data(input_path)
    driver = setup_driver()
    wait = WebDriverWait(driver, 30)

    results = []
    for rec in records:
        project_title = rec["Project Title"]
        award_date = rec["Award Date"]
        print(f"Submitting for: {project_title} / {award_date}")
        try:
            msg = fill_and_submit(driver, wait, project_title, award_date)
        except Exception as e:
            msg = f"Error: {e}"
        results.append({
            "Project Title": project_title,
            "Award Date": award_date,
            "Success Msg": msg
        })

    driver.quit()

    # Save results
    df_out = pd.DataFrame(results)
    df_out.to_excel(output_path, index=False)
    print(f"Results saved to {output_path}")

if __name__ == "__main__":
    automate_requests(INPUT_XLSX, OUTPUT_XLSX)