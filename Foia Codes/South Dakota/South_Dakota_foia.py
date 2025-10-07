import time
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ------------------ CONFIG ------------------
URL = "https://www.sd.gov/cs?id=cs_guided_cat_item&sys_id=f7f939eddbd4b150b2fb93d4f39619c0"
EXCEL_INPUT = r"C:\Users\ADMIN\Desktop\Scraper\South_Dakota_Awarded_Opportunities.xlsx"
EXCEL_OUTPUT = r"C:\Users\ADMIN\Desktop\Scraper\South_Dakota_FOIA_Submissions.xlsx"

FIRST_NAME = "Raaj"
LAST_NAME = "Thipparthy"
EMAIL = "raajnrao@gmail.com"
AGENCY_NAME = "Bureau of Administration"  # Always select this agency

# ------------------ FUNCTIONS ------------------
def get_chrome_options():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    possible_binaries = [
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
        "/usr/bin/chromium-browser",
        "/usr/bin/chromium",
        "/usr/bin/brave-browser",
    ]

    for binary in possible_binaries:
        if shutil.which(binary):
            options.binary_location = binary
            print(f"✅ Using browser binary: {binary}")
            break
    else:
        print("⚠️ No Chrome/Chromium/Brave binary found. Please install one.")
    return options


def wait_for_manual_login(driver):
    """Pause execution until user logs in manually and main form is visible."""
    print("⏳ Please log in manually in the opened browser...")
    wait = WebDriverWait(driver, 300)  # Wait up to 5 minutes
    wait.until(EC.presence_of_element_located((By.ID, "sp_formfield_first_name")))
    print("✅ Login detected, proceeding with automation.")


def select_agency(driver, wait):
    """Select the agency from the dropdown."""
    # Click the Select2 wrapper
    agency_wrapper = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.select2-choice")))
    driver.execute_script("arguments[0].click();", agency_wrapper)
    time.sleep(1)

    # Type the agency name
    agency_input = driver.find_element(By.CSS_SELECTOR, "input.select2-focusser")
    agency_input.send_keys(AGENCY_NAME)
    time.sleep(1)

    # Wait for the dropdown result and click it
    agency_result = wait.until(EC.element_to_be_clickable((By.XPATH,
        f"//div[@class='select2-result-label' and text()='{AGENCY_NAME}']"
    )))
    agency_result.click()
    time.sleep(0.5)


def fill_static_fields(driver, wait):
    """Fill first name, last name, email, and always select the desired agency."""

    # First Name
    first_name = wait.until(EC.visibility_of_element_located((By.ID, "sp_formfield_first_name")))
    driver.execute_script("arguments[0].scrollIntoView(true);", first_name)
    time.sleep(0.5)
    first_name.clear()
    first_name.send_keys(FIRST_NAME)

    # Last Name
    last_name = wait.until(EC.visibility_of_element_located((By.ID, "sp_formfield_last_name")))
    driver.execute_script("arguments[0].scrollIntoView(true);", last_name)
    time.sleep(0.5)
    last_name.clear()
    last_name.send_keys(LAST_NAME)

    # Email
    email = wait.until(EC.visibility_of_element_located((By.ID, "sp_formfield_email")))
    driver.execute_script("arguments[0].scrollIntoView(true);", email)
    time.sleep(0.5)
    email.clear()
    email.send_keys(EMAIL)

    # Agency
    select_agency(driver, wait)


def process_events():
    df = pd.read_excel(EXCEL_INPUT)
    results = []

    options = get_chrome_options()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(URL)
    wait = WebDriverWait(driver, 20)

    # Wait for manual login
    wait_for_manual_login(driver)

    try:
        fill_static_fields(driver, wait)

        for idx, row in df.iterrows():
            event_id = str(row["Event Name"])
            event_name = str(row["Event ID"])

            try:
                # Fill the FOIA request textarea
                textarea = wait.until(EC.visibility_of_element_located(
                    (By.ID, "sp_formfield_what_specific_record_are_you_requesting_for_disclosure_record_request")
                ))
                driver.execute_script("arguments[0].scrollIntoView(true);", textarea)
                time.sleep(0.5)
                textarea.clear()
                text_value = (
                    f"I am requesting a copy of the winning and shortlisted proposals for {event_id}. "
                    f"I am requesting a copy of the winning and shortlisted proposals for the referenced award. "
                    f"The solicitation/contract number is {event_name}."
                )
                textarea.send_keys(text_value)

                # 👉 Wait for manual submission
                print(f"⏳ Please click the Submit button manually for Event: {event_id}")
                wait.until(EC.staleness_of(textarea))  # Wait until page reloads
                print(f"✅ Submission detected for Event: {event_id}")

                results.append({"Event Name": event_id, "Event ID": event_name, "Success Msg": "Submitted"})
                print(f"✅ Submitted for Event Name: {event_id}")

                # Reload page and refill static fields for next submission
                driver.get(URL)
                time.sleep(5)
                fill_static_fields(driver, wait)

            except Exception as e:
                results.append({"Event Name": event_id, "Event ID": event_name, "Success Msg": f"Failed - {e}"})
                print(f"❌ Failed for Event Name: {event_id} - {e}")

    finally:
        driver.quit()

    # Save results
    pd.DataFrame(results).to_excel(EXCEL_OUTPUT, index=False)
    print(f"📂 Output saved to {EXCEL_OUTPUT}")


# ------------------ MAIN ------------------
if __name__ == "__main__":
    process_events()
