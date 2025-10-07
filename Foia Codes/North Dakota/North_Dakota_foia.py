import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import os

def setup_driver():
    """Setup Chrome driver with options"""
    chrome_options = Options()
    
    # Uncomment the line below to run in headless mode
    # chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument('--start-maximized')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    
    # Try to find Chrome/Chromium binary
    possible_chrome_paths = [
        '/usr/bin/google-chrome',
        '/usr/bin/google-chrome-stable',
        '/usr/bin/chromium',
        '/usr/bin/chromium-browser',
        '/snap/bin/chromium'
    ]
    
    chrome_binary = None
    for path in possible_chrome_paths:
        if os.path.exists(path):
            chrome_binary = path
            break
    
    if chrome_binary:
        chrome_options.binary_location = chrome_binary
        print(f"Using Chrome/Chromium from: {chrome_binary}")
    
    try:
        # Use webdriver-manager to automatically handle driver setup
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        print(f"Error setting up Chrome driver: {e}")
        print("\nPlease install Chrome or Chromium:")
        print("  sudo apt update")
        print("  sudo apt install chromium-browser")
        print("  OR")
        print("  sudo apt install google-chrome-stable")
        raise

def fill_form(driver, notice_id, title, first_name="Raaj", last_name="Thipparthy", email="raajnrao@gmail.com"):
    """Fill and submit the public records request form with proper success/error detection"""
    try:
        url = "https://www.bismarcknd.gov/FormCenter/Administration-2/Request-for-Public-Records-246"
        driver.get(url)
        wait = WebDriverWait(driver, 10)

        # Fill First Name
        first_name_field = wait.until(EC.presence_of_element_located((By.ID, "e_1")))
        first_name_field.clear()
        first_name_field.send_keys(first_name)
        time.sleep(1)

        # Fill Last Name
        last_name_field = driver.find_element(By.ID, "e_2")
        last_name_field.clear()
        last_name_field.send_keys(last_name)
        time.sleep(1)

        # Fill Email
        email_field = driver.find_element(By.ID, "e_9")
        email_field.clear()
        email_field.send_keys(email)
        time.sleep(1)

        # Fill Records
        records_text = f"I am requesting a copy of the winning and shortlisted proposals for {notice_id}. The solicitation/contract number is {title}."
        records_field = driver.find_element(By.ID, "e_11")
        records_field.clear()
        records_field.send_keys(records_text)
        time.sleep(1)

        # Select 'Other' Department
        other_checkbox = driver.find_element(By.ID, "e_14_7")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", other_checkbox)
        time.sleep(1)
        if not other_checkbox.is_selected():
            try:
                other_checkbox.click()
            except:
                driver.execute_script("arguments[0].click();", other_checkbox)
        time.sleep(1)

        # Uncheck 'Receive an email copy'
        email_copy_checkbox = driver.find_element(By.ID, "wantCopy")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", email_copy_checkbox)
        time.sleep(0.3)
        if email_copy_checkbox.is_selected():
            try:
                email_copy_checkbox.click()
            except:
                driver.execute_script("arguments[0].click();", email_copy_checkbox)
        time.sleep(1)

        time.sleep(10)

        # Click Submit
        submit_button = driver.find_element(By.ID, "btnFormSubmit")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", submit_button)
        time.sleep(1)
        try:
            submit_button.click()
        except:
            driver.execute_script("arguments[0].click();", submit_button)

        # Wait for response
        time.sleep(3)

        # Check for errors
        error_elements = driver.find_elements(By.CSS_SELECTOR, ".form-error, .error, .alert-danger")
        if error_elements:
            errors = [e.text for e in error_elements if e.text.strip() != ""]
            return "Form Rejected: " + "; ".join(errors)

        # Check for success messages
        success_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Thank you') or contains(text(), 'successfully')]")
        if success_elements:
            return "Form Submitted Successfully"
        else:
            return "Submission Status Unknown - Please Verify"

    except Exception as e:
        return f"Error: {str(e)}"


def process_excel_file(input_file, output_file):
    """Process Excel file and submit forms (all rows)"""
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    if 'Notice ID' not in df.columns or 'Title' not in df.columns:
        print("Error: Excel file must contain 'Notice ID' and 'Title' columns")
        return
    
    if 'Success Msg' not in df.columns:
        df['Success Msg'] = ''
    
    driver = setup_driver()
    
    try:
        for index, row in df.iterrows():
            notice_id = row['Notice ID']
            title = row['Title']
            
            print(f"\nProcessing row {index + 1}/{len(df)}")
            print(f"Notice ID: {notice_id}")
            print(f"Title: {title}")
            
            success_msg = fill_form(driver, notice_id, title)
            
            df.at[index, 'Success Msg'] = success_msg
            print(f"Result: {success_msg}")
            
            time.sleep(2)

            if (index + 1) % 15 == 0:
                print("⏸ Taking a 30-second break...")
                time.sleep(30)

        df.to_excel(output_file, index=False)
        print(f"\n✅ Process completed! Results saved to: {output_file}")
        
    except Exception as e:
        print(f"Error during processing: {e}")
        df.to_excel(output_file, index=False)
        print(f"Partial results saved to: {output_file}")
        
    finally:
        driver.quit()

if __name__ == "__main__":
    INPUT_FILE = r"C:\Users\ADMIN\Desktop\Scraper\north_dakota_closed_rfps.xlsx"
    OUTPUT_FILE = r"C:\Users\ADMIN\Desktop\Scraper\north_dakota_foia_results.xlsx"

    print("Starting Bismarck Public Records Form Automation")
    print("=" * 50)
    
    process_excel_file(INPUT_FILE, OUTPUT_FILE)
    
    print("\n" + "=" * 50)
    print("Automation completed!")
