import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

input_file = '/home/developer/Desktop/Scraper/rhode_island_awarded_bids.xlsx'


NOTICE_ID_COL = 'notice_id'  
TITLE_COL = 'title'          


try:
    df = pd.read_excel(input_file)
except FileNotFoundError:
    print(f"Error: The file {input_file} was not found.")
    exit(1)

if NOTICE_ID_COL not in df.columns or TITLE_COL not in df.columns:
    print(f"Error: Required columns '{NOTICE_ID_COL}' and/or '{TITLE_COL}' not found in the Excel file.")
    print("Available columns:", df.columns.tolist())
    exit(1)

df = df.head(20)
print(f"Processing first 20 rows out of {len(df)} total rows.")

results = []

chrome_options = Options()


chromedriver_path = '/usr/bin/chromedriver'

if not os.path.exists(chromedriver_path):
    print(f"Error: ChromeDriver not found at {chromedriver_path}. Please install it.")
    exit(1)

service = Service(chromedriver_path)
try:
    driver = webdriver.Chrome(service=service, options=chrome_options)
except Exception as e:
    print(f"Error initializing WebDriver: {str(e)}")
    exit(1)

for index, row in df.iterrows():
    notice_id = row[NOTICE_ID_COL]
    title = row[TITLE_COL]
    
    try:
        driver.get('https://ri.ng.mil/FOIA-Request/')
    except Exception as e:
        results.append({
            'notice id': notice_id,
            'title': title,
            'Success Msg': f"Failed - Could not load website: {str(e)}"
        })
        continue
    
    time.sleep(3)
    
    try:
        dropdown = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'dnn_ctr8348_View_ddlContactFormRecipient'))
        )
        dropdown.click()
        option = driver.find_element(By.CSS_SELECTOR, 'option[value="114"]')
        option.click()
        
        name_field = driver.find_element(By.ID, 'dnn_ctr8348_View_defName')
        name_field.clear()
        name_field.send_keys('Raaj Thipparthy')
        
        email_field = driver.find_element(By.ID, 'dnn_ctr8348_View_defEmail')
        email_field.clear()
        email_field.send_keys('raajnrao@gmail.com')
        
        subject_field = driver.find_element(By.ID, 'dnn_ctr8348_View_defSubject')
        subject_field.clear()
        subject_field.send_keys('Requesting for records')
        
        message_text = (
            f"I am requesting a copy of the winning and shortlisted proposals for {notice_id}. "
            f"I am requesting a copy of the winning and shortlisted proposals for the referenced award. "
            f"The solicitation/contract number is {title}."
        )
        message_field = driver.find_element(By.ID, 'dnn_ctr8348_View_defMessage')
        message_field.clear()
        message_field.send_keys(message_text)
        
        checkbox = driver.find_element(By.ID, 'dnn_ctr8348_View_cbContactFormContactMe')
        if not checkbox.is_selected():
            checkbox.click()
        
        print(f"Please solve the reCAPTCHA for row {index + 1} (notice id: {notice_id}) and press Enter when done.")
        input()
        
        submit_btn = driver.find_element(By.ID, 'btnContactFormSubmit8348')
        submit_btn.click()
        
        time.sleep(5)
        
        if "sent" in driver.page_source.lower() or "thank you" in driver.page_source.lower():
            success_msg = "Submitted"
        else:
            success_msg = "Failed - Check page for errors"
    
    except Exception as e:
        success_msg = f"Failed - Error: {str(e)}"
    
    results.append({
        'notice id': notice_id,
        'title': title,
        'Success Msg': success_msg
    })
    
    time.sleep(2)

try:
    driver.quit()
except Exception as e:
    print(f"Error closing WebDriver: {str(e)}")

output_df = pd.DataFrame(results)
try:
    output_df.to_excel('output.xlsx', index=False)
    print("Automation complete. Results saved to 'output.xlsx'.")
except Exception as e:
    print(f"Error saving output file: {str(e)}")