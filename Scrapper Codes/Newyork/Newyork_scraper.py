from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
from datetime import datetime

# Set up Selenium WebDriver (assuming ChromeDriver is installed and in PATH)
service = Service()  # You may need to specify the path to chromedriver.exe
options = webdriver.ChromeOptions()
options.add_argument('--headless')  # Run in headless mode (optional)
driver = webdriver.Chrome(service=service, options=options)

# Navigate to the URL
url = "https://a856-cityrecord.nyc.gov/Search/Advanced"
driver.get(url)

# Wait for the page to load
time.sleep(2)  # Basic wait; consider using WebDriverWait for better practice

# Select "Procurement" from section dropdown (value=6)
section_select = Select(driver.find_element(By.ID, "ddlSection"))
section_select.select_by_value("6")

# Select "Award" from notice type dropdown (value=2)
notice_type_select = Select(driver.find_element(By.ID, "ddlNoticeType"))
notice_type_select.select_by_value("2")

# Click the submit button
submit_button = driver.find_element(By.ID, "AdvancedSubmitButton")
submit_button.click()

# Wait for results to load
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "notice-container")))

# Prepare list to store data
data = []

# Current month and year (based on October 2025)
current_month = 10  # October
current_year = 2025

# Function to parse date from string like "10/2/2025"
def parse_date(date_str):
    try:
        return datetime.strptime(date_str, "%m/%d/%Y")
    except ValueError:
        return None

# Loop through pages
page = 1
while True:
    print(f"Scraping page {page}...")
    
    # Find all notice containers
    notices = driver.find_elements(By.CLASS_NAME, "notice-container")
    
    for notice in notices:
        # Get title and link
        title_elem = notice.find_element(By.TAG_NAME, "a")
        title = title_elem.text.strip()
        link = title_elem.get_attribute("href")
        
        # Get agency
        agency = notice.find_element(By.TAG_NAME, "strong").text.strip()
        
        # Get date (from small text with fa-calendar)
        small_texts = notice.find_elements(By.CSS_SELECTOR, "small")
        date_str = ""
        for small in small_texts:
            if "fa-calendar" in small.get_attribute("innerHTML"):
                date_str = small.text.strip().split()[-1]  # e.g., "10/2/2025"
                break
        
        # Parse date and check if in current month
        notice_date = parse_date(date_str)
        if notice_date and (notice_date.month != current_month or notice_date.year != current_year):
            # Stop if date is not in current month
            print(f"Stopping: Found date {date_str} outside current month.")
            driver.quit()
            # Save data to Excel
            df = pd.DataFrame(data)
            df.to_excel("ny_city_record_awards_october_2025.xlsx", index=False)
            print("Data saved to ny_city_record_awards_october_2025.xlsx")
            exit()  # Or break/return depending on context
        
        # Get description (may be empty)
        description = notice.find_element(By.CLASS_NAME, "short-description").text.strip()
        
        # Append to data
        data.append({
            "Title": title,
            "Link": link,
            "Agency": agency,
            "Date": date_str,
            "Description": description
        })
    
    # Check for next button
    try:
        next_button = driver.find_element(By.CLASS_NAME, "next")
        if "disabled" in next_button.get_attribute("class"):
            break
        next_button.click()
        # Wait for next page to load
        WebDriverWait(driver, 10).until(EC.staleness_of(notices[0]))  # Wait until current notices are stale
        page += 1
    except:
        break

# Close driver
driver.quit()

# Save data to Excel if not already saved
if data:
    df = pd.DataFrame(data)
    df.to_excel("ny_city_record_awards_october_2025.xlsx", index=False)
    print("Data saved to ny_city_record_awards_october_2025.xlsx")
else:
    print("No data found for the current month.")
