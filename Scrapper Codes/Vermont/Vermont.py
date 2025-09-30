import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.options import Options

URL = "https://www.vermontbusinessregistry.com/BidSearch.aspx?type=10"
OUTPUT_XLSX = "vermont_awards_data.xlsx"

def setup_driver():
    """Setup Chrome driver with options"""
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(options=options)
    return driver

def parse_award_entry(table):
    """Parse a single award entry from table element"""
    try:
        date_element = table.find_element(By.ID, "lblRowBreakTitle")
        award_date = date_element.text.strip()
    except NoSuchElementException:
        award_date = ""
    
    try:
        title_elements = table.find_elements(By.CSS_SELECTOR, "td.copyReg")
        project_title = title_elements[0].text.strip() if title_elements else ""
    except:
        project_title = ""
    
    try:
        org_element = table.find_element(By.ID, "lblOrganization")
        organization = org_element.text.strip()
    except NoSuchElementException:
        organization = ""
    
    try:
        close_date_element = table.find_element(By.ID, "lblCloseDate")
        close_date = close_date_element.text.strip()
    except NoSuchElementException:
        close_date = ""
    
    return {
        "Award Date": award_date,
        "Project Title": project_title,
        "Organization": organization,
        "Close Date": close_date
    }

def scrape_current_page(driver):
    """Scrape all award entries from the current page"""
    records = []
    
    wait = WebDriverWait(driver, 10)
    
    try:
        tables = driver.find_elements(By.XPATH, "//table[@width='540']")
        
        for table in tables:
            try:
                if table.find_elements(By.ID, "lblRowBreakTitle"):
                    record = parse_award_entry(table)
                    if record["Project Title"]:
                        records.append(record)
            except Exception as e:
                print(f"Error parsing table: {e}")
                continue
                
    except Exception as e:
        print(f"Error finding tables: {e}")
    
    return records

def get_next_page_link(driver, current_page):
    """Find and return the next page link"""
    try:
        next_page_num = current_page + 1
        next_link = driver.find_element(
            By.XPATH, 
            f"//a[contains(@href, \"javascript:__doPostBack('gvResults','Page${next_page_num}')\")]"
        )
        return next_link
    except NoSuchElementException:
        return None

def scrape_vermont_awards():
    """Main scraping function"""
    driver = setup_driver()
    wait = WebDriverWait(driver, 20)
    all_records = []
    current_page = 1
    
    try:
        print(f"Loading initial page: {URL}")
        driver.get(URL)
        
        time.sleep(3)
        
        while True:
            print(f"Scraping page {current_page}...")
        
            page_records = scrape_current_page(driver)
            all_records.extend(page_records)
            
            print(f"Found {len(page_records)} records on page {current_page}")
            
            next_link = get_next_page_link(driver, current_page)
            
            if next_link:
                print(f"Navigating to page {current_page + 1}...")
                driver.execute_script("arguments[0].click();", next_link)
                
                time.sleep(3)
                
                try:
                    wait.until(EC.presence_of_element_located((By.XPATH, "//table[@width='540']")))
                except TimeoutException:
                    print("Timeout waiting for page to load")
                    break
                
                current_page += 1
            else:
                print("No more pages found. Scraping complete.")
                break
        
        if all_records:
            df = pd.DataFrame(all_records)
            df = df.drop_duplicates()
            df.to_excel(OUTPUT_XLSX, index=False)
            print(f"Saved {len(df)} records to {OUTPUT_XLSX}")
        else:
            print("No records found!")
            
    except Exception as e:
        print(f"An error occurred: {e}")
    
    finally:
        driver.quit()

if __name__ == "__main__":
    scrape_vermont_awards()