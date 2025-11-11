from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import pandas as pd
from datetime import datetime, timedelta


def setup_driver():
    """Initialize and return Chrome driver"""
    service = Service()
    options = webdriver.ChromeOptions()
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(service=service, options=options)
    return driver


def get_date_range():
    """Calculate start date (2 years back) and end date (yesterday)"""
    end_date = datetime.now() - timedelta(days=1)
    start_date = end_date - timedelta(days=730)
    
    return start_date.strftime("%m/%d/%Y"), end_date.strftime("%m/%d/%Y")

def scrape_opportunity_details(driver, opportunity_url):
    """Scrape detailed information from individual opportunity page"""
    try:
        driver.get(opportunity_url)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "form-body"))
        )
        
        details = {}
        
        form_groups = driver.find_elements(By.CLASS_NAME, "form-md-line-input")
        
        for form_group in form_groups:
            try:
                label_elem = form_group.find_element(By.TAG_NAME, "label")
                value_elem = form_group.find_element(By.CLASS_NAME, "form-control-static")
                
                label = label_elem.text.strip()
                value = value_elem.text.strip()
                
                if label and value:
                    details[label] = value
            except NoSuchElementException:
                continue
        
        return details
    
    except Exception as e:
        print(f"Error scraping opportunity details: {e}")
        return {}


def scrape_page_opportunities(driver, data):
    """Scrape all opportunities from current page"""
    try:
        # Wait for opportunities to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "notice-container"))
        )
        
        notices = driver.find_elements(By.CLASS_NAME, "notice-container")
        opportunity_links = []
        
        for notice in notices:
            try:
                title_elem = notice.find_element(By.TAG_NAME, "a")
                title = title_elem.text.strip()
                link = title_elem.get_attribute("href")
                
                agency_elem = notice.find_element(By.TAG_NAME, "strong")
                agency = agency_elem.text.strip()
                
                small_texts = notice.find_elements(By.CSS_SELECTOR, "small")
                date_str = ""
                for small in small_texts:
                    if "fa-calendar" in small.get_attribute("innerHTML"):
                        date_str = small.text.strip().split()[-1]
                        break
                
                description_elem = notice.find_element(By.CLASS_NAME, "short-description")
                description = description_elem.text.strip()
                
                opportunity_links.append({
                    "Title": title,
                    "Link": link,
                    "Agency": agency,
                    "Date": date_str,
                    "Description": description
                })
            except NoSuchElementException:
                continue
        
        # Visit each opportunity and scrape details
        for i, opp in enumerate(opportunity_links):
            print(f"  Scraping opportunity {i+1}/{len(opportunity_links)}: {opp['Title'][:50]}...")
            
            details = scrape_opportunity_details(driver, opp['Link'])
            
            complete_data = {**opp, **details}
            data.append(complete_data)
            
            driver.back()
            
            # Wait for page to reload
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "notice-container"))
            )
            time.sleep(1)
        
        return len(opportunity_links)
    
    except Exception as e:
        print(f"Error scraping page opportunities: {e}")
        return 0


def main():
    print("Starting NYC City Record scraper...")
    
    driver = setup_driver()
    
    try:
        url = "https://a856-cityrecord.nyc.gov/Search/Advanced"
        driver.get(url)
        time.sleep(2)
        
        # Select Section Procurement
        print("Selecting Section: Procurement...")
        section_select = Select(driver.find_element(By.ID, "ddlSection"))
        section_select.select_by_value("6")
        time.sleep(1)
        
        # Select Notice Type Award
        print("Selecting Notice Type: Award...")
        notice_type_select = Select(driver.find_element(By.ID, "ddlNoticeType"))
        notice_type_select.select_by_value("2")
        time.sleep(1)
        
        # Set Date Range
        start_date, end_date = get_date_range()
        print(f"Setting date range: {start_date} to {end_date}")
        
        start_date_input = driver.find_element(By.ID, "txtStartDate")
        start_date_input.clear()
        start_date_input.send_keys(start_date)
        time.sleep(1)
        
        end_date_input = driver.find_element(By.ID, "txtEndDate")
        end_date_input.clear()
        end_date_input.send_keys(end_date)
        time.sleep(1)
        
        # Click Submit button
        print("Submitting search...")
        submit_button = driver.find_element(By.ID, "AdvancedSubmitButton")
        submit_button.click()
        
        # Wait for results to load
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CLASS_NAME, "notice-container"))
        )
        time.sleep(2)
        
        data = []
        page = 1
        
        # Scrape all pages
        while True:
            print(f"\nScraping page {page}...")
            
            # Scrape opportunities from current page
            opportunities_found = scrape_page_opportunities(driver, data)
            print(f"  Found and scraped {opportunities_found} opportunities on page {page}")
            
            try:
                next_button = driver.find_element(By.CSS_SELECTOR, "a.page-link.next")
                
                # Check if next button is disabled
                parent_li = next_button.find_element(By.XPATH, "..")
                if "disabled" in parent_li.get_attribute("class"):
                    print("Reached last page.")
                    break
                
                print(f"Moving to page {page + 1}...")
                next_button.click()
                
                # Wait for new page to load
                time.sleep(2)
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "notice-container"))
                )
                
                page += 1
                
            except NoSuchElementException:
                print("No more pages found.")
                break
            except TimeoutException:
                print("Timeout waiting for next page to load.")
                break
        
        # Save data to Excel and JSON
        if data:
            df = pd.DataFrame(data)
            filename = f"ny_city_record_awards_{start_date.replace('/', '_')}_to_{end_date.replace('/', '_')}.xlsx"
            json_filename = f"ny_city_record_awards_{start_date.replace('/', '_')}_to_{end_date.replace('/', '_')}.json"
            df.to_excel(filename, index=False)
            df.to_json(json_filename, orient="records", indent=2)
            print(f"\nSuccessfully scraped {len(data)} opportunities")
            print(f"Data saved to {filename}")
            print(f"Data also saved to {json_filename}")
        else:
            print("\nNo data found.")
    
    except Exception as e:
        print(f"Error during scraping: {e}")
    
    finally:
        driver.quit()
        print("\nScraping completed.")


if __name__ == "__main__":
    main()