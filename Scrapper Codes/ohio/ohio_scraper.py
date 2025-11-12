import os
import time
import zipfile
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd

class OhioBuysScraper:
    def __init__(self, download_path="downloads"):
        self.download_path = os.path.abspath(download_path)
        os.makedirs(self.download_path, exist_ok=True)
        
        chrome_options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": self.download_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver, 20)
        
        self.data = []
        self.current_page = 0
        
    def wait_for_page_load(self, timeout=10):
        """Wait for page to finish loading"""
        time.sleep(2)
        try:
            self.wait.until(
                lambda driver: driver.execute_script("return document.readyState") == "complete"
            )
        except:
            pass
    
    def apply_awarded_filter(self):
        """Apply the 'Awarded: Yes' filter"""
        try:
            awarded_input = self.wait.until(
                EC.element_to_be_clickable((By.ID, "body_x_cbRfpPubAward_search"))
            )
            awarded_input.click()
            time.sleep(1)
            
            yes_option = self.wait.until(
                EC.element_to_be_clickable((By.ID, "body_x__ctl1_True"))
            )
            yes_option.click()
            time.sleep(1)
            
            search_button = self.wait.until(
                EC.element_to_be_clickable((By.ID, "body_x_prxFilterBar_x_cmdSearchBtn"))
            )
            search_button.click()
            self.wait_for_page_load()
            
            print("Filter applied: Awarded = Yes")
            return True
        except Exception as e:
            print(f"Error applying filter: {e}")
            return False
    
    def navigate_to_page(self, page_index):
        """Navigate to a specific page using pagination"""
        try:
            if page_index == 0:
                return True
            
            print(f"Navigating to page {page_index + 1}...")
            
            script = f"__ivCtrl['body_x_grid_grd'].GoToPageOfGrid(0, {page_index});"
            self.driver.execute_script(script)
            self.wait_for_page_load()
            time.sleep(2)
            
            print(f"Now on page {page_index + 1}")
            return True
            
        except Exception as e:
            print(f"✗ Error navigating to page {page_index + 1}: {e}")
            return False
    
    def get_current_page_number(self):
        """Get the current page number from the UI"""
        try:
            active_page = self.driver.find_element(
                By.XPATH,
                "//button[contains(@class, 'active') and contains(@data-page-index, '')]"
            )
            page_index = int(active_page.get_attribute('data-page-index'))
            return page_index
        except:
            return 0
    
    def has_next_page(self):
        """Check if there's a next page available"""
        try:
            next_button = self.driver.find_element(By.ID, "body_x_grid_PagerBtnNextPage")
            
            is_disabled = 'disable' in next_button.get_attribute('class').lower()
            aria_disabled = next_button.get_attribute('aria-disabled') == 'true'
            
            return not (is_disabled or aria_disabled)
        except:
            return False
    
    def go_to_next_page(self):
        """Navigate to the next page"""
        try:
            next_button = self.wait.until(
                EC.element_to_be_clickable((By.ID, "body_x_grid_PagerBtnNextPage"))
            )
            
            next_page_index = int(next_button.get_attribute('data-page-index'))
            
            # Click next page button
            next_button.click()
            self.wait_for_page_load()
            time.sleep(2)
            
            self.current_page = next_page_index
            print(f"Moved to page {self.current_page + 1}")
            return True
            
        except Exception as e:
            print(f"Error going to next page: {e}")
            return False
    
    def get_all_opportunity_links(self):
        """Extract all opportunity links from the current page"""
        try:
            self.wait.until(
                EC.presence_of_element_located((By.ID, "body_x_grid_grd"))
            )
            time.sleep(2)
            
            # Find all edit buttons (opportunity links)
            links = self.driver.find_elements(
                By.XPATH, 
                "//a[contains(@id, '_img___colManagegrid') and contains(@class, 'iv-button')]"
            )
            
            opportunity_urls = []
            for link in links:
                href = link.get_attribute('href')
                if href:
                    opportunity_urls.append(href)
            
            print(f"Found {len(opportunity_urls)} opportunities on this page")
            return opportunity_urls
        except Exception as e:
            print(f"Error getting opportunity links: {e}")
            return []
    
    def scrape_opportunity_details(self):
        """Scrape details from the opportunity page"""
        try:
            self.wait_for_page_load()
            
            solicitation_id = self.wait.until(
                EC.presence_of_element_located(
                    (By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_lblProcessCode")
                )
            ).text
            
            solicitation_name = self.driver.find_element(
                By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_lblLabel"
            ).text
            
            begin_date = self.driver.find_element(
                By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_lblBeginDate"
            ).text
            
            end_date = self.driver.find_element(
                By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_lblEndDate"
            ).text
            
            status = self.driver.find_element(
                By.XPATH, 
                "//div[@data-iv-control='body_x_tabc_rfp_ext_prxrfp_ext_x_selStatusCode']//div[@class='text']"
            ).text
            
            opportunity_data = {
                'Solicitation ID': solicitation_id,
                'Solicitation Name': solicitation_name,
                'Begin Date': begin_date,
                'End Date': end_date,
                'Solicitation Status': status
            }
            
            print(f"Scraped: {solicitation_id} - {solicitation_name}")
            return opportunity_data
        except Exception as e:
            print(f"Error scraping opportunity details: {e}")
            return None
    
    def download_documents(self, solicitation_id):
        """Download all documents from Solicitation Documents section"""
        try:
            # Create a folder for this solicitation's documents
            doc_folder = os.path.join(self.download_path, f"{solicitation_id}_docs")
            os.makedirs(doc_folder, exist_ok=True)
            
            # Find all download links in the Att. column
            download_links = self.driver.find_elements(
                By.XPATH,
                "//a[contains(@class, 'iv-download-file') and contains(@href, '/bare.aspx/en/fil/download_public/')]"
            )
            
            if not download_links:
                print(f"No documents found for {solicitation_id}")
                return None
            
            print(f"Downloading {len(download_links)} document(s)...")
            
            downloaded_files = []
            for i, link in enumerate(download_links, 1):
                try:
                    file_name = link.text.strip()
                    if not file_name:
                        file_name = f"document_{i}.pdf"
                    
                    # Click to download
                    self.driver.execute_script("arguments[0].click();", link)
                    time.sleep(3)
                    
                    # Wait for file to complete downloading
                    self.wait_for_download_complete()
                    
                    # Move the downloaded file to the specific folder
                    latest_file = self.get_latest_file(self.download_path)
                    if latest_file:
                        new_path = os.path.join(doc_folder, file_name)
                        os.rename(latest_file, new_path)
                        downloaded_files.append(new_path)
                        print(f"Downloaded: {file_name}")
                    
                except Exception as e:
                    print(f"Error downloading document {i}: {e}")
                    continue
            
            # Create a zip file
            if downloaded_files:
                zip_filename = os.path.join(self.download_path, f"{solicitation_id}_documents.zip")
                with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for file in downloaded_files:
                        zipf.write(file, os.path.basename(file))
                
                # Remove the temporary folder
                import shutil
                shutil.rmtree(doc_folder)
                
                print(f"Created zip: {zip_filename}")
                return zip_filename
            
            return None
            
        except Exception as e:
            print(f"Error downloading documents: {e}")
            return None
    
    def wait_for_download_complete(self, timeout=60):
        """Wait for download to complete"""
        end_time = time.time() + timeout
        while time.time() < end_time:
            downloading = False
            for filename in os.listdir(self.download_path):
                if filename.endswith('.crdownload') or filename.endswith('.tmp'):
                    downloading = True
                    break
            
            if not downloading:
                time.sleep(1)
                return True
            
            time.sleep(0.5)
        
        return False
    
    def get_latest_file(self, folder):
        """Get the most recently created file in a folder"""
        files = [os.path.join(folder, f) for f in os.listdir(folder)]
        files = [f for f in files if os.path.isfile(f) and not f.endswith('.crdownload')]
        if not files:
            return None
        return max(files, key=os.path.getctime)
    
    def return_to_list_page_with_filters(self):
        """Navigate back to the main list page and restore filters and page position"""
        try:
            print(f"Returning to list page {self.current_page + 1}...")
            
            # Go back to the main page
            self.driver.get("https://ohiobuys.ohio.gov/page.aspx/en/rfp/request_browse_public?historyBack=1")
            self.wait_for_page_load()
            
            # Re-apply the filter
            if not self.apply_awarded_filter():
                return False
            
            # Navigate back to the correct page if not on page 1
            if self.current_page > 0:
                if not self.navigate_to_page(self.current_page):
                    return False
            
            return True
            
        except Exception as e:
            print(f"Error returning to list page: {e}")
            return False
    
    def save_to_excel(self, filename="ohiobuys_awarded_solicitations.xlsx"):
        """Save scraped data to Excel"""
        if not self.data:
            print("No data to save!")
            return
        
        df = pd.DataFrame(self.data)
        excel_path = os.path.join(self.download_path, filename)
        df.to_excel(excel_path, index=False)
        print(f"\nData saved to: {excel_path}")
        print(f"  Total records: {len(self.data)}")
    
    def scrape_all_opportunities(self, max_pages=None):
        """Main scraping function"""
        try:
            print("Opening OhioBuys website...")
            self.driver.get("https://ohiobuys.ohio.gov/page.aspx/en/rfp/request_browse_public?historyBack=1")
            self.wait_for_page_load()
            
            if not self.apply_awarded_filter():
                print("Failed to apply filter. Exiting.")
                return
            
            self.current_page = 0
            total_opportunities_processed = 0
            
            while True:
                if max_pages and self.current_page >= max_pages:
                    print(f"\nReached maximum page limit: {max_pages}")
                    break
                
                print(f"Processing Page {self.current_page + 1}")
                
                # Get all opportunity links on current page
                opportunity_urls = self.get_all_opportunity_links()
                
                if not opportunity_urls:
                    print("No opportunities found on this page.")
                    break
                
                # Process each opportunity on the current page
                for idx, url in enumerate(opportunity_urls, 1):
                    print(f"\n[Page {self.current_page + 1} - {idx}/{len(opportunity_urls)}] Processing opportunity...")
                    
                    # Navigate to opportunity page
                    self.driver.get(url)
                    self.wait_for_page_load()
                    
                    # Scrape details
                    opportunity_data = self.scrape_opportunity_details()
                    if opportunity_data:
                        # Download documents
                        zip_file = self.download_documents(opportunity_data['Solicitation ID'])
                        
                        self.data.append(opportunity_data)
                        total_opportunities_processed += 1
                    
                    # Return to list page with filters and correct page
                    if not self.return_to_list_page_with_filters():
                        print("Failed to return to list page. Stopping.")
                        break
                
                # Check if there's a next page
                if not self.has_next_page():
                    print("\nNo more pages available. Scraping complete!")
                    break
                
                # Move to next page
                if not self.go_to_next_page():
                    print("Failed to navigate to next page. Stopping.")
                    break
            
            print(f"\n{'='*60}")
            print(f"Scraping Summary")
            print(f"{'='*60}")
            print(f"Total pages processed: {self.current_page + 1}")
            print(f"Total opportunities scraped: {total_opportunities_processed}")
            
            # Save all data to Excel
            self.save_to_excel()
            
        except Exception as e:
            print(f"\nCritical error in main scraping function: {e}")
            if self.data:
                self.save_to_excel()
    
    def close(self):
        """Close the browser"""
        if self.driver:
            self.driver.quit()
            print("\nBrowser closed")


def main():
    scraper = OhioBuysScraper(download_path="downloads")
    
    try:
        scraper.scrape_all_opportunities(max_pages=None)
        
    except KeyboardInterrupt:
        print("\n\nScraping interrupted by user")
        scraper.save_to_excel()
    
    except Exception as e:
        print(f"\nUnexpected error: {e}")
    
    finally:
        scraper.close()


if __name__ == "__main__":
    main()