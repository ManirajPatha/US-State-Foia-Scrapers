import time
import os
import zipfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import pandas as pd
from datetime import datetime

class HawaiiProcurementScraper:
    def __init__(self, download_path=None):
        """Initialize the scraper with Chrome WebDriver"""
        self.download_path = download_path or os.path.join(os.getcwd(), "downloads")
        
        # Create downloads directory if it doesn't exist
        if not os.path.exists(self.download_path):
            os.makedirs(self.download_path)
        
        # Setup Chrome options
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
        self.wait = WebDriverWait(self.driver, 15)
        self.scraped_data = []
        
    def close_any_modal(self):
        """Close any open modal dialogs"""
        try:
            # Try to find and close Cancel button
            cancel_btn = self.driver.find_element(By.XPATH, "//button[@class='btn btn-default' and text()='Cancel']")
            if cancel_btn.is_displayed():
                cancel_btn.click()
                time.sleep(1)
        except:
            pass
        
        try:
            # Try to find and close any close button
            close_btns = self.driver.find_elements(By.CSS_SELECTOR, ".modal .close, .modal-header button.close")
            for btn in close_btns:
                if btn.is_displayed():
                    btn.click()
                    time.sleep(1)
                    break
        except:
            pass
    
    def scroll_to_element(self, element):
        """Scroll element into view"""
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.5)
    
    def start_scraping(self):
        """Main method to start the scraping process"""
        try:
            # Open the website
            print("Opening website...")
            self.driver.get("https://hands.ehawaii.gov/hands/opportunities")
            time.sleep(5)
            
            # Close any initial modals
            self.close_any_modal()
            
            # Click on "Show More Search Criteria"
            print("Clicking 'Show More Search Criteria'...")
            self.click_show_more_criteria()
            time.sleep(2)
            
            # Change status filter from Posted/Released to Closed
            print("Changing status filter to 'Closed'...")
            self.change_status_filter()
            time.sleep(5)
            
            # Start scraping pages
            page_num = 1
            while True:
                print(f"\nProcessing Page {page_num}")
                
                # Close any modals before processing
                self.close_any_modal()
                
                if not self.scrape_current_page():
                    break
                
                # Try to click Next button
                if not self.click_next_page():
                    print("No more pages to process.")
                    break
                
                page_num += 1
                time.sleep(3)
            
            # Save data to Excel
            self.save_to_excel()
            
        except Exception as e:
            print(f"Error during scraping: {str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            print("\nClosing browser...")
            time.sleep(2)
            self.driver.quit()
    
    def click_show_more_criteria(self):
        """Click on 'Show More Search Criteria' link"""
        try:
            show_more = self.wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Show More Search Criteria"))
            )
            self.scroll_to_element(show_more)
            show_more.click()
        except TimeoutException:
            print("'Show More Search Criteria' button not found or already expanded")
    
    def change_status_filter(self):
        """Change status filter from Posted/Released to Closed"""
        try:
            # Scroll to the status dropdown
            self.driver.execute_script("window.scrollTo(0, 300);")
            time.sleep(1)
            
            # Step 1: Click on the X icon to remove "Posted/Released"
            print("Removing 'Posted/Released' filter...")
            try:
                # Find the remove button (X icon) for Posted/Released
                remove_btn = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.c-token span.c-remove"))
                )
                self.scroll_to_element(remove_btn)
                time.sleep(0.5)
                
                # Click using JavaScript
                self.driver.execute_script("arguments[0].click();", remove_btn)
                time.sleep(1)
                print("Removed 'Posted/Released'")
            except Exception as e:
                print(f"Could not remove Posted/Released: {str(e)}")
            
            # Step 2: Click on the dropdown to open it
            print("Opening status dropdown...")
            dropdown_btn = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.cuppa-dropdown div.c-btn"))
            )
            self.scroll_to_element(dropdown_btn)
            time.sleep(0.5)
            
            # Click to open dropdown
            self.driver.execute_script("arguments[0].click();", dropdown_btn)
            time.sleep(2)
            
            # Step 3: Select "Closed" from the dropdown
            print("Selecting 'Closed' from dropdown...")
            try:
                # Wait for dropdown to be visible
                dropdown_list = self.wait.until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, "div.dropdown-list"))
                )
                
                # Find the "Closed" checkbox
                closed_checkbox = self.driver.find_element(
                    By.XPATH, 
                    "//label[text()='Closed']/preceding-sibling::input[@type='checkbox']"
                )
                
                # Check if it's not already selected
                if not closed_checkbox.is_selected():
                    # Click the label instead of checkbox (more reliable)
                    closed_label = self.driver.find_element(By.XPATH, "//label[text()='Closed']")
                    self.driver.execute_script("arguments[0].click();", closed_label)
                    time.sleep(1)
                    print("Selected 'Closed'")
                else:
                    print("Closed already selected")
                
            except Exception as e:
                print(f"Error selecting Closed: {str(e)}")
            
            # Step 4: Close the dropdown
            print("Closing dropdown...")
            try:
                # Click outside or on the button again to close
                self.driver.execute_script("arguments[0].click();", dropdown_btn)
                time.sleep(1)
            except:
                # Try clicking outside the dropdown
                body = self.driver.find_element(By.TAG_NAME, "body")
                body.click()
                time.sleep(1)
            
            print("✓ Status filter changed successfully!")
            
        except Exception as e:
            print(f"Error changing status filter: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def scrape_current_page(self):
        """Scrape all opportunities on the current page"""
        try:
            # Wait for table to load
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table.table-opportunity tbody"))
            )
            time.sleep(2)
            
            # Get all opportunity rows (refresh the list each time)
            while True:
                try:
                    # Close any open modals first
                    self.close_any_modal()
                    
                    # Get fresh list of rows
                    rows = self.driver.find_elements(By.CSS_SELECTOR, "table.table-opportunity tbody tr")
                    print(f"Found {len(rows)} opportunities on this page")
                    
                    # Process each row
                    processed_count = 0
                    for idx in range(len(rows)):
                        try:
                            # Re-fetch rows to avoid stale element
                            rows = self.driver.find_elements(By.CSS_SELECTOR, "table.table-opportunity tbody tr")
                            
                            if idx >= len(rows):
                                break
                                
                            row = rows[idx]
                            
                            print(f"\nProcessing opportunity {idx + 1}/{len(rows)}...")
                            
                            # Get solicitation number
                            try:
                                solicitation_num = row.find_element(By.CSS_SELECTOR, "td:nth-child(2) span").text
                                print(f"Solicitation Number: {solicitation_num}")
                            except:
                                solicitation_num = "Unknown"
                                print("Could not get solicitation number")
                            
                            # Scroll to row and click
                            self.scroll_to_element(row)
                            time.sleep(0.5)
                            
                            # Try clicking with JavaScript
                            try:
                                self.driver.execute_script("arguments[0].click();", row)
                            except:
                                row.click()
                            
                            time.sleep(2)
                            
                            # Handle the popup - check for closed notice or OK button
                            try:
                                # First check if the notice is closed
                                try:
                                    closed_notice = WebDriverWait(self.driver, 2).until(
                                        EC.presence_of_element_located((By.XPATH, "//strong[contains(., 'This notice was closed.')]"))
                                    )
                                    if closed_notice:
                                        print("Notice is closed - skipping this opportunity")
                                        # Click on Bidding Opportunities link to go back
                                        try:
                                            back_link = self.driver.find_element(By.XPATH, "//a[contains(@href, 'javascript: void(0);')]//span[contains(text(), 'Bidding Opportunities')]")
                                            back_link.click()
                                            time.sleep(2)
                                            print("Returned to main list")
                                        except:
                                            self.close_any_modal()
                                            time.sleep(1)
                                        continue
                                except TimeoutException:
                                    pass
                                
                                # Try to find and click OK button
                                ok_button = WebDriverWait(self.driver, 5).until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@class='btn btn-primary' and text()='OK']"))
                                )
                                ok_button.click()
                                time.sleep(2)
                                
                                # Switch to the new tab
                                original_window = self.driver.current_window_handle
                                self.driver.switch_to.window(self.driver.window_handles[-1])
                                time.sleep(3)
                                
                                # Scrape opportunity details
                                self.scrape_opportunity_details()
                                
                                # Close the tab and switch back to main window
                                self.driver.close()
                                self.driver.switch_to.window(original_window)
                                time.sleep(1)
                                
                                processed_count += 1
                                
                            except TimeoutException:
                                print("Popup not found - may have opened directly")
                                # Check if new tab was opened
                                if len(self.driver.window_handles) > 1:
                                    original_window = self.driver.current_window_handle
                                    self.driver.switch_to.window(self.driver.window_handles[-1])
                                    time.sleep(3)
                                    
                                    self.scrape_opportunity_details()
                                    
                                    self.driver.close()
                                    self.driver.switch_to.window(original_window)
                                    time.sleep(1)
                                    processed_count += 1
                                else:
                                    # No popup and no new tab - close any modal and skip this opportunity
                                    print("No popup or new tab found - skipping this opportunity")
                                    self.close_any_modal()
                                    time.sleep(1)
                                    continue
                            
                        except Exception as e:
                            print(f"Error processing row {idx + 1}: {str(e)} - Skipping and continuing...")
                            # Make sure we're back on the main window
                            try:
                                if len(self.driver.window_handles) > 1:
                                    self.driver.switch_to.window(self.driver.window_handles[-1])
                                    self.driver.close()
                                self.driver.switch_to.window(self.driver.window_handles[0])
                                self.close_any_modal()
                            except:
                                pass
                            continue
                    
                    print(f"\nProcessed {processed_count} opportunities on this page")
                    break
                    
                except Exception as e:
                    print(f"Error in page processing loop: {str(e)}")
                    break
            
            return True
            
        except Exception as e:
            print(f"Error scraping page: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def download_individual_attachments(self, solicitation_num):
        """Download individual attachments and create a zip file"""
        try:
            print("Looking for individual attachments...")
            
            # Find all attachment links
            attachment_links = self.driver.find_elements(
                By.XPATH, 
                "//a[contains(@href, '/static-resources/') or contains(@href, '.pdf') or contains(@href, '.doc')]"
            )
            
            if not attachment_links:
                print("No individual attachments found")
                return False
            
            print(f"Found {len(attachment_links)} individual attachment(s)")
            
            # Create a folder for this solicitation's attachments
            sol_folder = os.path.join(self.download_path, solicitation_num.replace('/', '_'))
            if not os.path.exists(sol_folder):
                os.makedirs(sol_folder)
            
            downloaded_files = []
            
            # Download each attachment
            for i, link in enumerate(attachment_links):
                try:
                    # Get the filename from the link text or href
                    filename = link.text.strip() if link.text.strip() else f"attachment_{i+1}"
                    if not filename.endswith(('.pdf', '.doc', '.docx', '.xls', '.xlsx')):
                        href = link.get_attribute('href')
                        if href:
                            filename = href.split('/')[-1]
                    
                    print(f"  Downloading: {filename}")
                    
                    # Scroll to element and click
                    self.scroll_to_element(link)
                    time.sleep(0.5)
                    
                    # Click to download
                    try:
                        self.driver.execute_script("arguments[0].click();", link)
                    except:
                        link.click()
                    
                    time.sleep(2)
                    downloaded_files.append(filename)
                    
                except Exception as e:
                    print(f"  Error downloading attachment {i+1}: {str(e)}")
                    continue
            
            # If files were downloaded, create a zip file
            if downloaded_files:
                print(f"Creating zip file for {solicitation_num}...")
                zip_filename = os.path.join(self.download_path, f"{solicitation_num.replace('/', '_')}_attachments.zip")
                
                # Wait a bit for downloads to complete
                time.sleep(3)
                
                try:
                    with zipfile.ZipFile(zip_filename, 'w') as zipf:
                        # Add all downloaded files to zip
                        for filename in downloaded_files:
                            # Look for the file in download directory
                            file_path = os.path.join(self.download_path, filename)
                            if os.path.exists(file_path):
                                zipf.write(file_path, filename)
                                print(f"  Added {filename} to zip")
                    
                    print(f"✓ Created zip file: {zip_filename}")
                    return True
                    
                except Exception as e:
                    print(f"Error creating zip file: {str(e)}")
                    return False
            
            return len(downloaded_files) > 0
            
        except Exception as e:
            print(f"Error downloading individual attachments: {str(e)}")
            return False
    
    def scrape_opportunity_details(self):
        """Scrape details from the opportunity detail page"""
        try:
            # Wait for page to load
            time.sleep(3)
            
            # Extract all details
            data = {}
            
            # Solicitation Number
            try:
                sol_num_elem = self.driver.find_element(By.XPATH, "//dt[text()='Solicitation Number']/following-sibling::dd[1]")
                data['Solicitation Number'] = sol_num_elem.text.split('version:')[0].strip()
            except NoSuchElementException:
                data['Solicitation Number'] = 'Unknown'
                print("Warning: Solicitation Number not found")
            
            # Status - CRITICAL: Check if status is "Awarded"
            try:
                status_elem = self.driver.find_element(By.XPATH, "//dt[text()='Status']/following-sibling::dd[1]")
                data['Status'] = status_elem.text.strip()
                
                # Only continue if status is "Awarded"
                if data['Status'] != 'Awarded':
                    print(f"Status is '{data['Status']}', not 'Awarded'. Skipping this opportunity.")
                    return
                    
            except NoSuchElementException:
                data['Status'] = 'Unknown'
                print("Warning: Status not found. Skipping this opportunity.")
                return
            
            print(f"✓ Status is 'Awarded'. Continuing with scraping...")
            
            # Department
            try:
                dept_elem = self.driver.find_element(By.XPATH, "//dt[contains(text(),'Department')]/following-sibling::dd[1]")
                data['Department'] = dept_elem.text.strip()
            except NoSuchElementException:
                data['Department'] = ''
            
            # Islands
            try:
                islands_elem = self.driver.find_element(By.XPATH, "//dt[contains(text(),'Islands')]/following-sibling::dd[1]")
                data['Islands'] = islands_elem.text.strip()
            except NoSuchElementException:
                data['Islands'] = ''
            
            # Category
            try:
                category_elem = self.driver.find_element(By.XPATH, "//dt[text()='Category']/following-sibling::dd[1]")
                data['Category'] = category_elem.text.strip()
            except NoSuchElementException:
                data['Category'] = ''
            
            # Release Date
            try:
                release_elem = self.driver.find_element(By.XPATH, "//dt[text()='Release Date ']/following-sibling::dd[1]")
                data['Release Date'] = release_elem.text.strip()
            except NoSuchElementException:
                data['Release Date'] = ''
            
            # Amendment Date & Time
            try:
                amend_elem = self.driver.find_element(By.XPATH, "//dt[contains(text(),'Amendment Date')]/following-sibling::dd[1]")
                data['Amendment Date & Time'] = amend_elem.text.strip()
            except NoSuchElementException:
                data['Amendment Date & Time'] = ''
            
            # Offer Due Date & Time
            try:
                due_elem = self.driver.find_element(By.XPATH, "//dt[contains(text(),'Offer Due Date')]/following-sibling::dd[1]")
                data['Offer Due Date & Time'] = due_elem.text.strip()
            except NoSuchElementException:
                data['Offer Due Date & Time'] = ''
            
            # Download attachments
            download_success = False
            
            # Try to download all attachments first
            try:
                download_all_btn = self.driver.find_element(By.XPATH, "//a[contains(@href,'downloadAllAttachments')]")
                print("Found 'Download All' button - downloading all attachments...")
                self.scroll_to_element(download_all_btn)
                download_all_btn.click()
                time.sleep(3)
                download_success = True
                print("✓ Downloaded all attachments")
            except NoSuchElementException:
                print("'Download All' button not found - trying individual attachments...")
                # Try downloading individual attachments
                download_success = self.download_individual_attachments(data['Solicitation Number'])
            except Exception as e:
                print(f"Error with 'Download All': {str(e)} - trying individual attachments...")
                # Try downloading individual attachments as fallback
                try:
                    download_success = self.download_individual_attachments(data['Solicitation Number'])
                except Exception as e2:
                    print(f"Error downloading individual attachments: {str(e2)}")
            
            if not download_success:
                print("No attachments downloaded - continuing anyway...")
            
            # Append to scraped data
            self.scraped_data.append(data)
            print(f"✓ Successfully scraped data for {data['Solicitation Number']}")
            
        except Exception as e:
            print(f"Error scraping opportunity details: {str(e)} - Skipping and continuing...")
            import traceback
            traceback.print_exc()
    
    def click_next_page(self):
        """Click the Next button to go to the next page"""
        try:
            # Close any open modals first
            self.close_any_modal()
            time.sleep(1)
            
            # Find all dataNav buttons
            nav_buttons = self.driver.find_elements(By.CSS_SELECTOR, "button.dataNav")
            
            next_button = None
            for btn in nav_buttons:
                if "Next" in btn.text or ">" in btn.text:
                    next_button = btn
                    break
            
            if not next_button:
                print("Next button not found")
                return False
            
            # Check if button is disabled
            if next_button.get_attribute('disabled') is not None:
                print("Next button is disabled - no more pages")
                return False
            
            # Scroll to button
            self.scroll_to_element(next_button)
            time.sleep(1)
            
            print("Clicking Next button...")
            
            # Try clicking with JavaScript
            try:
                self.driver.execute_script("arguments[0].click();", next_button)
            except:
                next_button.click()
            
            time.sleep(3)
            print("Successfully navigated to next page")
            return True
            
        except NoSuchElementException:
            print("Next button not found")
            return False
        except Exception as e:
            print(f"Error clicking next page: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def save_to_excel(self):
        """Save scraped data to Excel file"""
        if not self.scraped_data:
            print("\nNo data to save.")
            return
        
        try:
            # Create DataFrame
            df = pd.DataFrame(self.scraped_data)
            
            # Define the column order we want in the Excel
            columns_to_keep = [
                'Solicitation Number',
                'Status',
                'Department',
                'Islands',
                'Category',
                'Release Date',
                'Amendment Date & Time',
                'Offer Due Date & Time'
            ]
            
            # Only keep columns that exist in the dataframe
            final_columns = [col for col in columns_to_keep if col in df.columns]
            df = df[final_columns]
            
            # Create filename with timestamp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"hawaii_opportunities_{timestamp}.xlsx"
            filepath = os.path.join(self.download_path, filename)
            
            # Save to Excel
            df.to_excel(filepath, index=False, engine='openpyxl')
            print(f"\n✓ Data saved to: {filepath}")
            print(f"✓ Total opportunities scraped: {len(self.scraped_data)}")
            
        except Exception as e:
            print(f"Error saving to Excel: {str(e)}")
            import traceback
            traceback.print_exc()


def main():
    """Main function to run the scraper"""
    print("=== Hawaii Procurement Scraper ===\n")
    
    download_path = os.path.join(os.getcwd(), "downloads")
    
    scraper = HawaiiProcurementScraper(download_path=download_path)
    scraper.start_scraping()
    
    print("\n=== Scraping Complete ===")


if __name__ == "__main__":
    main()