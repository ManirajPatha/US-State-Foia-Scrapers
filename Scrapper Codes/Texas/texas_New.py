import os
import time
import zipfile
import pandas as pd
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from multiprocessing import Process, Queue, Manager
from queue import Empty
import traceback


def wait_for_download(download_dir, timeout=60):
    """Wait for download to complete by checking for .crdownload files"""
    seconds = 0
    while seconds < timeout:
        time.sleep(1)
        downloading = False
        for file in os.listdir(download_dir):
            if file.endswith('.crdownload') or file.endswith('.tmp'):
                downloading = True
                break
        if not downloading:
            time.sleep(2)
            return True
        seconds += 1
    return False


def get_latest_file(download_dir, exclude_files=None):
    """Get the most recently downloaded file"""
    if exclude_files is None:
        exclude_files = set()
    
    files = [os.path.join(download_dir, f) for f in os.listdir(download_dir) 
             if os.path.isfile(os.path.join(download_dir, f)) 
             and not f.endswith('.crdownload') 
             and not f.endswith('.tmp')
             and os.path.join(download_dir, f) not in exclude_files]
    
    if not files:
        return None
    return max(files, key=os.path.getctime)


def create_zip(files, zip_path):
    """Create a zip file from a list of files"""
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file in files:
            if os.path.exists(file):
                zipf.write(file, os.path.basename(file))


def sanitize_filename(filename):
    """Remove invalid characters from filename"""
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename[:200]


def extract_text_safely(element, selector, attr=None):
    """Safely extract text from an element"""
    try:
        found = element.find_element(By.CSS_SELECTOR, selector)
        if attr:
            return found.get_attribute(attr)
        return found.text.strip()
    except:
        return ""


def extract_field_value(row, field_label):
    """Extract field value after the label from paragraph text"""
    try:
        # Find all p tags in the row
        paragraphs = row.find_elements(By.CSS_SELECTOR, "p")
        for p in paragraphs:
            text = p.text.strip()
            if field_label in text:
                # Split by the label and get the part after it
                value = text.split(field_label, 1)[1].strip()
                return value
        return ""
    except:
        return ""


def extract_contact_info(driver):
    """Extract contact information from the opportunity page"""
    contact_info = {
        'contact_name': '',
        'contact_number': '',
        'contact_email': '',
        'response_due_date': ''
    }
    
    try:
        # Look for contact information cells
        cells = driver.find_elements(By.CSS_SELECTOR, "div.esbd-result-cell")
        
        for cell in cells:
            try:
                cell_text = cell.text.strip()
                
                if "Contact Name:" in cell_text:
                    p_elem = cell.find_element(By.CSS_SELECTOR, "p")
                    contact_info['contact_name'] = p_elem.text.strip()
                
                elif "Contact Number:" in cell_text:
                    p_elem = cell.find_element(By.CSS_SELECTOR, "p")
                    contact_info['contact_number'] = p_elem.text.strip()
                
                elif "Contact Email:" in cell_text:
                    p_elem = cell.find_element(By.CSS_SELECTOR, "p")
                    contact_info['contact_email'] = p_elem.text.strip()
                
                elif "Response Due Date:" in cell_text:
                    p_elem = cell.find_element(By.CSS_SELECTOR, "p")
                    contact_info['response_due_date'] = p_elem.text.strip()
                    
            except:
                continue
                
    except Exception as e:
        print(f"Error extracting contact info: {str(e)[:50]}")
    
    return contact_info


def extract_awards_info(driver):
    """Extract awards information from the opportunity page"""
    awards_data = []
    
    try:
        # Look for award rows - make sure we're getting the data rows, not the header
        award_rows = driver.find_elements(By.CSS_SELECTOR, "div.esbd-awards-row")
        
        print(f"Found {len(award_rows)} award row elements")
        
        for idx, row in enumerate(award_rows):
            try:
                award = {
                    'contractor': '',
                    'mailing_address': '',
                    'value_per_contractor': '',
                    'hub_status': '',
                    'award_date': '',
                    'award_status': ''
                }
                
                # Extract each column from the award row
                columns = row.find_elements(By.CSS_SELECTOR, "div.esbd-award-result-column")
                
                print(f"Award row {idx}: found {len(columns)} columns")
                
                if len(columns) >= 6:
                    try:
                        award['contractor'] = columns[0].find_element(By.CSS_SELECTOR, "p").text.strip()
                    except:
                        award['contractor'] = ''
                    
                    try:
                        award['mailing_address'] = columns[1].find_element(By.CSS_SELECTOR, "p").text.strip()
                    except:
                        award['mailing_address'] = ''
                    
                    try:
                        award['value_per_contractor'] = columns[2].find_element(By.CSS_SELECTOR, "p").text.strip()
                    except:
                        award['value_per_contractor'] = ''
                    
                    try:
                        award['hub_status'] = columns[3].find_element(By.CSS_SELECTOR, "p").text.strip()
                    except:
                        award['hub_status'] = ''
                    
                    try:
                        award['award_date'] = columns[4].find_element(By.CSS_SELECTOR, "p").text.strip()
                    except:
                        award['award_date'] = ''
                    
                    try:
                        award['award_status'] = columns[5].find_element(By.CSS_SELECTOR, "p").text.strip()
                    except:
                        award['award_status'] = ''
                    
                    # Only add if we got at least the contractor name
                    if award['contractor']:
                        awards_data.append(award)
                        print(f"Extracted award: {award['contractor']}")
                
            except Exception as e:
                print(f"Error extracting award row {idx}: {str(e)[:100]}")
                continue
                
    except Exception as e:
        print(f"Error extracting awards: {str(e)[:100]}")
    
    return awards_data


def extract_attachment_names(driver):
    """Extract attachment names without downloading"""
    attachment_names = []
    
    try:
        # Look for attachment links
        download_links = driver.find_elements(
            By.CSS_SELECTOR, "a[data-action='downloadURL']"
        )
        
        for link in download_links:
            try:
                att_name = link.text.strip()
                if att_name:
                    attachment_names.append(att_name)
            except:
                continue
                
    except Exception as e:
        print(f"Error extracting attachment names: {str(e)[:50]}")
    
    return attachment_names


def worker_process(worker_id, task_queue, result_queue, download_dir):
    """Worker process that handles scraping attachment names for opportunities"""
    
    # Create worker-specific download directory (not used but kept for future)
    worker_download_dir = os.path.join(download_dir, f"temp_worker_{worker_id}")
    os.makedirs(worker_download_dir, exist_ok=True)
    
    # Setup Chrome driver for this worker
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": os.path.abspath(worker_download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_setting_values.automatic_downloads": 1
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    
    driver = None
    
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        print(f"[Worker {worker_id}] Started and ready")
        
        while True:
            try:
                # Get task from queue with timeout
                task = task_queue.get(timeout=10)
                
                if task is None:  # Poison pill to stop worker
                    print(f"[Worker {worker_id}] Received stop signal")
                    break
                
                opp = task
                print(f"[Worker {worker_id}] Processing: {opp['title'][:50]}...")
                
                try:
                    # Navigate to opportunity page
                    driver.get(opp['href'])
                    time.sleep(3)
                    
                    # Extract contact information
                    contact_info = extract_contact_info(driver)
                    opp.update(contact_info)
                    print(f"[Worker {worker_id}] ✓ Extracted contact info")
                    
                    # Extract awards information
                    awards_data = extract_awards_info(driver)
                    
                    # If multiple awards, we'll store them as pipe-separated values
                    if awards_data:
                        opp['awards_count'] = len(awards_data)
                        # Store first award details or aggregate
                        opp['contractor'] = ' | '.join([a['contractor'] for a in awards_data])
                        opp['mailing_address'] = ' | '.join([a['mailing_address'] for a in awards_data])
                        opp['value_per_contractor'] = ' | '.join([a['value_per_contractor'] for a in awards_data])
                        opp['hub_status'] = ' | '.join([a['hub_status'] for a in awards_data])
                        opp['award_date'] = ' | '.join([a['award_date'] for a in awards_data])
                        opp['award_status'] = ' | '.join([a['award_status'] for a in awards_data])
                        print(f"[Worker {worker_id}] ✓ Extracted {len(awards_data)} award(s)")
                    else:
                        opp['awards_count'] = 0
                        opp['contractor'] = ''
                        opp['mailing_address'] = ''
                        opp['value_per_contractor'] = ''
                        opp['hub_status'] = ''
                        opp['award_date'] = ''
                        opp['award_status'] = ''
                        print(f"[Worker {worker_id}] No awards found")
                    
                    # Extract attachment names (NEW - replaces download functionality)
                    attachment_names = extract_attachment_names(driver)
                    
                    if attachment_names:
                        opp['attachment_count'] = len(attachment_names)
                        opp['attachments'] = ' | '.join(attachment_names)
                        print(f"[Worker {worker_id}] ✓ Found {len(attachment_names)} attachment(s)")
                    else:
                        opp['attachment_count'] = 0
                        opp['attachments'] = ''
                        print(f"[Worker {worker_id}] No attachments found")
                    
                    # COMMENTED OUT: Download functionality
                    # # Track existing files before downloading
                    existing_files = set(os.path.join(worker_download_dir, f) 
                                       for f in os.listdir(worker_download_dir))
                    
                    # Look for attachment links
                    download_links = driver.find_elements(
                        By.CSS_SELECTOR, "a[data-action='downloadURL']"
                    )
                    
                    if not download_links:
                        print(f"[Worker {worker_id}] No attachments found")
                        opp['attachment_count'] = 0
                        opp['zip_file'] = ""
                        result_queue.put(opp)
                        continue
                    
                    print(f"[Worker {worker_id}] Found {len(download_links)} attachments")
                    opp['attachment_count'] = len(download_links)
                    downloaded_files = []
                    
                    # Download each attachment
                    for att_idx, link in enumerate(download_links, 1):
                        try:
                            att_name = link.text.strip()
                            print(f"[Worker {worker_id}] Downloading {att_idx}/{len(download_links)}: {att_name[:30]}...")
                            
                            driver.execute_script("arguments[0].click();", link)
                            
                            if wait_for_download(worker_download_dir, timeout=60):
                                latest_file = get_latest_file(worker_download_dir, existing_files)
                                if latest_file:
                                    downloaded_files.append(latest_file)
                                    existing_files.add(latest_file)
                                    print(f"[Worker {worker_id}] ✓ Downloaded")
                                else:
                                    print(f"[Worker {worker_id}] ✗ Could not find downloaded file")
                            else:
                                print(f"[Worker {worker_id}] ✗ Download timeout")
                            
                            time.sleep(2)
                            
                        except Exception as e:
                            print(f"[Worker {worker_id}] ✗ Attachment error: {str(e)[:50]}")
                    
                    # Create zip file for this opportunity
                    if downloaded_files:
                        safe_title = sanitize_filename(opp['title'] or opp['solicitation_id'])
                        zip_filename = f"{safe_title}_w{worker_id}.zip"
                        zip_path = os.path.join(download_dir, zip_filename)
                        
                        # Handle duplicate zip names
                        counter = 1
                        while os.path.exists(zip_path):
                            zip_filename = f"{safe_title}_w{worker_id}_{counter}.zip"
                            zip_path = os.path.join(download_dir, zip_filename)
                            counter += 1
                        
                        create_zip(downloaded_files, zip_path)
                        opp['zip_file'] = zip_filename
                        print(f"[Worker {worker_id}] ✓ Created zip: {zip_filename} ({len(downloaded_files)} files)")
                        
                        # Clean up downloaded files
                        for file in downloaded_files:
                            try:
                                os.remove(file)
                            except:
                                pass
                    else:
                        opp['zip_file'] = ""
                    
                    result_queue.put(opp)
                    
                except Exception as e:
                    print(f"[Worker {worker_id}] ✗ Error processing opportunity: {str(e)[:100]}")
                    opp['attachment_count'] = 0
                    opp['attachments'] = ''
                    opp['error'] = str(e)[:200]
                    result_queue.put(opp)
                    
            except Empty:
                continue
            except Exception as e:
                print(f"[Worker {worker_id}] ✗ Unexpected error: {str(e)}")
                traceback.print_exc()
                
    except Exception as e:
        print(f"[Worker {worker_id}] ✗ Fatal error: {str(e)}")
        traceback.print_exc()
        
    finally:
        if driver:
            driver.quit()
        
        # Clean up worker directory
        try:
            for file in os.listdir(worker_download_dir):
                os.remove(os.path.join(worker_download_dir, file))
            os.rmdir(worker_download_dir)
        except:
            pass
        
        print(f"[Worker {worker_id}] Stopped")


def scrape_and_download(download_dir="downloads", max_pages=None, num_workers=2):
    """Main function with multiprocessing support"""
    os.makedirs(download_dir, exist_ok=True)
    
    # Create queues for task distribution
    task_queue = Queue(maxsize=20)  # Limit queue size to prevent memory issues
    result_queue = Queue()
    
    # Start worker processes
    workers = []
    for i in range(num_workers):
        worker = Process(target=worker_process, args=(i, task_queue, result_queue, download_dir))
        worker.start()
        workers.append(worker)
    
    print(f"Started {num_workers} worker processes")
    
    # Setup main driver for page navigation
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(driver, 20)
    
    all_opportunities_data = []
    total_opportunities_found = 0

    try:
        print(f"\n{'='*60}")
        print("STARTING SCRAPING PROCESS")
        print(f"{'='*60}")
        
        base_url = "https://www.txsmartbuy.gov"
        url = f"{base_url}/esbd?status=2&dateRange=lastFiscalYear&startDate=09%2F01%2F2024&endDate=08%2F31%2F2025"
        driver.get(url)

        # Apply filters
        status_dropdown = Select(wait.until(EC.presence_of_element_located((By.NAME, "status"))))
        status_dropdown.select_by_value("2")

        date_range_dropdown = Select(wait.until(EC.presence_of_element_located((By.NAME, "dateRange"))))
        date_range_dropdown.select_by_value("lastFiscalYear")

        search_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))
        search_btn.click()
        time.sleep(5)

        page_num = 1
        
        while True:
            if max_pages and page_num > max_pages:
                break
                
            print(f"\n{'='*60}")
            print(f"Processing Page {page_num}")
            print(f"{'='*60}")
            
            time.sleep(3)
            
            # Get all opportunity rows on current page
            opportunity_rows = driver.find_elements(By.CSS_SELECTOR, "div.esbd-result-row")
            
            print(f"Found {len(opportunity_rows)} opportunities on page {page_num}")
            
            # Extract data from each opportunity row
            for row in opportunity_rows:
                try:
                    # Extract title and link
                    title_elem = row.find_element(By.CSS_SELECTOR, "div.esbd-result-title a")
                    title = title_elem.text.strip()
                    href = title_elem.get_attribute('href')
                    
                    # Extract other details using the new function
                    solicitation_id = extract_field_value(row, "Solicitation ID:")
                    due_date = extract_field_value(row, "Due Date:")
                    due_time = extract_field_value(row, "Due Time:")
                    agency = extract_field_value(row, "Agency/Texas SmartBuy Member Number:")
                    status = extract_field_value(row, "Status:")
                    posting_date = extract_field_value(row, "Posting Date:")
                    created_date = extract_field_value(row, "Created Date:")
                    last_updated = extract_field_value(row, "Last Updated:")
                    
                    opp_data = {
                        'title': title,
                        'href': href,
                        'solicitation_id': solicitation_id,
                        'due_date': due_date,
                        'due_time': due_time,
                        'agency': agency,
                        'status': status,
                        'posting_date': posting_date,
                        'created_date': created_date,
                        'last_updated': last_updated
                    }
                    
                    # Add to task queue for workers to process
                    total_opportunities_found += 1
                    print(f"[Main] Queuing opportunity {total_opportunities_found}: {title[:50]}")
                    task_queue.put(opp_data)
                    
                except Exception as e:
                    print(f"[Main] ✗ Error extracting row data: {str(e)[:50]}")
            
            # Increment page number
            page_num += 1
            
            # Check if we've reached the page limit BEFORE trying to go to next page
            if max_pages and page_num > max_pages:
                print(f"\n{'='*60}")
                print(f"Reached page limit ({max_pages} pages)")
                print(f"{'='*60}")
                break
            
            # Try to go to next page
            try:
                next_button = driver.find_element(By.CSS_SELECTOR, "a#Next[aria-label='Next']")
                next_class = next_button.get_attribute('class') or ""
                
                if 'disabled' in next_class or not next_button.is_enabled():
                    print(f"\n{'='*60}")
                    print("No more pages available")
                    print(f"{'='*60}")
                    break
                
                print(f"\n{'='*60}")
                print(f"Moving to page {page_num}")
                print(f"{'='*60}")
                
                driver.execute_script("arguments[0].click();", next_button)
                time.sleep(5)
                
            except NoSuchElementException:
                print(f"\n{'='*60}")
                print("Reached last page")
                print(f"{'='*60}")
                break

        print(f"\n[Main] Finished queuing {total_opportunities_found} opportunities")
        print(f"[Main] Waiting for workers to complete...")
        
        # Send stop signal to all workers
        for _ in range(num_workers):
            task_queue.put(None)
        
        # Collect results from workers
        collected = 0
        while collected < total_opportunities_found:
            try:
                result = result_queue.get(timeout=60)
                all_opportunities_data.append(result)
                collected += 1
                print(f"[Main] Collected {collected}/{total_opportunities_found} results")
            except Empty:
                print(f"[Main] Timeout waiting for results. Collected {collected}/{total_opportunities_found}")
                break

        # Wait for all workers to finish
        for worker in workers:
            worker.join(timeout=30)
            if worker.is_alive():
                print(f"[Main] Warning: Worker still alive, terminating...")
                worker.terminate()

        # Save all scraped data to Excel
        if all_opportunities_data:
            # Remove columns we don't want in Excel
            for opp in all_opportunities_data:
                opp.pop('href', None)
                opp.pop('error', None)
            
            # Reorder columns for better readability
            column_order = [
                'title', 'solicitation_id', 'status', 
                'due_date', 'due_time', 'response_due_date',
                'agency', 'posting_date', 'created_date', 'last_updated',
                'contact_name', 'contact_number', 'contact_email',
                'awards_count', 'contractor', 'mailing_address', 
                'value_per_contractor', 'hub_status', 'award_date', 'award_status',
                'attachment_count', 'attachments'
            ]
            
            df = pd.DataFrame(all_opportunities_data)
            # Reorder columns (keep any extra columns at the end)
            existing_cols = [col for col in column_order if col in df.columns]
            other_cols = [col for col in df.columns if col not in column_order]
            df = df[existing_cols + other_cols]
            
            excel_file = os.path.join(download_dir, "opportunities_data.xlsx")
            df.to_excel(excel_file, index=False)
            print(f"\n✓ Saved data for {len(all_opportunities_data)} opportunities to: {excel_file}")

            # Save as JSON as well
            json_file = os.path.join(download_dir, "opportunities_data.json")
            df.to_json(json_file, orient="records", force_ascii=False, indent=2)
            print(f"✓ Saved data as JSON to: {json_file}")

    except Exception as e:
        print(f"\n✗ Critical error: {str(e)}")
        traceback.print_exc()
        
    finally:
        driver.quit()
        
        # Ensure all workers are terminated
        for worker in workers:
            if worker.is_alive():
                worker.terminate()
                worker.join()
        
        print(f"\n{'='*60}")
        print("SCRAPING COMPLETED!")
        print(f"{'='*60}")
        if all_opportunities_data:
            print(f"📊 Data file: {os.path.abspath(download_dir)}/opportunities_data.xlsx")
            print(f"📊 Total opportunities: {len(all_opportunities_data)}")
        print(f"{'='*60}")


if __name__ == "__main__":
    # Set max_pages=None to process all pages, or set a number to limit
    # num_workers=2 is safer, use 3 if you have good CPU/memory
    scrape_and_download(download_dir="downloads", max_pages=None, num_workers=2)