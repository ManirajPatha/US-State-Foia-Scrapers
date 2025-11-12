import os
import time
import zipfile
import json
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
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
        time.sleep(2)
        try:
            self.wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        except:
            pass

    def apply_awarded_filter(self, retries=3):
        for attempt in range(retries):
            try:
                self.wait_for_page_load()
                
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
                time.sleep(2)

                print("Filter applied: Awarded = Yes")
                return True
            except StaleElementReferenceException:
                if attempt < retries - 1:
                    print(f"Retrying filter application (attempt {attempt + 2}/{retries})...")
                    time.sleep(2)
                    continue
                else:
                    print(f"Error applying filter after {retries} attempts")
                    return False
            except Exception as e:
                print(f"Error applying filter: {e}")
                return False
        return False

    def navigate_to_page(self, page_index, retries=3):
        for attempt in range(retries):
            try:
                if page_index == 0:
                    return True
                print(f"Navigating to page {page_index + 1}...")
                
                self.wait.until(EC.presence_of_element_located((By.ID, "body_x_grid_grd")))
                time.sleep(1)
                
                script = f"__ivCtrl['body_x_grid_grd'].GoToPageOfGrid(0, {page_index});"
                self.driver.execute_script(script)
                self.wait_for_page_load()
                time.sleep(3)
                
                self.current_page = page_index
                print(f"Now on page {page_index + 1}")
                return True
            except Exception as e:
                if attempt < retries - 1:
                    print(f"Retrying navigation (attempt {attempt + 2}/{retries})...")
                    time.sleep(2)
                    continue
                else:
                    print(f"Error navigating to page {page_index + 1}: {e}")
                    return False
        return False

    def has_next_page(self):
        try:
            next_button = self.driver.find_element(By.ID, "body_x_grid_PagerBtnNextPage")
            is_disabled = 'disable' in next_button.get_attribute('class').lower()
            aria_disabled = next_button.get_attribute('aria-disabled') == 'true'
            return not (is_disabled or aria_disabled)
        except:
            return False

    def go_to_next_page(self, retries=3):
        for attempt in range(retries):
            try:
                self.wait.until(EC.presence_of_element_located((By.ID, "body_x_grid_grd")))
                time.sleep(1)
                
                next_button = self.wait.until(
                    EC.presence_of_element_located((By.ID, "body_x_grid_PagerBtnNextPage"))
                )

                if 'disable' in next_button.get_attribute('class').lower():
                    print("Next page button is disabled - no more pages")
                    return False
                
                # Get the next page index before clicking
                next_page_index = int(next_button.get_attribute('data-page-index'))
                
                self.driver.execute_script("arguments[0].click();", next_button)
                self.wait_for_page_load()
                time.sleep(3)
                
                self.current_page = next_page_index
                print(f"Moved to page {self.current_page + 1}")
                return True
                
            except StaleElementReferenceException:
                if attempt < retries - 1:
                    print(f"Retrying next page (attempt {attempt + 2}/{retries})...")
                    time.sleep(2)
                    continue
                else:
                    print(f"Error going to next page after {retries} attempts")
                    return False
            except Exception as e:
                if attempt < retries - 1:
                    print(f"Retrying next page due to error (attempt {attempt + 2}/{retries})...")
                    time.sleep(2)
                    continue
                else:
                    print(f"Error going to next page: {e}")
                    return False
        return False

    def get_all_opportunity_links(self):
        try:
            self.wait.until(EC.presence_of_element_located((By.ID, "body_x_grid_grd")))
            time.sleep(2)
            links = self.driver.find_elements(
                By.XPATH,
                "//a[contains(@id, '_img___colManagegrid') and contains(@class, 'iv-button')]"
            )
            urls = [link.get_attribute('href') for link in links if link.get_attribute('href')]
            print(f"Found {len(urls)} opportunities on this page")
            return urls
        except Exception as e:
            print(f"Error getting opportunity links: {e}")
            return []

    def scrape_awarded_suppliers(self):
        try:
            try:
                table = self.driver.find_element(
                    By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_grdSupplierResponse_grd"
                )
            except NoSuchElementException:
                print("No Public Supplier Response table found")
                return []

            awarded_suppliers = []
            rows = table.find_elements(By.XPATH, ".//tbody/tr")

            for row in rows:
                try:
                    row.find_element(By.XPATH, ".//input[@type='checkbox' and @checked='checked']")
                    cells = row.find_elements(By.XPATH, ".//td[@data-iv-role='cell']")
                    if len(cells) >= 3:
                        awarded_suppliers.append({
                            'Supplier Name': cells[0].text.strip(),
                            'Item': cells[1].text.strip(),
                            'Submitted Unit Price': cells[2].text.strip()
                        })
                except NoSuchElementException:
                    continue

            print(f"Found {len(awarded_suppliers)} awarded supplier(s)")
            return awarded_suppliers
        except Exception as e:
            print(f"Error scraping awarded suppliers: {e}")
            return []

    def scrape_opportunity_details(self):
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

            data = {
                'Solicitation ID': solicitation_id,
                'Solicitation Name': solicitation_name,
                'Begin Date': begin_date,
                'End Date': end_date,
                'Solicitation Status': status
            }
            print(f"Scraped: {solicitation_id} - {solicitation_name}")
            return data
        except Exception as e:
            print(f"Error scraping opportunity details: {e}")
            return None

    def download_documents(self, solicitation_id):
        try:
            rows = self.driver.find_elements(
                By.XPATH,
                "//table[contains(@id, '_proxy_rfp_') and contains(@id, '_grid_grd')]//tbody/tr"
            )
            if not rows:
                print(f"No documents found for {solicitation_id}")
                return None

            print(f"Found {len(rows)} document(s)...")

            doc_folder = os.path.join(self.download_path, f"{solicitation_id}_docs")
            os.makedirs(doc_folder, exist_ok=True)

            downloaded_files = []
            attachments_info = []

            for i, row in enumerate(rows, 1):
                try:
                    title_cell = row.find_element(By.XPATH, ".//td[@data-iv-role='cell'][1]")
                    doc_title = title_cell.text.strip() or f"document_{i}"
                    link = row.find_element(
                        By.XPATH,
                        ".//a[contains(@class, 'iv-download-file') and contains(@href, '/bare.aspx/en/fil/download_public/')]"
                    )

                    file_name = f"{doc_title}.pdf"
                    self.driver.execute_script("arguments[0].click();", link)
                    time.sleep(3)
                    self.wait_for_download_complete()

                    latest_file = self.get_latest_file(self.download_path)
                    if latest_file:
                        new_path = os.path.join(doc_folder, file_name)
                        os.rename(latest_file, new_path)
                        downloaded_files.append(new_path)
                        attachments_info.append({
                            "Title": doc_title,
                            "File Name": file_name
                        })
                        print(f"Downloaded: {file_name}")
                except Exception as e:
                    print(f"Error processing document {i}: {e}")
                    continue

            if downloaded_files:
                zip_filename = os.path.join(self.download_path, f"{solicitation_id}_documents.zip")
                with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for f in downloaded_files:
                        zipf.write(f, os.path.basename(f))

                import shutil
                shutil.rmtree(doc_folder)
                print(f"Created zip: {zip_filename}")
                return {"Zip File": zip_filename, "Attachments": attachments_info}

            return None
        except Exception as e:
            print(f"Error downloading documents: {e}")
            return None

    def wait_for_download_complete(self, timeout=60):
        end_time = time.time() + timeout
        while time.time() < end_time:
            if not any(f.endswith(('.crdownload', '.tmp')) for f in os.listdir(self.download_path)):
                time.sleep(1)
                return True
            time.sleep(0.5)
        return False

    def get_latest_file(self, folder):
        files = [os.path.join(folder, f) for f in os.listdir(folder)
                 if os.path.isfile(os.path.join(folder, f)) and not f.endswith('.crdownload')]
        return max(files, key=os.path.getctime) if files else None

    def return_to_list_page_with_filters(self, retries=3):
        for attempt in range(retries):
            try:
                print(f"Returning to list page {self.current_page + 1}...")
                
                # Go back to the list page
                self.driver.get("https://ohiobuys.ohio.gov/page.aspx/en/rfp/request_browse_public")
                self.wait_for_page_load()
                time.sleep(2)
                
                # Reapply the filter
                if not self.apply_awarded_filter():
                    if attempt < retries - 1:
                        print(f"Retrying return to list (attempt {attempt + 2}/{retries})...")
                        time.sleep(3)
                        continue
                    return False
                
                if self.current_page > 0:
                    if not self.navigate_to_page(self.current_page):
                        if attempt < retries - 1:
                            print(f"Retrying return to list (attempt {attempt + 2}/{retries})...")
                            time.sleep(3)
                            continue
                        return False
                
                print(f"Successfully returned to page {self.current_page + 1}")
                return True
                
            except Exception as e:
                if attempt < retries - 1:
                    print(f"Retrying return to list due to error (attempt {attempt + 2}/{retries})...")
                    time.sleep(3)
                    continue
                else:
                    print(f"Error returning to list page after {retries} attempts: {e}")
                    return False
        return False

    def save_to_json(self, filename="ohiobuys_awarded_solicitations.json"):
        if not self.data:
            print("No data to save!")
            return
        path = os.path.join(self.download_path, filename)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)
        print(f"Data saved to: {path}")

    def save_to_excel(self, filename="ohiobuys_awarded_solicitations.xlsx"):
        if not self.data:
            print("No data to save!")
            return
        all_rows = []
        for record in self.data:
            base = {
                'Solicitation ID': record.get('Solicitation ID', ''),
                'Solicitation Name': record.get('Solicitation Name', ''),
                'Begin Date': record.get('Begin Date', ''),
                'End Date': record.get('End Date', ''),
                'Solicitation Status': record.get('Solicitation Status', ''),
                'Documents Zip': record.get('Documents Zip', '')
            }
            suppliers = record.get('Awarded Suppliers', [])
            attachments = record.get('Attachments', [])
            
            # Handle case where there are no suppliers or attachments
            if not suppliers:
                suppliers = [{'Supplier Name': '', 'Item': '', 'Submitted Unit Price': ''}]
            if not attachments:
                attachments = [{'Title': '', 'File Name': ''}]
            
            for supplier in suppliers:
                for attach in attachments:
                    row = base.copy()
                    row.update({
                        'Supplier Name': supplier.get('Supplier Name', ''),
                        'Item': supplier.get('Item', ''),
                        'Submitted Unit Price': supplier.get('Submitted Unit Price', ''),
                        'Attachment Title': attach.get('Title', ''),
                        'Attachment File': attach.get('File Name', '')
                    })
                    all_rows.append(row)
        df = pd.DataFrame(all_rows)
        path = os.path.join(self.download_path, filename)
        df.to_excel(path, index=False)
        print(f"Data saved to: {path}")
        print(f"  Total records: {len(all_rows)}")

    def scrape_all_opportunities(self, max_pages=None):
        try:
            print("Opening OhioBuys website...")
            self.driver.get("https://ohiobuys.ohio.gov/page.aspx/en/rfp/request_browse_public")
            self.wait_for_page_load()

            if not self.apply_awarded_filter():
                print("Failed to apply filter. Exiting.")
                return

            total = 0
            consecutive_failures = 0
            max_consecutive_failures = 3
            
            while True:
                if max_pages and self.current_page >= max_pages:
                    print(f"\nReached maximum page limit ({max_pages})")
                    break

                urls = self.get_all_opportunity_links()
                if not urls:
                    print("No opportunities found on this page")
                    break

                page_success = False
                for idx, url in enumerate(urls, 1):
                    print(f"\n[Page {self.current_page + 1} - {idx}/{len(urls)}] Processing...")
                    self.driver.get(url)
                    self.wait_for_page_load()

                    data = self.scrape_opportunity_details()
                    if data:
                        suppliers = self.scrape_awarded_suppliers()
                        data['Awarded Suppliers'] = suppliers
                        doc_info = self.download_documents(data['Solicitation ID'])
                        if doc_info:
                            data['Attachments'] = doc_info['Attachments']
                            data['Documents Zip'] = doc_info['Zip File']
                        self.data.append(data)
                        total += 1
                        page_success = True

                    if not self.return_to_list_page_with_filters():
                        print("Failed to return to list page, attempting to continue...")
                        consecutive_failures += 1
                        if consecutive_failures >= max_consecutive_failures:
                            print(f"Too many consecutive failures ({max_consecutive_failures}), stopping...")
                            break
                        # Try to recover by going back to the main page
                        self.driver.get("https://ohiobuys.ohio.gov/page.aspx/en/rfp/request_browse_public")
                        self.wait_for_page_load()
                        if not self.apply_awarded_filter():
                            break
                        if self.current_page > 0:
                            self.navigate_to_page(self.current_page)
                    else:
                        consecutive_failures = 0

                if consecutive_failures >= max_consecutive_failures:
                    break

                # Check if there's a next page
                if not self.has_next_page():
                    print("\nNo more pages available.")
                    break
                    
                # Go to next page
                if not self.go_to_next_page():
                    print("\nFailed to navigate to next page")
                    break

            print(f"Total pages processed: {self.current_page + 1}")
            print(f"Total opportunities scraped: {total}")

            self.save_to_excel()
            self.save_to_json()
        except Exception as e:
            print(f"Critical error: {e}")
            import traceback
            traceback.print_exc()
            if self.data:
                print("\n⚠ Saving partial data...")
                self.save_to_excel()
                self.save_to_json()

    def close(self):
        if self.driver:
            self.driver.quit()
            print("\nBrowser closed")


def main():
    scraper = OhioBuysScraper(download_path="downloads")
    try:
        scraper.scrape_all_opportunities(max_pages=None)
    except KeyboardInterrupt:
        print("\nInterrupted by user")
        scraper.save_to_excel()
        scraper.save_to_json()
    finally:
        scraper.close()


if __name__ == "__main__":
    main()