from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import json
import time

class IdahoContractsScraper:
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.wait = WebDriverWait(self.driver, 10)
        self.all_contracts_data = []
        
        # Year configurations with their table and dropdown IDs
        self.year_configs = [
            {"year": "2025", "table_id": "tablepress-231", "dropdown_id": "dt-length-0"},
            {"year": "2024", "table_id": "tablepress-229", "dropdown_id": "dt-length-1"},
            {"year": "2023", "table_id": "tablepress-233", "dropdown_id": "dt-length-2"},
            {"year": "2022", "table_id": "tablepress-151", "dropdown_id": "dt-length-3"},
            {"year": "2021", "table_id": "tablepress-108", "dropdown_id": "dt-length-4"},
            {"year": "2020", "table_id": "tablepress-26", "dropdown_id": "dt-length-5"},
            {"year": "2019", "table_id": "tablepress-27", "dropdown_id": "dt-length-6"},
            {"year": "2018", "table_id": "tablepress-28", "dropdown_id": "dt-length-7"},
            {"year": "2017", "table_id": "tablepress-29", "dropdown_id": "dt-length-8"},
            {"year": "2016", "table_id": "tablepress-30", "dropdown_id": "dt-length-9"},
        ]
        
    def scrape_contracts(self):
        try:
            # Navigate to the website
            print("Opening Idaho Department of Lands website...")
            self.driver.get("https://www.idl.idaho.gov/contracting-bid-board/contracts-awarded/")
            time.sleep(3)
            
            # Process each year
            for idx, config in enumerate(self.year_configs):
                # Scroll down for years after 2025
                if idx > 0:
                    print(f"\nScrolling to {config['year']} section...")
                    self.driver.execute_script("window.scrollBy(0, 800);")
                    time.sleep(2)
                
                # Scrape the year's table
                print("\n" + "="*60)
                print(f"Processing {config['year']} Contracts")
                print("="*60)
                self.scrape_year_table(
                    year=config['year'],
                    table_id=config['table_id'],
                    dropdown_id=config['dropdown_id']
                )
            
            # Save to Excel and JSON
            self.save_to_excel()
            self.save_to_json()
            
        except Exception as e:
            print(f"Error during scraping: {str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            print("\nClosing browser...")
            self.driver.quit()
    
    def scrape_year_table(self, year, table_id, dropdown_id):
        """Scrape contracts table for a specific year"""
        try:
            # Find and select dropdown to show 100 entries
            print(f"Setting dropdown to show 100 entries for {year}...")
            dropdown = self.wait.until(
                EC.presence_of_element_located((By.ID, dropdown_id))
            )
            select = Select(dropdown)
            select.select_by_value("100")
            time.sleep(3)  # Wait for table to reload
            
            # Find the table
            print(f"Finding table for {year}...")
            table = self.wait.until(
                EC.presence_of_element_located((By.ID, table_id))
            )
            
            # Find all rows in tbody
            rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
            print(f"Found {len(rows)} contracts in {year}")
            
            # Process each row
            for index, row in enumerate(rows, 1):
                try:
                    print(f"Processing contract {index} of {len(rows)} ({year})...")
                    contract_data = self.extract_row_data(row, year)
                    if contract_data:
                        self.all_contracts_data.append(contract_data)
                except Exception as e:
                    print(f"Error processing row {index} in {year}: {str(e)}")
                    continue
            
            print(f"Completed scraping {year} contracts")
            
        except Exception as e:
            print(f"Error scraping {year} table: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def extract_row_data(self, row, year):
        """Extract data from a single table row"""
        try:
            cells = row.find_elements(By.TAG_NAME, "td")
            
            if len(cells) < 4:
                print("Row doesn't have enough columns, skipping...")
                return None
            
            # Extract text data from columns
            project = cells[0].text.strip()
            description = cells[1].text.strip()
            awarded_to = cells[3].text.strip()
            
            # Extract link from Responsive Vendors column (column 3, index 2)
            responsive_vendors_cell = cells[2]
            evaluation_link = ""
            evaluation_text = responsive_vendors_cell.text.strip()
            
            try:
                link_element = responsive_vendors_cell.find_element(By.TAG_NAME, "a")
                evaluation_link = link_element.get_attribute("href")
                print(f"Found evaluation link: {evaluation_link}")
            except NoSuchElementException:
                print("No evaluation link found in this row")
            
            return {
                'Year': year,
                'Project': project,
                'Description': description,
                'Responsive Vendors Text': evaluation_text,
                'Evaluation Link': evaluation_link,
                'Awarded To': awarded_to
            }
            
        except Exception as e:
            print(f"Error extracting row data: {str(e)}")
            return None
    
    def save_to_excel(self):
        """Save scraped data to Excel file"""
        if not self.all_contracts_data:
            print("No data to save!")
            return
        
        # Create DataFrame
        df = pd.DataFrame(self.all_contracts_data)
        
        # Reorder columns for better readability
        column_order = ['Year', 'Project', 'Description', 'Responsive Vendors Text', 
                       'Evaluation Link', 'Awarded To']
        df = df[column_order]
        
        # Create filename with timestamp
        filename = f'idaho_contracts_awarded_{time.strftime("%Y%m%d_%H%M%S")}.xlsx'
        
        # Save to Excel
        df.to_excel(filename, index=False, engine='openpyxl')
        
        print(f"\n{'='*60}")
        print(f"📊 Excel file saved: {filename}")
        print(f"Total contracts scraped: {len(df)}")
        print(f"\nBreakdown by year:")
        year_counts = df['Year'].value_counts().sort_index(ascending=False)
        for year, count in year_counts.items():
            print(f"  {year}: {count} contracts")
        print(f"{'='*60}")
    
    def save_to_json(self):
        """Save scraped data to JSON file"""
        if not self.all_contracts_data:
            print("No data to save!")
            return
        
        # Create filename with timestamp
        filename = f'idaho_contracts_awarded_{time.strftime("%Y%m%d_%H%M%S")}.json'
        
        # Save to JSON with proper formatting
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(self.all_contracts_data, f, indent=2, ensure_ascii=False)
        
        print(f"📄 JSON file saved: {filename}")
        print(f"Total contracts in JSON: {len(self.all_contracts_data)}")

if __name__ == "__main__":
    scraper = IdahoContractsScraper()
    scraper.scrape_contracts()