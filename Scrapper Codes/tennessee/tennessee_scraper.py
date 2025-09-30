#!/usr/bin/env python3
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def scrape_awards_from_index(driver, index_url):
    driver.get(index_url)
    time.sleep(2)
    awards = []
    
    links = driver.find_elements(By.XPATH, "//a[contains(text(),'Intent to Award')]")
    for link in links:
        award_url = link.get_attribute('href')
        
        row = link.find_element(By.XPATH, "./ancestor::tr")
        cells = row.find_elements(By.TAG_NAME, "td")
        event_id = cells[0].text.strip() if len(cells) > 0 else ''
        name     = cells[2].text.strip() if len(cells) > 2 else ''
        awards.append({
            "Event ID":    event_id,
            "Name":        name,
            "Award PDF":   award_url,
            "Index Page":  index_url
        })
    return awards

def main():
    service = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')

    driver = webdriver.Chrome(service=service, options=options)

    index_pages = [
        "https://www.tn.gov/generalservices/procurement/central-procurement-office--cpo-/supplier-information/invitations-to-bid--itb-.html",
        "https://www.tn.gov/generalservices/procurement/central-procurement-office--cpo-/supplier-information/request-for-proposals--rfp--opportunities1.html"
    ]

    all_awards = []
    for page in index_pages:
        all_awards.extend(scrape_awards_from_index(driver, page))

    driver.quit()

    df = pd.DataFrame(all_awards)
    df.to_excel("tn_awards.xlsx", index=False)
    print(f"Saved {len(df)} award entries to tn_awards.xlsx")

if __name__ == "__main__":
    main()