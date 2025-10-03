# Required libraries:
# pip install requests pandas openpyxl

import logging
import time
from datetime import datetime
from typing import Optional, List, Dict, Any, Tuple
import requests
import json
from urllib.parse import quote
import pandas as pd


logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


def classify_opportunity(title: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
    """
    Classifies an opportunity based on its title.
    Returns (industry, sub_sector).
    """
    if not title:
        return ("Miscellaneous", "General")

    title_lower = title.lower()
    keywords = {
        "Construction": ["construction", "building", "renovation", "roofing", "hvac"],
        "Technology": ["software", "hardware", "it services", "cybersecurity", "network"],
        "Consulting": ["consulting", "consultant", "professional services", "study"],
        "Transportation": ["vehicles", "fleet", "automotive", "trucks", "transportation"],
        "Medical": ["medical", "health", "pharmaceutical", "ppe", "hospital"],
    }

    for industry, terms in keywords.items():
        if any(term in title_lower for term in terms):
            return (industry, "General")

    return ("Miscellaneous", "General")


def parse_date_from_api(api_date: Optional[int]) -> Optional[str]:
    """
    Parses a Unix timestamp (in milliseconds) from the API into a date string (YYYY-MM-DD).
    """
    if not api_date:
        return None
    try:
        return datetime.fromtimestamp(api_date / 1000).date().isoformat()
    except (ValueError, TypeError) as e:
        logging.warning(f"Could not parse date from API data: '{api_date}'. Error: {e}")
        return None


def scrape_connecticut_awarded():
    """
    Scrapes AWARDED bidding opportunities from the Connecticut (CTSource)
    procurement portal and saves them to an Excel file for the current year.
    """
    CUSTOMER_ID = "51"
    SOURCE_NAME = "Connecticut"
    scraped_data: List[Dict[str, Any]] = []

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json, text/plain, */*",
        "Referer": "https://webprocure.proactiscloud.com/wp-web-public/en/",
    }

    list_url = "https://webprocure.proactiscloud.com/wp-full-text-search/search/sols"
    list_params = {
        "customerid": CUSTOMER_ID,
        "q": "*",
        "sort": "r",
        "f": "ps=Awarded",  
        "oids": "",
    }

    download_base_url = "https://webprocure.proactiscloud.com/main/sol/viewdoc.do"

    try:
        with requests.Session() as session:
            session.headers.update(headers)
            current_offset = 0
            total_hits = 0
            is_first_page = True

            while True:
                logging.info(f"Fetching page of awarded bids with offset {current_offset}...")
                list_params["from"] = current_offset

                try:
                    list_response = session.get(list_url, params=list_params, timeout=60)
                    list_response.raise_for_status()
                    list_data = list_response.json()
                except requests.RequestException as e:
                    logging.error(f"Failed to fetch page at offset {current_offset}: {e}. Stopping scrape.")
                    break

                if is_first_page:
                    total_hits = list_data.get("hits", 0)
                    if total_hits == 0:
                        logging.info("No awarded solicitations found.")
                        break
                    logging.info(f"Found {total_hits} total awarded solicitations. Starting scrape...")
                    is_first_page = False

                records_on_page = list_data.get("records", [])
                if not records_on_page:
                    logging.info("Finished processing all pages.")
                    break

                for record in records_on_page:
                    bid_id = record.get("bidid")
                    notice_id = record.get("bidNumber")

                    if not notice_id or not bid_id:
                        logging.warning(f"Skipping record due to missing ID: {record}")
                        continue

                    logging.info(f"Processing awarded bid: {notice_id} - {record.get('title')}")

                    try:
                        detail_url = f"https://webprocure.proactiscloud.com/wp-full-text-search/soldetail/{bid_id}"
                        detail_params = {"customerid": CUSTOMER_ID}
                        detail_response = session.get(detail_url, params=detail_params, timeout=60)
                        detail_response.raise_for_status()
                        detail_data_wrapper = detail_response.json()

                        if not detail_data_wrapper.get("records"):
                            logging.warning(f"No detail record found for bid {notice_id}")
                            continue

                        detail_data = detail_data_wrapper["records"][0]

                        
                        download_links = []
                        for doc in detail_data.get("bidDocs", []):
                            doc_details = doc.get("docAssoc", {}).get("docDoc", {})
                            file_id = doc_details.get("docid")
                            file_name = doc_details.get("name")
                            mime_type = doc_details.get("mimeType")

                            if all([file_id, file_name, mime_type]):
                                encoded_file_name = quote(file_name)
                                final_download_url = (
                                    f"{download_base_url}?docid={file_id}&eboid={CUSTOMER_ID}"
                                    f"&mimeType={mime_type}&docName={encoded_file_name}"
                                    f"&docUniqueName={encoded_file_name}&bidid={bid_id}"
                                )
                                download_links.append(f"{file_name}: {final_download_url}")

                        
                        issuer, email = None, None
                        contacts = detail_data.get("bidContacts", [])
                        if contacts:
                            contact_detail = contacts[0].get("bidContactDetail", {})
                            contact_string = contact_detail.get("contactinfo", "")
                            parts = [p.strip() for p in contact_string.split("\r\n") if p.strip()]
                            if parts:
                                issuer = parts[0]
                                email = next((p for p in parts if "@" in p), None)

                        
                        bid_title = detail_data.get("title", record.get("title"))
                        industry, sub_sector = classify_opportunity(bid_title)
                        closing_date_dt = parse_date_from_api(record.get("openDate"))

                        rfp_data = {
                            "notice_id": notice_id,
                            "title": bid_title,
                            "source": SOURCE_NAME,
                            "publish_date": parse_date_from_api(detail_data.get("cdate")),
                            "closing_date": closing_date_dt,
                            "issuer": issuer,
                            "email": email,
                            "page_url": (
                                f"https://webprocure.proactiscloud.com/wp-web-public/en/#/solicitation/{bid_id}"
                            ),
                            "industry": industry,
                            "type": sub_sector,
                            "description": detail_data.get("description", record.get("description")),
                            "download_links": "\n".join(download_links) if download_links else None,
                        }

                        cleaned_data = {k: v for k, v in rfp_data.items() if v is not None}
                        scraped_data.append(cleaned_data)
                        logging.info(f"Successfully processed and staged for export: {notice_id}")
                        time.sleep(1)

                    except requests.RequestException as e:
                        logging.error(f"Failed to process bid {notice_id}: {e}")
                    except json.JSONDecodeError as e:
                        logging.error(f"Failed to decode JSON for bid {notice_id}: {e}")

                current_offset += len(records_on_page)
                if current_offset >= total_hits:
                    logging.info("Processed all available records.")
                    break

                time.sleep(0.5)

    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}", exc_info=True)
    finally:
        if scraped_data:
            output_filename = "connecticut_awarded_bids.xlsx"
            df = pd.DataFrame(scraped_data)

            
            df["publish_date"] = pd.to_datetime(df["publish_date"], errors="coerce")
            df["closing_date"] = pd.to_datetime(df["closing_date"], errors="coerce")

           
            current_year = datetime.now().year
            df = df[df["publish_date"].dt.year == current_year]

            if not df.empty:
                
                df = df.sort_values(by="closing_date", ascending=False, na_position="last")

              l
                output_filename = f"connecticut_awarded_bids_{current_year}.xlsx"
                logging.info(f"Saving {len(df)} awarded bids from {current_year} to {output_filename}...")
                df.to_excel(output_filename, index=False, engine="openpyxl")
                logging.info(f"Successfully saved data to {output_filename}")
            else:
                logging.warning(f"No awarded bids found for {current_year}.")
        else:
            logging.warning("No data was scraped, so no Excel file was created.")


if __name__ == "__main__":
    scrape_connecticut_awarded()
