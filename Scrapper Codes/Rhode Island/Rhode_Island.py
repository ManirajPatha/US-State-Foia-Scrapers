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



def scrape_rhode_island_awarded():
    """
    Scrapes AWARDED bidding opportunities from Rhode Island's OSP Bid Board (WebProcure)
    and saves them to an Excel file.
    """
    CUSTOMER_ID = "46"
    SOURCE_NAME = "Rhode Island"
    scraped_data: List[Dict[str, Any]] = []

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json, text/plain, */*",
        "Referer": "https://webprocure.proactiscloud.com/wp-web-public/en/",
        "Origin": "https://webprocure.proactiscloud.com",
    }

    list_url = "https://webprocure.proactiscloud.com/wp-full-text-search/search/sols"
    list_params = {
        "customerid": CUSTOMER_ID,
        "q": "*",
        "sort": "r",
        "f": "ps=Awarded",
        "oids": "",
    }

    detail_url_tmpl = "https://webprocure.proactiscloud.com/wp-full-text-search/soldetail/{bidid}"
    download_base_url = "https://webprocure.proactiscloud.com/main/sol/viewdoc.do"

    try:
        with requests.Session() as session:
            session.headers.update(headers)
            current_offset = 0
            total_hits = 0
            is_first_page = True

            while True:
                logging.info(f"Fetching awarded results at offset {current_offset} ...")
                list_params["from"] = current_offset

                try:
                    resp = session.get(list_url, params=list_params, timeout=60)
                    resp.raise_for_status()
                    data = resp.json()
                except requests.RequestException as e:
                    logging.error(f"Failed to fetch page at offset {current_offset}: {e}. Stopping.")
                    break
                except json.JSONDecodeError as e:
                    logging.error(f"Failed to parse listing JSON at offset {current_offset}: {e}. Stopping.")
                    break

                if is_first_page:
                    total_hits = data.get("hits", 0)
                    if total_hits == 0:
                        logging.info("No awarded solicitations found.")
                        break
                    logging.info(f"Found {total_hits} awarded solicitations. Starting scrape ...")
                    is_first_page = False

                records = data.get("records", [])
                if not records:
                    logging.info("Finished processing all pages.")
                    break

                for record in records:
                    bid_id = record.get("bidid")
                    notice_id = record.get("bidNumber")
                    if not notice_id or not bid_id:
                        logging.warning(f"Skipping record due to missing ID: {record}")
                        continue

                    logging.info(f"Processing: {notice_id} - {record.get('title')}")

                    try:
                        detail_url = detail_url_tmpl.format(bidid=bid_id)
                        dparams = {"customerid": CUSTOMER_ID}
                        dresp = session.get(detail_url, params=dparams, timeout=60)
                        dresp.raise_for_status()
                        dwrapper = dresp.json()

                        if not dwrapper.get("records"):
                            logging.warning(f"No detail found for {notice_id}")
                            continue

                        drec = dwrapper["records"][0]

                        download_links = []
                        for doc in drec.get("bidDocs", []):
                            doc_details = doc.get("docAssoc", {}).get("docDoc", {})
                            file_id = doc_details.get("docid")
                            file_name = doc_details.get("name")
                            mime_type = doc_details.get("mimeType")
                            if all([file_id, file_name, mime_type]):
                                encoded_file_name = quote(file_name)
                                final_url = (
                                    f"{download_base_url}?docid={file_id}&eboid={CUSTOMER_ID}"
                                    f"&mimeType={mime_type}&docName={encoded_file_name}"
                                    f"&docUniqueName={encoded_file_name}&bidid={bid_id}"
                                )
                                download_links.append(f"{file_name}: {final_url}")

                        issuer, email = None, None
                        contacts = drec.get("bidContacts", [])
                        if contacts:
                            cdetail = contacts[0].get("bidContactDetail", {})
                            contact_string = cdetail.get("contactinfo", "")
                            parts = [p.strip() for p in contact_string.split("\r\n") if p.strip()]
                            if parts:
                                issuer = parts[0]
                                email = next((p for p in parts if "@" in p), None)

                        bid_title = drec.get("title", record.get("title"))
                        industry, sub_sector = classify_opportunity(bid_title)
                        closing_date = parse_date_from_api(record.get("openDate"))

                        rfp_data = {
                            "notice_id": notice_id,
                            "title": bid_title,
                            "source": SOURCE_NAME,
                            "publish_date": parse_date_from_api(drec.get("cdate")),
                            "closing_date": closing_date,
                            "issuer": issuer,
                            "email": email,
                            "page_url": f"https://webprocure.proactiscloud.com/wp-web-public/en/#/bidboard/bid/{bid_id}?customerid={CUSTOMER_ID}",
                            "industry": industry,
                            "type": sub_sector,
                            "description": drec.get("description", record.get("description")),
                            "download_links": "\n".join(download_links) if download_links else None,
                        }

                        cleaned = {k: v for k, v in rfp_data.items() if v is not None}
                        scraped_data.append(cleaned)
                        logging.info(f"Staged: {notice_id}")
                        time.sleep(0.5)

                    except requests.RequestException as e:
                        logging.error(f"Detail fetch failed for {notice_id}: {e}")
                    except json.JSONDecodeError as e:
                        logging.error(f"Detail JSON decode failed for {notice_id}: {e}")

                current_offset += len(records)
                if current_offset >= total_hits:
                    logging.info("Processed all available records.")
                    break

                time.sleep(0.25)

    except Exception as e:
        logging.error(f"Unexpected error: {e}", exc_info=True)
    finally:
        if scraped_data:
            output_filename = "rhode_island_awarded_bids.xlsx"
            logging.info(f"Saving {len(scraped_data)} awarded bids to {output_filename} ...")
            df = pd.DataFrame(scraped_data)
            df.to_excel(output_filename, index=False, engine="openpyxl")
            logging.info(f"Saved to {output_filename}")
        else:
            logging.warning("No data scraped; no Excel file created.")


if __name__ == "__main__":
    scrape_rhode_island_awarded()