import logging
import re
import time
import os
import json
import zipfile
import tempfile
from datetime import datetime
from typing import Optional, Dict, List

import requests
from bs4 import BeautifulSoup
from urllib3.util.retry import Retry
from requests.adapters import HTTPAdapter

# --- Set Aside Map ---
SET_ASIDE_MAP = {
    "8A": "8(a) Competed",
    "HS3": "8(a) with HUBZone Preference",
    "8AN": "8(a) Sole Source",
    "HMP": "HBCU or MI Set Aside - Partial",
    "HMT": "HBCU or MI Set Aside - Total",
    "HZC": "HUBZone Set Aside",
    "HZS": "HUBZone Sole Source",
    "BI": "Buy Indian",
    "IEE": "Indian Economic Enterprise",
    "ISBEE": "Indian Small Business Economic Enterprise",
    "NONE": "No Set Aside Used",
    "ESB": "Emerging Small Business Set Aside",
    "RSB": "Reserved for Small Business",
    "SBP": "Small Business Set Aside - Partial",
    "SBA": "Small Business Set Aside - Total",
    "VSB": "Very Small Business",
    "SDVOSBC": "Service-Disabled Veteran-Owned Small Business Set Aside",
    "SDVOSBS": "SDVOSB Sole Source",
    "VSA": "Veteran Set Aside",
    "VSS": "Veteran Sole Source",
    "EDWOSBSS": "Economically Disadvantaged Women Owned Small Business Sole Source",
    "EDWOSB": "Economically Disadvantaged Women Owned Small Business",
    "WOSBSS": "Women-Owned Small Business Sole Source",
    "WOSB": "Women-Owned Small Business",
}


def resolve_set_aside(code: Optional[str]) -> Optional[str]:
    if not code:
        return None
    return SET_ASIDE_MAP.get(code.upper(), code)


# --- Configuration ---
LIST_API_URL = "https://sam.gov/api/prod/sgs/v1/search/"
DETAIL_API_BASE_URL = "https://sam.gov/api/prod/opps/v2/opportunities/"
RESOURCES_API_BASE_URL = "https://sam.gov/api/prod/opps/v3/opportunities/"

DOWNLOAD_DIR = "sam_gov_downloads"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

JSON_OUTPUT_FILE = "sam_gov_data_inactive_f1.json"
REQUEST_TIMEOUT = 45
RECORDS_PER_PAGE = 25

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/5.0 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    ),
    "Content-Type": "application/json",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


def create_session() -> requests.Session:
    """
    Create a requests session with retries configured.
    """
    session = requests.Session()
    session.headers.update(HEADERS)

    retries = Retry(
        total=3,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=frozenset(["GET", "POST", "PUT", "DELETE", "HEAD", "OPTIONS"]),
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


# --- Helper: Generate year chunks for SBA ---
def generate_year_chunks(start_year: int = 2005, end_year: Optional[int] = None, tz_offset: str = "+05:30") -> List[tuple]:
    """
    Generate (modified_date.from, modified_date.to) pairs for each year.
    Uses a timezone offset suffix such as '+05:30' (Asia/Kolkata) by default.

    Example returns:
    [
      ("2005-01-01+05:30", "2005-12-31+23:59:59+05:30"),
      ("2006-01-01+05:30", "2006-12-31+23:59:59+05:30"),
      ...
    ]
    """
    if end_year is None:
        end_year = datetime.now().year

    chunks = []
    for y in range(start_year, end_year + 1):
        date_from = f"{y}-01-01{tz_offset}"
        date_to = f"{y}-12-31T23:59:59{tz_offset}"
        # Keep formats compatible with examples (API tolerates both with/without T),
        # but include time for `to` to capture full day.
        chunks.append((date_from, date_to))
    return chunks


# --- Fetch attachment metadata ---
def fetch_attachment_metadata(session: requests.Session, internal_id: str) -> List[Dict]:
    try:
        resources_url = f"{RESOURCES_API_BASE_URL}{internal_id}/resources"
        api_params = {"api_key": "null", "random": int(time.time() * 1000)}
        response = session.get(resources_url, params=api_params, timeout=REQUEST_TIMEOUT)

        if response.status_code != 200:
            return []

        resources = (
            response.json().get("_embedded", {}).get("opportunityAttachmentList", [])
        )
        if not resources:
            return []

        return [
            {"name": att.get("name"), "resourceId": att.get("resourceId")}
            for att in resources[0].get("attachments", [])
            if att
            and att.get("deletedFlag") == "0"
            and att.get("name")
            and att.get("resourceId")
        ]

    except Exception as e:
        logging.error(f"Error parsing attachment metadata for {internal_id}: {e}")
        return []


# --- Download and ZIP attachments ---
def download_and_zip_attachments(session: requests.Session, notice_id: str, attachment_metadata: List[Dict]) -> Optional[str]:
    if not attachment_metadata:
        return None

    zip_filename = f"{notice_id}.zip"
    zip_filepath = os.path.join(DOWNLOAD_DIR, zip_filename)

    if os.path.exists(zip_filepath):
        logging.info(f" -> Zip already exists: {zip_filename}")
        return zip_filepath

    with tempfile.TemporaryDirectory() as temp_dir:
        files_downloaded = False

        for attachment in attachment_metadata:
            redirect_url = (
                f"{RESOURCES_API_BASE_URL}resources/files/"
                f"{attachment['resourceId']}/download"
            )
            try:
                redirect_res = session.get(
                    redirect_url,
                    params={"api_key": "null", "token": ""},
                    allow_redirects=False,
                    timeout=REQUEST_TIMEOUT,
                )

                if (
                    redirect_res.status_code == 303
                    and "Location" in redirect_res.headers
                ):
                    s3_url = redirect_res.headers["Location"]
                    safe_filename = re.sub(
                        r'[\\/*?:"<>|]', "_", attachment["name"]
                    )
                    temp_filepath = os.path.join(temp_dir, safe_filename)

                    logging.info(f" -> Downloading: {safe_filename}...")

                    file_resp = session.get(
                        s3_url, stream=True,
                        headers={"User-Agent": "Mozilla/5.0"},
                        timeout=REQUEST_TIMEOUT,
                    )

                    if file_resp.status_code == 200:
                        with open(temp_filepath, "wb") as f:
                            for chunk in file_resp.iter_content(chunk_size=8192):
                                f.write(chunk)
                        files_downloaded = True
                    else:
                        logging.warning(f" -> Failed to download content for {safe_filename}")

                else:
                    logging.warning(f" -> No redirect found for {attachment['name']}")

            except Exception as e:
                logging.error(f" -> Error processing attachment {attachment['name']}: {e}")

        if files_downloaded:
            logging.info(f" -> Zipping files into {zip_filename}...")
            try:
                with zipfile.ZipFile(zip_filepath, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for root, _, files in os.walk(temp_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            zipf.write(file_path, arcname=file)
                return zip_filepath
            except Exception as e:
                logging.error(f" -> Error creating zip file: {e}")
                return None

        return None


# --- Process each opportunity ---
def process_opportunity(session: requests.Session, list_item: Dict) -> Optional[Dict]:
    internal_id = list_item.get("_id")
    if not internal_id:
        return None

    try:
        logging.info(f" - Processing Internal ID: {internal_id}")

        api_params = {"api_key": "null", "random": int(time.time() * 1000)}
        detail_response = session.get(
            f"{DETAIL_API_BASE_URL}{internal_id}",
            params=api_params,
            timeout=REQUEST_TIMEOUT,
        )
        detail_response.raise_for_status()

        details = detail_response.json()
        data2 = details.get("data2", {})

        solicitation_number = data2.get("solicitationNumber")
        notice_id = (
            solicitation_number.strip()
            if solicitation_number and solicitation_number.strip()
            else internal_id
        )

        # Attachments
        attachment_metadata = fetch_attachment_metadata(session, internal_id)
        zip_path = download_and_zip_attachments(
            session, notice_id, attachment_metadata
        ) if attachment_metadata else None

        # Text fields
        title = data2.get("title")
        description_html = (details.get("description") or [{}])[0].get("body", "")
        description_text = BeautifulSoup(description_html, "html.parser").get_text(separator=" ", strip=True)

        # Set aside
        set_aside_code = (data2.get("solicitation") or {}).get("setAside")
        set_aside = resolve_set_aside(set_aside_code)

        # Org hierarchy
        issuer, sub_tier, office = None, None, None
        org_hierarchy = list_item.get("organizationHierarchy")
        if org_hierarchy:
            for org in org_hierarchy:
                if org.get("level") == 1:
                    issuer = org.get("name")
                elif org.get("level") == 2:
                    sub_tier = org.get("name")
            office = org_hierarchy[-1].get("name")

        # Contacts
        all_contacts = data2.get("pointOfContact", []) or []
        primary_contact = next((p for p in all_contacts if p and p.get("type") == "primary"), {})
        poc_name = primary_contact.get("fullName")
        poc_email = primary_contact.get("email")
        phone_number = primary_contact.get("phone")
        combined_email = f"{poc_name}, {poc_email}" if poc_name and poc_email else poc_name or poc_email

        # Award details
        award = data2.get("award") or {}
        award_date = award.get("date")
        award_amount = award.get("amount")
        award_number = award.get("number")
        award_line_item_number = award.get("lineItemNumber")

        awardee = award.get("awardee") or {}
        awardee_location = awardee.get("location") or {}

        # Final structure
        rfp_data = {
            "notice_id": notice_id,
            "title": title,
            "publish_date": (details.get("postedDate") or "").split("T")[0],
            "proposal_deadline": (
                (data2.get("solicitation") or {})
                .get("deadlines", {})
                .get("response", "")
                .split("T")[0]
            ) or None,
            "source": "Federal",
            "issuer": issuer,
            "sub_tier": sub_tier,
            "office": office,
            "email": combined_email,
            "phone_number": phone_number,
            "naics_code": (data2.get("naics") or [{}])[0].get("code"),
            "description": description_text,
            "page_url": f"https://sam.gov/workspace/contract/opp/{internal_id}/view",
            "local_zip_path": zip_path,
            "place_of_performance": ", ".join(
                filter(
                    None,
                    [
                        (data2.get("placeOfPerformance") or {}).get("city", {}).get("name"),
                        (data2.get("placeOfPerformance") or {}).get("state", {}).get("name"),
                        (data2.get("placeOfPerformance") or {}).get("country", {}).get("code"),
                    ],
                )
            ) or "Not Specified",
            "notice_type": (list_item.get("type") or {}).get("value") or "Not Specified",
            "setaside": set_aside,
            "award_date": award_date,
            "award_amount": award_amount,
            "award_number": award_number,
            "award_line_item_number": award_line_item_number,
            "awardee_name": awardee.get("name"),
            "awardee_city": (awardee_location.get("city") or {}).get("name"),
            "awardee_state": (awardee_location.get("state") or {}).get("name"),
            "awardee_country": (awardee_location.get("country") or {}).get("name"),
            "awardee_zip": awardee_location.get("zip"),
        }

        logging.info(f"   Processed: {title[:50]}..." if title else "   Processed an item...")
        return rfp_data

    except Exception as e:
        logging.error(f"Unexpected error occurred while processing {internal_id}: {e}")
        return None


# MAIN SCRAPER WITH CHUNKING BY SET-ASIDE CODE
def scrape_sam_gov():
    logging.info("--- Running SAM.gov Scraper (Chunking by Set-Aside) ---")

    session = create_session()
    all_results = []

    # Chunking list
    target_set_asides = [
        "SBA", "SBP", "8A", "8AN", "HZC", "HZS",
        "SDVOSBC", "SDVOSBS", "WOSB", "WOSBSS",
        "EDWOSB", "EDWOSBSS", "IEE", "BI",
        "ISBEE", "VSA", "VSS",
    ]

    # Process each set-aside chunk individually
    for set_aside_code in target_set_asides:
        logging.info(f"\n========== Processing SET-ASIDE: {set_aside_code} ==========\n")

        # --- Special handling for SBA (10k+ records) ---
        if set_aside_code == "SBA":
            # Adjust start_year if you want to go further back or forward
            year_chunks = generate_year_chunks(start_year=2005, end_year=datetime.now().year, tz_offset="+05:30")

            for (date_from, date_to) in year_chunks:
                logging.info(f"[SBA] Year Chunk: {date_from[:4]}  →  {date_from} to {date_to}")

                page = 0
                total_pages = 1

                while True:
                    params = {
                        "random": int(time.time() * 1000),
                        "index": "opp",
                        "page": page,
                        "sort": "-modifiedDate",
                        "size": RECORDS_PER_PAGE,
                        "mode": "search",
                        "responseType": "json",
                        "is_active": "false",
                        "notice_type": "a",
                        "set_aside": "SBA",
                        "modified_date.from": date_from,
                        "modified_date.to": date_to,
                    }

                    try:
                        logging.info(f"[SBA] {date_from[:4]} → Fetching page {page + 1}...")

                        response = session.get(
                            LIST_API_URL, params=params, timeout=REQUEST_TIMEOUT
                        )
                        response.raise_for_status()

                        data = response.json()

                        if page == 0:
                            total_pages = data.get("page", {}).get("totalPages", 1)
                            total_records = data.get("page", {}).get("totalElements", 0)

                            logging.info(
                                f"[SBA] Year {date_from[:4]} → {total_records} records across {total_pages} pages."
                            )

                            if total_records >= 10000:
                                logging.warning(
                                    f"[SBA] WARNING: Year chunk {date_from[:4]} still exceeds 10k records!"
                                )

                        opportunities = data.get("_embedded", {}).get("results", [])

                        if not opportunities:
                            logging.info(f"[SBA] Year {date_from[:4]} → No more results for this year chunk.")
                            break

                        for item in opportunities:
                            result = process_opportunity(session, item)
                            if result:
                                all_results.append(result)

                        page += 1
                        if page >= total_pages:
                            logging.info(f"[SBA] Completed year {date_from[:4]} ({total_pages} pages).")
                            break

                        # polite pause (tune as needed)
                        time.sleep(1)

                    except Exception as e:
                        logging.error(f"[SBA] Error while fetching year {date_from[:4]}: {e}")
                        # break this year's loop and continue with next year
                        break

            # After finishing SBA chunks, continue to next set-aside
            continue

        # --- Normal pagination for other set-asides ---
        page = 0
        total_pages = 1

        while True:
            params = {
                "random": int(time.time() * 1000),
                "index": "opp",
                "page": page,
                "sort": "-modifiedDate",
                "size": RECORDS_PER_PAGE,
                "mode": "search",
                "responseType": "json",
                "is_active": "false",
                "notice_type": "a",
                "set_aside": set_aside_code,
            }

            try:
                logging.info(f"[{set_aside_code}] Fetching page {page + 1}...")

                response = session.get(
                    LIST_API_URL, params=params, timeout=REQUEST_TIMEOUT
                )
                response.raise_for_status()

                data = response.json()

                if page == 0:
                    total_pages = data.get("page", {}).get("totalPages", 1)
                    total_records = data.get("page", {}).get("totalElements", 0)

                    logging.info(
                        f"[{set_aside_code}] Found {total_records} results across {total_pages} pages."
                    )

                    if total_records >= 10000:
                        logging.warning(
                            f"[{set_aside_code}] WARNING: This set-aside still exceeds 10k records!"
                        )

                opportunities = data.get("_embedded", {}).get("results", [])

                if not opportunities:
                    logging.info(f"[{set_aside_code}] No more results.")
                    break

                for item in opportunities:
                    result = process_opportunity(session, item)
                    if result:
                        all_results.append(result)

                page += 1
                if page >= total_pages:
                    logging.info(
                        f"[{set_aside_code}] Completed all {total_pages} pages."
                    )
                    break

                time.sleep(1)

            except Exception as e:
                logging.error(f"[{set_aside_code}] Error: {e}")
                break

    # --- Final JSON Output ---
    logging.info(f"\nWriting {len(all_results)} total records to {JSON_OUTPUT_FILE}...")

    with open(JSON_OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(all_results, f, indent=4, default=str)

    logging.info("\n--- Scraper Finished Successfully ---")


if __name__ == "__main__":
    scrape_sam_gov()
