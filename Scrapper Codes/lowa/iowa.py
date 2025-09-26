import time
import re
import argparse
from datetime import datetime
from typing import Optional, List, Dict

import requests
import pandas as pd
from bs4 import BeautifulSoup

RETRIES = 3
BASE = "https://bidopportunities.iowa.gov"
LIST_API = f"{BASE}/Home/DT_HostedBidsSearch?agencyId=&enteredSearchText="

def parse_date_string(date_str: str):
    """Parse formats like '5/31/2025' or '05/31/2025' to ISO date string."""
    if not date_str:
        return None
    date_str = date_str.strip()
    try:
        return datetime.strptime(date_str, "%m/%d/%Y").date().isoformat()
    except ValueError:
        return None

def parse_time_string(time_str: str):
    """Parse '2:00:00 P.M.', '02:00 PM', or '5:01:00 PM' to 'HH:MM:SS'."""
    if not time_str:
        return None
    normalized = re.sub(r"\.", "", time_str).strip().upper()
    for fmt in ("%I:%M:%S %p", "%I:%M %p"):
        try:
            return datetime.strptime(normalized, fmt).time().strftime("%H:%M:%S")
        except ValueError:
            continue
    return None

def extract_panel_pairs(panel_body) -> Dict[str, Optional[str]]:
    """Extract label/value pairs from a 'panel-body' section."""
    out = {}
    if not panel_body:
        return out
    rows = panel_body.find_all("div", class_="row")
    for row in rows:
        label_elem = row.find("label")
        val_div = row.find("div", class_="col-md-9")
        if not label_elem or not val_div:
            continue
        key = (label_elem.get("for") or label_elem.get_text(strip=True)) or ""
        val = val_div.get_text(strip=True) or None
        if key:
            out[key] = val
    return out

def scrape_detail_page(url: str, headers: dict) -> Dict[str, Optional[str]]:
    """Scrape a single bid detail page and return a flat dict of fields."""
    resp = requests.get(url, headers=headers, timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    # Find all panels by headings
    panels = soup.find_all("div", class_="panel-group")

    general_info = {}
    agency_info = {}
    contact_info = {}
    description_info = {}
    valid_dates_info = {}
    attachments: List[Dict[str, str]] = []

    for panel in panels:
        try:
            heading = panel.find("div", class_="panel-heading")
            h3 = heading.find("h3") if heading else None
            title = h3.get_text(strip=True) if h3 else ""
            body = panel.find("div", class_="panel-body")

            if "Bid Information" in title:
                general_info = extract_panel_pairs(body)
            elif "Agency Information" in title:
                agency_info = extract_panel_pairs(body)
            elif "Contact Information" in title:
                contact_info = extract_panel_pairs(body)
            elif "Description" in title:
                description_info = extract_panel_pairs(body)
            elif "Valid Dates" in title:
                valid_dates_info = extract_panel_pairs(body)
            elif "Documents/Attachments" in title:
                # Rows typically: [file link] [download icon] [date]
                rows = body.find_all("div", class_="row") if body else []
                for r in rows:
                    try:
                        name_a = r.find("div", class_="col-md-8").find("a")
                        filename = name_a.get_text(strip=True)
                        view_url = name_a.get("href")
                        if view_url and view_url.startswith("/"):
                            view_url = f"{BASE}{view_url}"

                        dl_a = r.find("div", class_="col-md-1").find("a", class_="glyphicon-download")
                        download_url = dl_a.get("href") if dl_a else None
                        if download_url and download_url.startswith("/"):
                            download_url = f"{BASE}{download_url}"

                        url_final = download_url or view_url
                        if filename and url_final:
                            attachments.append({"name": filename, "url": url_final})
                    except Exception:
                        continue
        except Exception:
            continue

    # Parse valid dates: "From" and "Until" may include time after a space
    def split_dt(raw: Optional[str]):
        if not raw:
            return None, None
        try:
            date_part, time_part = raw.split(" ", 1)
            return parse_date_string(date_part), parse_time_string(time_part)
        except Exception:
            return parse_date_string(raw), None

    pub_date, pub_time = split_dt(valid_dates_info.get("From"))
    due_date, due_time = split_dt(valid_dates_info.get("Until"))

    # Build output row
    addr_parts = []
    if agency_info.get("AgencyAddress1"):
        addr_parts.append(agency_info.get("AgencyAddress1"))
    if agency_info.get("AgencyAddress2"):
        addr_parts.append(agency_info.get("AgencyAddress2"))
    addr_citystzip = agency_info.get("AgencyCityStateZip")
    full_addr = ", ".join(addr_parts)
    if full_addr and addr_citystzip:
        full_addr = f"{full_addr}, {addr_citystzip}"
    elif addr_citystzip:
        full_addr = addr_citystzip or None
    else:
        full_addr = full_addr or None

    row = {
        "source": "Iowa",
        "page_url": url,
        "bid_number": general_info.get("BidNumber"),
        "title": general_info.get("Solicitation"),
        "issuer": agency_info.get("AgencyName"),
        "place_of_performance": full_addr,
        "contact_email": contact_info.get("ContactEmail"),
        "contact_phone": contact_info.get("ContactPhoneNumber"),
        "description": description_info.get("Description"),
        "publish_date": pub_date,
        "publish_time": pub_time,
        "proposal_deadline": due_date,
        "deadline_time": due_time,
        "attachments": "; ".join([a["url"] for a in attachments]) if attachments else None,
    }
    return row

def scrape_iowa(max_rows: Optional[int] = None) -> List[Dict[str, Optional[str]]]:
    print("Starting Iowa bidding opportunities scraping...")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "X-Requested-With": "XMLHttpRequest",
    }

    # 1) Fetch bid list
    data = None
    for attempt in range(1, RETRIES + 1):
        try:
            resp = requests.get(LIST_API, headers=headers, timeout=20)
            resp.raise_for_status()
            data = resp.json()
            break
        except requests.RequestException as e:
            print(f"[Attempt {attempt}/{RETRIES}] List fetch failed: {e}")
            time.sleep(2)

    if not data:
        print("No data returned from list API.")
        return []

    detail_urls: List[str] = []
    for item in data.get("aaData", []):
        bid_uuid = (item.get("ID") or "").strip()
        if bid_uuid:
            detail_urls.append(f"{BASE}/Home/BidInfo?bidId={bid_uuid}")

    print(f"Collected {len(detail_urls)} bid detail URLs.")
    if max_rows is not None:
        detail_urls = detail_urls[:max_rows]

    # 2) Visit each detail page
    out_rows: List[Dict[str, Optional[str]]] = []
    for idx, u in enumerate(detail_urls, start=1):
        print(f"[{idx}/{len(detail_urls)}] {u}")
        row = None
        for attempt in range(1, RETRIES + 1):
            try:
                row = scrape_detail_page(u, headers)
                break
            except requests.RequestException as e:
                print(f"  [Attempt {attempt}/{RETRIES}] detail fetch failed: {e}")
                time.sleep(1.5)
            except Exception as e:
                print(f"  Unexpected error: {e}")
                break
        if row:
            out_rows.append(row)

    return out_rows

def save_to_excel(rows: List[Dict[str, Optional[str]]]) -> str:
    if not rows:
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        fname = f"iowa_bids_{ts}.xlsx"
        # write empty frame with known columns
        cols = [
            "source","page_url","bid_number","title","issuer","place_of_performance",
            "contact_email","contact_phone","description",
            "publish_date","publish_time","proposal_deadline","deadline_time","attachments"
        ]
        pd.DataFrame(columns=cols).to_excel(fname, index=False)
        return fname

    df = pd.DataFrame(rows)
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    fname = f"iowa_bids_{ts}.xlsx"
    df.to_excel(fname, index=False)
    return fname

def main():
    parser = argparse.ArgumentParser(description="Scrape Iowa Bid Opportunities to Excel (no DB).")
    parser.add_argument("--max-rows", type=int, default=None,
                        help="Limit how many bid detail pages to scrape.")
    args = parser.parse_args()

    rows = scrape_iowa(max_rows=args.max_rows)
    out = save_to_excel(rows)
    print(f"\nExcel written to: {out} (rows: {len(rows)})")

if __name__ == "__main__":
    main()