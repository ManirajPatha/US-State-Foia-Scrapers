# maine_rfp_awarded_archives_to_excel.py
# Scrapes Awarded opportunities from:
# https://www.maine.gov/dafs/bbm/procurementservices/vendors/rfps/rfp-archives
#
# Output: Excel (.xlsx) with columns:
# Title, Title URL, RFP #, Issuing Department, Date Posted, Q&A/Amendments (JSON),
# Proposal Due Date, RFP Status, Awarded Vendor(s) (JSON), Next Anticipated RFP Release, Source Page

import argparse
import json
import os
from datetime import datetime
from typing import List, Dict
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

DEFAULT_URL = "https://www.maine.gov/dafs/bbm/procurementservices/vendors/rfps/rfp-archives"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    )
}

def get_soup(url: str) -> BeautifulSoup:
    resp = requests.get(url, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    return BeautifulSoup(resp.text, "html.parser")

def discover_all_pages(start_url: str) -> List[str]:
    """
    Collects all pagination pages from the archive.
    Works whether the page has explicit pagination links or is a single page.
    Also follows the 'Older Archives' / '2024 to present' links found at top.
    """
    to_visit = [start_url]
    seen = set()

    pages = []

    while to_visit:
        url = to_visit.pop(0)
        if url in seen:
            continue
        seen.add(url)

        soup = get_soup(url)
        pages.append(url)

        # 1) Follow explicit pagination links if present
        # (look for typical pager containers)
        for sel in ["ul.pagination a", ".pager a", "nav[aria-label*='pagination'] a"]:
            for a in soup.select(sel):
                href = a.get("href")
                if not href:
                    continue
                nxt = urljoin(url, href)
                if nxt not in seen and nxt not in to_visit:
                    to_visit.append(nxt)

        # 2) Follow the "Older Archives" / "2024 to present" quick links if present
        #    We add them once; subsequent pages’ discovery handles their own pagination.
        top_links = soup.select("a")
        for a in top_links:
            text = (a.get_text(strip=True) or "").lower()
            if "older archives" in text or "2023 and prior" in text or "2024 to present" in text:
                href = a.get("href")
                if href:
                    nxt = urljoin(url, href)
                    if nxt not in seen and nxt not in to_visit:
                        to_visit.append(nxt)

    return pages

def extract_links(cell, base_url: str) -> List[Dict[str, str]]:
    out = []
    if not cell:
        return out
    for a in cell.find_all("a"):
        text = a.get_text(strip=True) or ""
        href = a.get("href")
        if not href:
            continue
        out.append({"text": text, "url": urljoin(base_url, href)})
    return out

def parse_table_on_page(page_url: str) -> List[Dict[str, str]]:
    soup = get_soup(page_url)

    # Find the archive table (there’s usually just one primary table on the page)
    table = soup.find("table")
    if not table:
        return []

    tbody = table.find("tbody") or table
    rows = tbody.find_all("tr")
    results = []

    for tr in rows:
        tds = tr.find_all("td")
        if len(tds) < 9:
            # Unexpected layout; skip
            continue

        # Columns by visual order in the snippet:
        # 0 Title (linked to RFP doc/page)
        # 1 RFP #
        # 2 Issuing Department
        # 3 Date Posted
        # 4 Q and A Summary and Amendment (links)
        # 5 Proposal Due Date
        # 6 RFP Status
        # 7 Awarded Vendor(s)
        # 8 Next Anticipated RFP Release

        # Title + URL
        title_a = tds[0].find("a")
        title_text = title_a.get_text(strip=True) if title_a else tds[0].get_text(strip=True)
        title_url = urljoin(page_url, title_a["href"]) if (title_a and title_a.get("href")) else ""

        rfp_no = tds[1].get_text(strip=True)
        issuing_dept = tds[2].get_text(strip=True)
        date_posted = tds[3].get_text(strip=True)

        qa_amend_links = extract_links(tds[4], page_url)

        proposal_due = tds[5].get_text(strip=True)
        rfp_status = tds[6].get_text(strip=True)

        # Only keep "Awarded"
        if rfp_status.strip().lower() != "awarded":
            continue

        # Awarded Vendors (may contain one or more links/text)
        vendors_links = extract_links(tds[7], page_url)
        # If there were no anchors, capture raw text
        vendors_text = tds[7].get_text(strip=True)
        if not vendors_links and vendors_text:
            vendors_links = [{"text": vendors_text, "url": ""}]

        next_anticipated = tds[8].get_text(strip=True)

        results.append({
            "Title": title_text,
            "Title URL": title_url,
            "RFP #": rfp_no,
            "Issuing Department": issuing_dept,
            "Date Posted": date_posted,
            "Q&A/Amendments (JSON)": json.dumps(qa_amend_links, ensure_ascii=False),
            "Proposal Due Date": proposal_due,
            "RFP Status": rfp_status,
            "Awarded Vendor(s) (JSON)": json.dumps(vendors_links, ensure_ascii=False),
            "Next Anticipated RFP Release": next_anticipated,
            "Source Page": page_url
        })

    return results

def scrape_awarded_archives(start_url: str) -> pd.DataFrame:
    pages = discover_all_pages(start_url)
    all_rows: List[Dict[str, str]] = []
    for i, page in enumerate(pages, 1):
        print(f"[INFO] Parsing page {i}/{len(pages)}: {page}")
        all_rows.extend(parse_table_on_page(page))

    # Deduplicate by (RFP #, Title) if needed
    seen = set()
    deduped = []
    for r in all_rows:
        key = (r.get("RFP #", ""), r.get("Title", ""))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(r)

    cols = [
        "Title",
        "Title URL",
        "RFP #",
        "Issuing Department",
        "Date Posted",
        "Q&A/Amendments (JSON)",
        "Proposal Due Date",
        "RFP Status",
        "Awarded Vendor(s) (JSON)",
        "Next Anticipated RFP Release",
        "Source Page",
    ]
    df = pd.DataFrame(deduped, columns=cols)
    return df

def main():
    parser = argparse.ArgumentParser(
        description="Maine RFP Archives (Awarded only) → Excel"
    )
    parser.add_argument(
        "--url",
        default=DEFAULT_URL,
        help=f"Archives URL (default: {DEFAULT_URL})",
    )
    parser.add_argument(
        "--out",
        default=f"maine_rfp_awarded_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        help="Output Excel filepath (.xlsx).",
    )
    args = parser.parse_args()

    df = scrape_awarded_archives(args.url)

    # Ensure directory exists
    out_path = os.path.abspath(args.out)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    df.to_excel(out_path, index=False)
    print(f"[INFO] Wrote {len(df)} awarded row(s) → {out_path}")

if __name__ == "__main__":
    main()