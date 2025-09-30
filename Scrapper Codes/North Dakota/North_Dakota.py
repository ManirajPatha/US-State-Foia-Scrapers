import requests
from bs4 import BeautifulSoup
import re
import os
from urllib.parse import parse_qs, urlparse
import time
import logging
from datetime import datetime
from dateutil.relativedelta import relativedelta
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import tempfile
import shutil
import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

HEADERS = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36"
}

DOWNLOAD_DIR = tempfile.mkdtemp(prefix="nd_downloads_")
LISTING_URL = "https://apps.nd.gov/csd/spo/services/bidder/searchSolicitation.do"
ATTACHMENT_URL = "https://apps.nd.gov/csd/spo/services/bidder/streamFileServlet"


def create_session_with_retry() -> requests.Session:
    session = requests.Session()
    retry_strategy = Retry(total=3, status_forcelist=[429, 500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session

def make_request(session: requests.Session, method: str, url: str, **kwargs):
    try:
        response = session.request(method, url, **kwargs)
        response.raise_for_status()
        return response
    except requests.exceptions.RequestException as e:
        logger.error(f"Request failed for {url}: {e}")
        return None

def parse_date(date_str):
    if not date_str:
        return None
    try:
        return datetime.strptime(date_str, "%m/%d/%Y").strftime("%Y-%m-%d")
    except:
        return None

def parse_time(time_str):
    if not time_str:
        return None
    try:
        return datetime.strptime(time_str, "%I:%M %p").strftime("%H:%M:%S")
    except:
        return None

def split_close(close_str):
    if not close_str:
        return close_str, None
    match = re.match(r"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2}\s+[AP]M)", close_str)
    return match.groups() if match else (close_str, None)

def get_filename(response):
    cd = response.headers.get('Content-Disposition', '')
    if cd:
        match = re.search(r'filename="([^"]+)"', cd)
        if not match:
            match = re.search(r'filename=([^;]+)', cd)
        if match:
            from urllib.parse import unquote_plus
            return re.sub(r'[^\w\s.-]', '', unquote_plus(match.group(1)).replace(' ', '_'))
    return None


def get_session(session, url):
    response = make_request(session, "GET", url, headers=HEADERS)
    return response is not None

def get_listing_page(session):
    startdate = datetime.now() - relativedelta(months=6)
    stopdate = datetime.now()
    params = {
        "path": "/bidder/searchSolicitation",
        "command": "searchSearchSolicitation",
        "searchDT.startDate": startdate.strftime("%m/%d/%Y"),
        "searchDT.stopDate": stopdate.strftime("%m/%d/%Y")
    }
    response = make_request(session, "GET", LISTING_URL, params=params, headers=HEADERS)
    return response

def extract_solicitation_rows(soup):
    table = soup.find('table', {'summary': 'Results from Project Search'})
    return table.find_all('tr') if table else []

def extract_solicitation_info(row):
    cols = row.find_all('td')
    if len(cols) != 5:
        return None
    link = cols[1].find('a')
    if not link or 'href' not in link.attrs:
        return None
    match = re.search(r"submitViewDetails\('(.*?)'\)", link['href'])
    if not match:
        return None
    return match.group(1), cols[1].text.strip()

def get_detail_page(session, solicitation_id):
    params = {
        "path": "/bidder/searchSolicitation",
        "command": "viewSearchSolicitationDetails",
        "selectedSolicitation": solicitation_id
    }
    return make_request(session, "GET", LISTING_URL, params=params, headers=HEADERS)

def extract_solicitation_fields(soup):
    fields = {}
    tables = soup.find_all('table', {'summary': ' '})
    if tables:
        for tr in tables[0].find_all('tr'):
            th, td = tr.find('th'), tr.find('td')
            if th and td:
                fields[th.text.strip().rstrip(':')] = td.text.strip()
    return fields

def extract_attachments(session, soup, detail_url):
    downloads = []
    tables = soup.find_all('table')
    for table in tables:
        rows = table.find_all('tr')[1:]
        for tr in rows:
            cols = tr.find_all('td')
            if len(cols) < 4:
                continue
            link = cols[3].find('a')
            if not link or 'href' not in link.attrs:
                continue
            href = link['href']
            parsed = urlparse(href)
            attachment_id = parse_qs(parsed.query).get('selectedAttachmentId', [''])[0]
            if attachment_id:
                resp = make_request(session, "GET", ATTACHMENT_URL, params={"selectedAttachmentId": attachment_id}, headers=HEADERS)
                if not resp:
                    continue
                filename = get_filename(resp) or cols[0].text.strip()
                file_path = os.path.join(DOWNLOAD_DIR, filename)
                with open(file_path, 'wb') as f:
                    f.write(resp.content)
                downloads.append({"name": filename, "url": href})
    return downloads

def process_solicitation(session, solicitation_id, sol_number):
    detail_resp = get_detail_page(session, solicitation_id)
    if not detail_resp:
        return None
    soup = BeautifulSoup(detail_resp.text, 'html.parser')
    fields = extract_solicitation_fields(soup)
    attachments = extract_attachments(session, soup, detail_resp.url)
    
    close_full = fields.get('Closes', '')
    close_date_raw, close_time_raw = split_close(close_full)
    close_date = parse_date(close_date_raw)
    
    if close_date:
        close_dt = datetime.strptime(close_date, "%Y-%m-%d")
        if close_dt > datetime.now():
            return None
    
    return {
        "Notice ID": sol_number,
        "Title": fields.get("Title"),
        "Issued Date": parse_date(fields.get("Issued")),
        "Close Date": close_date,
        "Close Time": parse_time(close_time_raw),
        "Agency": fields.get("Issuing Agency"),
        "Detail URL": detail_resp.url,
        "Attachments": ", ".join([att['url'] for att in attachments])
    }

def scrape_north_dakota(url):
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    session = create_session_with_retry()
    if not get_session(session, url):
        logger.error("Failed to establish session.")
        return
    listing_resp = get_listing_page(session)
    if not listing_resp:
        logger.error("Failed to fetch listing page.")
        return
    soup = BeautifulSoup(listing_resp.text, 'html.parser')
    rows = extract_solicitation_rows(soup)

    all_data = []
    for row in rows:
        info = extract_solicitation_info(row)
        if not info:
            continue
        solicitation_id, sol_number = info
        data = process_solicitation(session, solicitation_id, sol_number)
        if data:
            all_data.append(data)

    if all_data:
        df = pd.DataFrame(all_data)
        excel_file = "north_dakota_closed_rfps.xlsx"
        df.to_excel(excel_file, index=False)
        logger.info(f"Saved {len(all_data)} closed solicitations to {excel_file}")
    else:
        logger.info("No closed solicitations found.")

    session.close()
    shutil.rmtree(DOWNLOAD_DIR, ignore_errors=True)

if __name__ == "__main__":
    scrape_north_dakota("https://apps.nd.gov/csd/spo/services/bidder/searchSolicitation.do")
