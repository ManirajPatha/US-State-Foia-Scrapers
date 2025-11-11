#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Nevada ePro — Closed solicitations scraper (stable pages + robust attachments + AUTOSAVE)

- Status=Closed
- Scrapes ALL pages; downloads each record's "File Attachments"
- Writes Excel with all rows across pages
- NEW: Autosave partial results after each page (and every 10 rows) and on any interruption.
"""

import argparse, os, re, sys, time, json, signal
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

ADV_URL = "https://nevadaepro.com/bso/view/search/external/advancedSearchBid.xhtml"

TARGET_HEADERS = [
    "Bid Solicitation #","Organization Name","Contract #","Buyer","Description",
    "Bid Opening Date","Bid Holder List","Awarded Vendor(s)","Status","Alternate Id",
]

FILE_PAT = re.compile(r"\.(pdf|docx?|xlsx?|txt|csv|zip)\b", re.I)
QUIET_SECONDS = 2
DOWNLOAD_TIMEOUT = 300
MIN_WAIT_AFTER_DOWNLOAD = 1.5
RETRY_STALE = 4
ROW_AUTOSAVE_INTERVAL = 10

stop_flag = {"stop": False}
def _sig_stop(signum, frame): stop_flag["stop"] = True
for _s in ("SIGINT","SIGTERM","SIGBREAK"):
    if hasattr(signal, _s):
        signal.signal(getattr(signal, _s), _sig_stop)

# ---------- small helpers ----------
def normalize(s: str) -> str:
    import re as _re
    return _re.sub(r"\s+", " ", (s or "").strip())

def ensure_dir(path: str) -> str:
    os.makedirs(path, exist_ok=True); return os.path.abspath(path)

def wait_for_downloads(dirpath: str, quiet=QUIET_SECONDS, timeout=DOWNLOAD_TIMEOUT):
    start=time.time(); last=start; prev=set(os.listdir(dirpath))
    while True:
        cur=set(os.listdir(dirpath))
        if cur!=prev: last=time.time(); prev=cur
        if not any(f.endswith((".crdownload",".tmp")) for f in cur) and time.time()-last>=quiet:
            return
        if time.time()-start>timeout: return
        time.sleep(0.4)

def timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M")

def save_progress(out_folder: str, results: List[dict], page_no: int, note: str, final: bool=False) -> str:
    cols = TARGET_HEADERS + ["Row URL","Attachment Files"]
    name = f"nevada_closed_with_attachments_{timestamp()}"
    name += "" if final else f"_partial_p{page_no}_r{len(results)}"
    xlsx = os.path.join(out_folder, f"{name}.xlsx")
    pd.DataFrame(results, columns=cols).to_excel(xlsx, index=False)
    # write a tiny JSON manifest too
    with open(os.path.join(out_folder,"progress.json"), "w", encoding="utf-8") as f:
        json.dump({
            "rows_collected": len(results),
            "last_page_completed": page_no,
            "note": note,
            "excel": os.path.basename(xlsx),
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }, f, indent=2)
    return xlsx

# ---------- driver ----------
def build_driver(headless: bool, download_dir: str) -> webdriver.Chrome:
    opts = Options()
    if headless: opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1600,1100")
    opts.add_argument("--disable-gpu"); opts.add_argument("--no-sandbox"); opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    prefs = {
        "download.default_directory": download_dir.replace("/", "\\"),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }
    opts.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opts)

# ---------- search ----------
def wait_ready(drv, timeout=30):
    WebDriverWait(drv, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR,"input[name='javax.faces.ViewState']"))
    )

def set_status_closed(drv, timeout=20):
    sel = WebDriverWait(drv, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR,"select[name$=':status']"))
    )
    s = Select(sel)
    for opt in s.options:
        if (opt.text or "").strip().lower()=="closed": s.select_by_visible_text(opt.text); return
    for opt in s.options:
        if "closed" in (opt.get_attribute("value") or "").lower(): s.select_by_value(opt.get_attribute("value")); return
    for opt in s.options:
        if "closed" in (opt.text or "").lower(): s.select_by_visible_text(opt.text); return
    raise RuntimeError("Status=Closed not found")

def click_search(drv):
    ids = ["advSearchForm:search","advSearchForm:searchBtn","advSearchForm:searchButton",
           "bidSearchForm:search","bidSearchForm:searchBtn","bidSearchForm:searchButton",
           "searchForm:search","searchForm:searchBtn","searchForm:searchButton"]
    for cid in ids:
        try:
            el=drv.find_element(By.ID,cid); drv.execute_script("arguments[0].scrollIntoView({block:'center'});", el); el.click(); return
        except Exception: pass
    for xp in ["//button[normalize-space()='Search']",
               "//*[@role='button' and contains(normalize-space(.),'Search')]",
               "//input[@type='submit' and translate(@value,'SEARCH','search')='search']"]:
        try:
            el=drv.find_element(By.XPATH,xp); drv.execute_script("arguments[0].scrollIntoView({block:'center'});", el); el.click(); return
        except Exception: pass
    try:
        form=drv.find_element(By.XPATH,"//form[contains(@id,'advSearchForm') or contains(@name,'advSearchForm')]")
        drv.execute_script("arguments[0].submit();", form); return
    except Exception: pass
    drv.find_element(By.TAG_NAME,"body").send_keys(Keys.ENTER)

# ---------- results grid targeting ----------
def find_results_table_context(drv) -> Optional[Tuple[object,object,object,object]]:
    tables = drv.find_elements(By.XPATH,"//table[.//thead]")
    for tbl in tables:
        try:
            thead=tbl.find_element(By.TAG_NAME,"thead")
            labels=[normalize(th.text) for th in thead.find_elements(By.TAG_NAME,"th")]
            if any("bid solicitation" in lab.lower() for lab in labels):
                tbody=tbl.find_element(By.XPATH,".//tbody[1]")
                try: pager = tbl.find_element(By.XPATH,".//following::*[contains(@class,'ui-paginator')][1]")
                except NoSuchElementException: pager=None
                return (tbl, thead, tbody, pager)
        except Exception: continue
    return None

def switch_to_results_window_and_table(drv, timeout=35):
    end=time.time()+timeout
    while time.time()<end:
        for h in drv.window_handles:
            try: drv.switch_to.window(h)
            except Exception: continue
            ctx=find_results_table_context(drv)
            if ctx: return h, ctx
        time.sleep(0.4)
    for h in drv.window_handles:
        try: drv.switch_to.window(h); ctx=find_results_table_context(drv)
        except Exception: ctx=None
        if ctx: return h, ctx
    raise RuntimeError("Could not locate results grid after Search.")

def header_index_map(thead) -> Dict[str, Optional[int]]:
    ths=thead.find_elements(By.TAG_NAME,"th")
    labels=[normalize(th.text) for th in ths]
    mp={}
    for want in TARGET_HEADERS:
        idx=None
        for i,l in enumerate(labels):
            if l.lower()==want.lower(): idx=i; break
        if idx is None:
            wn=re.sub(r"[^a-z0-9]+","",want.lower())
            for i,l in enumerate(labels):
                if re.sub(r"[^a-z0-9]+","",l.lower())==wn: idx=i; break
        if idx is None:
            for i,l in enumerate(labels):
                if want.lower() in l.lower(): idx=i; break
        mp[want]=idx
    return mp

def rows_in_tbody(tbody) -> List:
    rows=tbody.find_elements(By.CSS_SELECTOR,":scope > tr")
    return [r for r in rows if r.find_elements(By.CSS_SELECTOR,":scope > td, :scope > th")]

def paginator_has_more(pager_root) -> bool:
    if not pager_root: return False
    try:
        nxt=pager_root.find_element(By.XPATH,".//a[contains(@class,'ui-paginator-next') or @aria-label='Next Page' or @title='Next']")
        if "ui-state-disabled" not in (nxt.get_attribute("class") or "").lower(): return True
    except Exception: pass
    try:
        act=pager_root.find_element(By.XPATH,".//*[contains(@class,'ui-paginator-page') and contains(@class,'ui-state-active')]")
        cur=int(act.text.strip())
        for a in pager_root.find_elements(By.XPATH,".//a[contains(@class,'ui-paginator-page')]"):
            try:
                if int(a.text.strip())>cur: return True
            except ValueError: continue
    except Exception: pass
    return False

def paginator_click_next_and_wait(drv, table, pager_root, timeout=30) -> bool:
    try:
        try:
            before_first = table.find_element(By.XPATH, ".//tbody[1]/tr[1]").text
        except Exception:
            before_first = ""
        try:
            nxt=pager_root.find_element(By.XPATH,".//a[contains(@class,'ui-paginator-next') or @aria-label='Next Page' or @title='Next']")
            if "ui-state-disabled" not in (nxt.get_attribute("class") or "").lower():
                drv.execute_script("arguments[0].scrollIntoView({block:'center'});", nxt); nxt.click()
        except Exception:
            act=pager_root.find_element(By.XPATH,".//*[contains(@class,'ui-paginator-page') and contains(@class,'ui-state-active')]")
            cur=int(act.text.strip())
            nxt=pager_root.find_element(By.XPATH,f".//a[contains(@class,'ui-paginator-page') and normalize-space()='{cur+1}']")
            drv.execute_script("arguments[0].scrollIntoView({block:'center'});", nxt); nxt.click()

        def changed(_):
            try:
                ctx = find_results_table_context(drv)
                if not ctx: return False
                tbl, th, tb, pg = ctx
                now_first = tbl.find_element(By.XPATH, ".//tbody[1]/tr[1]").text
                return now_first != before_first
            except Exception:
                return False
        WebDriverWait(drv, timeout).until(changed)
        return True
    except Exception:
        return False

# ---------- detail page (robust attachment detection) ----------
def find_attachment_links(drv) -> List[object]:
    anchors: List[object] = []
    try:
        label=None
        for xp in ["//td[normalize-space()='File Attachments:']",
                   "//td[normalize-space()='File Attachments']",
                   "//td[contains(normalize-space(.),'File Attachments')]"]:
            els = drv.find_elements(By.XPATH, xp)
            if els: label=els[0]; break
        if label:
            drv.execute_script("arguments[0].scrollIntoView({block:'center'});", label)
            cand = label.find_elements(By.XPATH, "following-sibling::td[1]//a[@href]")
            for a in cand:
                href=(a.get_attribute('href') or ""); txt=(a.text or "")
                if FILE_PAT.search(href) or FILE_PAT.search(txt): anchors.append(a)
    except Exception: pass

    if not anchors:
        try:
            base = drv.find_elements(By.XPATH, "(//*[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'file attachments')])[1]")
            if base:
                drv.execute_script("arguments[0].scrollIntoView({block:'center'});", base[0])
                cand = drv.find_elements(By.XPATH, "(//*[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'file attachments')])[1]/following::a[@href]")
                for a in cand[:20]:
                    href=(a.get_attribute('href') or ""); txt=(a.text or "")
                    if FILE_PAT.search(href) or FILE_PAT.search(txt): anchors.append(a)
        except Exception: pass

    if not anchors:
        try:
            all_as = drv.find_elements(By.XPATH, "//a[@href]")
            h = drv.execute_script("return document.body.scrollHeight;")
            cutoff = h*0.35 if h else 0
            for a in all_as:
                if not a.is_displayed(): continue
                href=(a.get_attribute('href') or ""); txt=(a.text or "")
                if FILE_PAT.search(href) or FILE_PAT.search(txt):
                    try:
                        top = drv.execute_script("const r=arguments[0].getBoundingClientRect(); return r.top + window.scrollY;", a)
                    except Exception:
                        top = None
                    if top is None or top >= cutoff:
                        anchors.append(a)
        except Exception: pass

    seen=set(); uniq=[]
    for a in anchors:
        key=(a.get_attribute('href') or "")+"||"+(a.text or "")
        if key not in seen:
            seen.add(key); uniq.append(a)
    return uniq

def download_attachments_from_detail(drv, download_dir: str) -> List[str]:
    before=set(os.listdir(download_dir))
    anchors = find_attachment_links(drv)
    if not anchors:
        return []
    idx=0
    while idx < len(anchors) and not stop_flag["stop"]:
        a = anchors[idx]; idx += 1
        try:
            drv.execute_script("arguments[0].scrollIntoView({block:'center'});", a)
            drv.execute_script("arguments[0].click();", a)
        except Exception:
            try: a.click()
            except Exception: continue
        wait_for_downloads(download_dir); time.sleep(MIN_WAIT_AFTER_DOWNLOAD)
        if not drv.find_elements(By.TAG_NAME, "body"):
            drv.back(); WebDriverWait(drv, 15).until(EC.presence_of_element_located((By.TAG_NAME,"body")))
        anchors = find_attachment_links(drv)
    after=set(os.listdir(download_dir))
    return sorted(list(after-before))

# ---------- main scraping ----------
def scrape_all(out_folder: str, headless: bool) -> Optional[str]:
    attachments_dir = ensure_dir(os.path.join(out_folder,"nevada_attachments"))
    drv = build_driver(headless=headless, download_dir=attachments_dir)

    results: List[dict] = []
    page_no = 1
    partial_path = None

    try:
        drv.get(ADV_URL); wait_ready(drv); set_status_closed(drv); click_search(drv)
        results_handle, ctx = switch_to_results_window_and_table(drv)
        drv.switch_to.window(results_handle)

        while not stop_flag["stop"]:
            ctx = find_results_table_context(drv) or ctx
            table, thead, tbody, pager_root = ctx
            idx_map = header_index_map(thead)

            # rows for this page
            ok=False
            for _ in range(RETRY_STALE):
                try:
                    tbody = table.find_element(By.XPATH, ".//tbody[1]")
                    rows = rows_in_tbody(tbody); ok=True; break
                except StaleElementReferenceException:
                    time.sleep(0.3); ctx = find_results_table_context(drv) or ctx; table, thead, tbody, pager_root = ctx
            if not ok:
                tbody = table.find_element(By.XPATH, ".//tbody[1]")
                rows = rows_in_tbody(tbody)

            print(f"[INFO] Page {page_no} — {len(rows)} rows")

            for i in range(len(rows)):
                if stop_flag["stop"]: break
                # fresh row ref
                tr = table.find_element(By.XPATH, f".//tbody[1]/tr[{i+1}]")
                tds = tr.find_elements(By.CSS_SELECTOR, ":scope > td, :scope > th")

                rec={}
                for col in TARGET_HEADERS:
                    j = idx_map.get(col)
                    rec[col] = normalize(tds[j].text) if j is not None and j < len(tds) else ""
                rec["Row URL"] = ""; rec["Attachment Files"] = ""

                try:
                    j = idx_map["Bid Solicitation #"]
                    a = tds[j].find_element(By.CSS_SELECTOR,"a[href]")
                    href = a.get_attribute("href")
                    drv.execute_script("window.open(arguments[0], '_blank');", href)
                    detail_handle = drv.window_handles[-1]
                    drv.switch_to.window(detail_handle)
                    try:
                        WebDriverWait(drv, 25).until(EC.presence_of_element_located((By.TAG_NAME,"body")))
                    except TimeoutException:
                        pass
                    rec["Row URL"] = drv.current_url
                    files = download_attachments_from_detail(drv, attachments_dir)
                    rec["Attachment Files"] = "; ".join(files)
                    wait_for_downloads(attachments_dir); time.sleep(MIN_WAIT_AFTER_DOWNLOAD)
                    drv.close(); drv.switch_to.window(results_handle)
                    ctx = find_results_table_context(drv) or ctx; table, thead, tbody, pager_root = ctx
                except Exception:
                    try: drv.switch_to.window(results_handle)
                    except Exception: pass
                    ctx = find_results_table_context(drv) or ctx; table, thead, tbody, pager_root = ctx

                results.append(rec)

                # row-level autosave
                if len(results) % ROW_AUTOSAVE_INTERVAL == 0:
                    partial_path = save_progress(out_folder, results, page_no, note="autosave (rows)")

            # page-level autosave
            partial_path = save_progress(out_folder, results, page_no, note="autosave (page complete)")

            if stop_flag["stop"]: break

            if pager_root and paginator_has_more(pager_root):
                moved = paginator_click_next_and_wait(drv, table, pager_root, timeout=40)
                if not moved: break
                page_no += 1
                continue
            else:
                break

        # final save
        final_path = save_progress(out_folder, results, page_no, note="complete", final=True)
        print(f"[INFO] Wrote {len(results)} rows → {final_path}")
        print(f"[INFO] Attachments → {attachments_dir}")
        return final_path

    except (KeyboardInterrupt, WebDriverException, Exception) as e:
        # graceful partial save on any crash/close
        note = f"stopped due to: {type(e).__name__}"
        partial_path = save_progress(out_folder, results, page_no, note=note)
        print(f"[WARN] Interrupted, saved partial results → {partial_path}")
        print(f"[WARN] Reason: {e}")
        return partial_path
    finally:
        try: drv.quit()
        except Exception: pass

# ---------- CLI ----------
def main():
    ap = argparse.ArgumentParser(description="Nevada ePro Closed — autosave & robust attachment downloads")
    ap.add_argument("--out", default=".", help="Output folder")
    ap.add_argument("--headless", action="store_true", help="Run Chrome headless")
    args = ap.parse_args()
    out_dir = ensure_dir(args.out)
    path = scrape_all(out_dir, headless=args.headless)
    if not path: sys.exit(2)

if __name__ == "__main__":
    main()
