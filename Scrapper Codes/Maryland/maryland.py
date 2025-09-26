#!/usr/bin/env python3
# maryland.py — EMMA Public Solicitations: Status=Closed + Award Status=Awarded → Excel (headless)

import argparse, os, sys, re
from datetime import datetime
from typing import List, Dict

import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright, Page

URL = "https://emma.maryland.gov/page.aspx/en/rfp/request_browse_public"

# --------------------- generic waits ---------------------

def wait_form_ready(page: Page, timeout_ms: int = 40000):
    page.wait_for_selector("form, select, [role=combobox], button", timeout=timeout_ms)

def wait_results_ready(page: Page, timeout_ms: int = 30000):
    try:
        page.locator("table thead tr").first.wait_for(timeout=timeout_ms)
        return
    except Exception:
        pass
    page.locator("table tbody tr").first.wait_for(timeout=timeout_ms)

# --------------------- dropdown helpers ---------------------

def _try_native_select(page: Page, label_text: str, wanted: str) -> bool:
    """Try selecting via real <select> associated to a label."""
    # 1) exact label mapping
    try:
        sel = page.get_by_label(label_text, exact=True)
        sel.select_option(label=wanted)
        return True
    except Exception:
        pass
    # 2) nearest select after a label element
    try:
        lab = page.locator(f"label:has-text('{label_text}')").first
        lab.wait_for(timeout=1500)
        sel = lab.locator("xpath=following::select[1]").first
        sel.select_option(label=wanted)
        return True
    except Exception:
        pass
    return False

def _try_combobox_click(page: Page, label_text: str, wanted: str) -> bool:
    """
    Handle custom dropdowns (Select2/Kendo/etc.) exposed as a combobox.
    Strategy:
      - Find combobox near the label
      - Click to open listbox
      - Click option by visible text (role=option|li|div)
    """
    # Find a combobox by accessible name
    cb = page.get_by_role("combobox", name=re.compile(rf"^{re.escape(label_text)}\b", re.I)).first
    if cb.count() == 0:
        # look for a combobox near the label text
        try:
            lab = page.get_by_text(re.compile(rf"^{re.escape(label_text)}\b", re.I)).first
            lab.wait_for(timeout=1500)
            cb = lab.locator("xpath=following::*[@role='combobox'][1]").first
        except Exception:
            pass
    if cb.count() == 0:
        # sometimes the clickable element is just a button next to label
        try:
            lab = page.get_by_text(re.compile(rf"^{re.escape(label_text)}\b", re.I)).first
            lab.wait_for(timeout=1500)
            cb = lab.locator("xpath=following::button[1]").first
        except Exception:
            pass
    if cb.count() == 0:
        return False

    cb.click(timeout=3000)

    # Options commonly render with role=option; fallback to list items/buttons
    # Use has-text to match the visible label
    for sel in (
        f"[role=option]:has-text('{wanted}')",
        f"li[role=option]:has-text('{wanted}')",
        f"li:has-text('{wanted}')",
        f"div[role=option]:has-text('{wanted}')",
        f"button:has-text('{wanted}')",
    ):
        opt = page.locator(sel).first
        if opt.count() > 0:
            opt.click(timeout=3000)
            return True

    # final fallback: type to filter & press Enter
    try:
        cb.fill(wanted)
        page.keyboard.press("Enter")
        return True
    except Exception:
        return False

def set_dropdown(page: Page, label_text: str, wanted: str):
    """
    Robustly set a dropdown by visible label to the wanted option,
    handling both native <select> and custom combobox widgets.
    """
    if _try_native_select(page, label_text, wanted):
        return
    if _try_combobox_click(page, label_text, wanted):
        return
    raise RuntimeError(f"Could not select '{wanted}' in '{label_text}'")

# --------------------- UI bits ---------------------

def expand_advanced_if_needed(page: Page):
    # If Award Status not visible, click Advanced Search chevron/header
    try:
        page.get_by_label("Award Status", exact=True).first.wait_for(timeout=1500)
        return
    except Exception:
        pass
    for sel in ("button:has-text('Advanced Search')",
                "a:has-text('Advanced Search')",
                "text=Advanced Search"):
        try:
            page.locator(sel).first.click(timeout=1500)
            page.get_by_label("Award Status", exact=True).first.wait_for(timeout=3000)
            return
        except Exception:
            continue

def click_search(page: Page):
    for sel in ("button:has-text('Search')",
                "input[type='submit'][value*='Search']",
                "input[type='button'][value*='Search']"):
        try:
            page.locator(sel).first.click(timeout=4000)
            return
        except Exception:
            continue
    page.evaluate("""() => {
        const b = Array.from(document.querySelectorAll('button,input[type=submit],input[type=button]'))
          .find(x => (x.innerText||x.value||'').toLowerCase().includes('search'));
        if (b) b.click();
    }""")

def pager_has_next(page: Page) -> bool:
    for sel in ("a:has-text('Next')", "button:has-text('Next')", "a[aria-label*='Next']"):
        el = page.locator(sel).first
        if el.count() == 0:
            continue
        cls = (el.get_attribute("class") or "").lower()
        aria = (el.get_attribute("aria-disabled") or "").lower()
        if "disabled" in cls or aria == "true":
            continue
        return True
    return False

def pager_next(page: Page) -> bool:
    for sel in ("a:has-text('Next')", "button:has-text('Next')", "a[aria-label*='Next']"):
        el = page.locator(sel).first
        if el.count() == 0:
            continue
        cls = (el.get_attribute("class") or "").lower()
        aria = (el.get_attribute("aria-disabled") or "").lower()
        if "disabled" in cls or aria == "true":
            continue
        try:
            el.click(timeout=3000)
            page.wait_for_timeout(600)
            return True
        except Exception:
            continue
    return False

# --------------------- parsing ---------------------

def parse_rows(html: str) -> List[Dict]:
    soup = BeautifulSoup(html, "html.parser")
    best_tbl = None
    best_headers = []
    best_score = -1

    for tbl in soup.find_all("table"):
        thead, tbody = tbl.find("thead"), tbl.find("tbody")
        if not thead or not tbody:
            continue
        headers = [th.get_text(strip=True) for th in thead.find_all("th")]
        score = sum(k in " ".join(h.lower() for h in headers)
                    for k in ["id", "title", "status", "due", "close", "publish", "solicitation", "issuing"])
        if score > best_score:
            best_tbl, best_headers, best_score = tbl, headers, score

    rows: List[Dict] = []
    if not best_tbl:
        return rows

    # map headers to keys
    keys = []
    for h in best_headers:
        hl = h.lower()
        if hl.strip() == "id":
            keys.append("id")
        elif "title" in hl:
            keys.append("title")
        elif "status" in hl:
            keys.append("status")
        elif "due" in hl or "close" in hl:
            keys.append("due_close_date")
        elif "publish" in hl:
            keys.append("publish_date")
        elif "solicitation type" in hl:
            keys.append("solicitation_type")
        elif "issuing" in hl:
            keys.append("issuing_agency")
        elif "main category" in hl:
            keys.append("main_category")
        else:
            keys.append(hl.replace(" ", "_"))

    for tr in best_tbl.find("tbody").find_all("tr"):
        tds = tr.find_all("td")
        if not tds:
            continue
        rec = {}
        for i, key in enumerate(keys):
            val = tds[i].get_text(" ", strip=True) if i < len(tds) else ""
            if key == "title":
                a = tds[i].find("a", href=True)
                if a:
                    href = a["href"]
                    if href.startswith("/"):
                        href = "https://emma.maryland.gov" + href
                    rec["detail_url"] = href
                    rec[key] = a.get_text(strip=True)
                else:
                    rec[key] = val
            else:
                rec[key] = val
        rows.append(rec)
    return rows

# --------------------- main ---------------------

def main():
    ap = argparse.ArgumentParser(description="EMMA: Closed + Awarded → Excel (headless)")
    ap.add_argument("--out", default=None, help="Output Excel path")
    ap.add_argument("--show", action="store_true", help="Show browser (debug)")
    ap.add_argument("--max-pages", type=int, default=1000, help="Safety cap on pages")
    args = ap.parse_args()

    out_path = args.out or os.path.abspath(
        f"./maryland_emma_closed_awarded_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    )

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=not args.show)
        ctx = browser.new_context(locale="en-US", viewport={"width": 1500, "height": 1100})
        page = ctx.new_page()
        try:
            page.goto(URL, wait_until="domcontentloaded", timeout=45000)
            wait_form_ready(page)

            # Status = Closed  (handles native select or custom combobox)
            set_dropdown(page, "Status", "Closed")

            # Advanced Search -> Award Status = Awarded
            expand_advanced_if_needed(page)
            set_dropdown(page, "Award Status", "Awarded")

            # Search
            click_search(page)

            # Collect all pages
            all_rows: List[Dict] = []
            pages = 0
            while pages < args.max_pages:
                wait_results_ready(page)
                rows = parse_rows(page.content())
                all_rows.extend(rows)
                if not pager_has_next(page):
                    break
                if not pager_next(page):
                    break
                pages += 1

            # Save Excel
            df = pd.DataFrame(all_rows)
            ordered = ["id","title","status","due_close_date","publish_date",
                       "main_category","solicitation_type","issuing_agency","detail_url"]
            cols = [c for c in ordered if c in df.columns] + [c for c in df.columns if c not in ordered]
            df[cols].to_excel(out_path, index=False)
            print(f"[OK] Wrote {len(df)} rows to: {out_path}")

        finally:
            ctx.close()
            browser.close()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[FATAL] {e}")
        sys.exit(1)
