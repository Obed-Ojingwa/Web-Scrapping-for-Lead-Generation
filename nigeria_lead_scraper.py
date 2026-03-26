"""
Nigeria Corporate Lead Generation Scraper
==========================================
Searches the web for companies in Nigeria across:
  - Construction & Civil Engineering
  - Agriculture / Agritech / Agro-processing
  - Oil & Gas (Upstream, Midstream, Downstream, Oilfield Services)
  - Cadastral Survey / Land Surveying / Geospatial / GIS / Drone Mapping

REQUIREMENTS (install before running):
    pip install requests beautifulsoup4 googlesearch-python openpyxl lxml tqdm

USAGE:
    python nigeria_lead_scraper.py

OUTPUT:
    nigeria_leads_<timestamp>.xlsx  — Excel file with all results
    nigeria_leads_<timestamp>.csv   — CSV backup

HOW IT WORKS:
  1. Fires Google search queries per sector (via googlesearch-python)
  2. Fetches each result page and scrapes emails, phones, addresses
  3. Also tries common contact-page paths (/contact, /about, /contact-us)
  4. Deduplicates, validates, and writes to Excel

NOTE ON RATE LIMITS:
  Google will throttle aggressive scraping. The script uses:
    - Random delays between requests (2–6 seconds)
    - Rotating User-Agent strings
  For heavy use, swap SEARCH_BACKEND to "serpapi" (requires free API key)
  or "bing" (no key needed).
"""

import re
import csv
import time
import random
import logging
import hashlib
from datetime import datetime
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── optional: tqdm for progress bars ────────────────────────────────────────
try:
    from tqdm import tqdm
except ImportError:
    def tqdm(x, **kw): return x

# ─── optional: googlesearch-python ───────────────────────────────────────────
try:
    from googlesearch import search as google_search
    GOOGLE_AVAILABLE = True
except ImportError:
    GOOGLE_AVAILABLE = False
    print("[WARN] googlesearch-python not installed. Using Bing fallback.")

logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")
log = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════

# How many Google results to fetch per query
RESULTS_PER_QUERY = 8

# Seconds to sleep between HTTP requests (be polite to servers)
MIN_DELAY = 2.0
MAX_DELAY = 5.5

# Request timeout in seconds
REQUEST_TIMEOUT = 12

# Max pages to crawl per domain (1 = homepage only, 2 = homepage + contact page)
MAX_PAGES_PER_DOMAIN = 2

# ── Search queries by sector ──────────────────────────────────────────────────
SEARCH_QUERIES = {

    "Construction": [
        "large construction civil engineering companies Nigeria official website email",
        "construction company Nigeria Lagos Abuja site:linkedin.com/company",
        "building infrastructure company Nigeria contact email phone",
        "road bridge construction Nigeria contractor official website",
        "Abuja Lagos construction company email info@ Nigeria",
        "top construction companies Nigeria 2024 email contact",
        "civil engineering firm Nigeria official website contact us",
        "Nigeria construction company Lagos Abuja Port Harcourt email address",
    ],

    "Agriculture": [
        "agribusiness agritech company Nigeria official website email contact",
        "agriculture processing company Nigeria Lagos email info@",
        "farm estate Nigeria large-scale email contact website",
        "Nigeria agro-processing company email phone official site",
        "precision agriculture company Nigeria contact website 2024",
        "Nigeria food processing company email official website",
        "Nigeria farming estate company contact email website",
        "agritech startup Nigeria Lagos Abuja email contact",
    ],

    "Oil & Gas": [
        "oilfield services company Nigeria Lagos Port Harcourt email official website",
        "upstream oil gas company Nigeria email contact official site",
        "Nigeria oil gas exploration company email phone website 2024",
        "pipeline construction company Nigeria email contact address",
        "downstream oil gas company Nigeria official website email",
        "oil gas services Nigeria Port Harcourt email info@",
        "Nigeria midstream petroleum company email phone address",
        "Nigeria EPC company oil gas contact email website",
    ],

    "Cadastral Survey": [
        "land surveying GIS company Nigeria email official website",
        "cadastral survey company Nigeria Lagos Abuja email contact",
        "drone mapping geospatial services Nigeria email phone website",
        "Nigeria GIS mapping company email info@ official site",
        "surveying geoinformatics Nigeria contact website email 2024",
        "aerial survey photogrammetry Nigeria email contact official",
        "Nigeria remote sensing geospatial company email address",
        "drone survey LiDAR company Nigeria email contact website",
    ],
}

# ── User-Agent pool ───────────────────────────────────────────────────────────
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_3) AppleWebKit/605.1.15 "
    "(KHTML, like Gecko) Version/17.2 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64; rv:123.0) Gecko/20100101 Firefox/123.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
]

# ── Common contact-page suffixes to try ──────────────────────────────────────
CONTACT_PATHS = ["/contact", "/contact-us", "/about", "/about-us",
                 "/reach-us", "/get-in-touch", "/contactus"]

# ══════════════════════════════════════════════════════════════════════════════
# REGEX PATTERNS
# ══════════════════════════════════════════════════════════════════════════════

EMAIL_RE    = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", re.I)
PHONE_RE    = re.compile(
    r"(?:\+?234[\s\-.]?)?(?:0[789]\d[\s\-.]?\d{3}[\s\-.]?\d{4}|"
    r"0[789]\d{8}|\+234[\s\-.]?\d{2,3}[\s\-.]?\d{3,4}[\s\-.]?\d{4})",
    re.I
)
# Addresses: looks for street-number patterns common in Nigeria
ADDRESS_RE  = re.compile(
    r"\b(?:Plot|No\.?|House|Block|Suite|Flat|KM|Km|Road|Street|Close|"
    r"Avenue|Drive|Way|Crescent|Boulevard)\b[^\n<]{5,80}",
    re.I
)
# Nigeria state names for city/state extraction
NG_STATES   = re.compile(
    r"\b(Lagos|Abuja|FCT|Kano|Rivers|Port Harcourt|Kaduna|Oyo|Ibadan|"
    r"Enugu|Edo|Benin City|Delta|Imo|Ogun|Anambra|Bauchi|Plateau|"
    r"Cross River|Borno|Sokoto|Kwara|Kebbi|Adamawa|Taraba|Nasarawa|"
    r"Ekiti|Osun|Ondo|Bayelsa|Akwa Ibom|Ebonyi|Gombe|Niger|Zamfara|Jigawa|Katsina|Yobe)\b",
    re.I
)

# Domains to skip (generic, not company sites)
SKIP_DOMAINS = {
    "facebook.com", "twitter.com", "instagram.com", "youtube.com",
    "wikipedia.org", "nairaland.com", "indeed.com", "jobberman.com",
    "glassdoor.com", "yellowpages.com.ng", "vconnect.com",
    "businesslist.com.ng", "naijagoodnews.com", "google.com",
    "bing.com", "yahoo.com",
}

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def rand_headers():
    return {
        "User-Agent": random.choice(USER_AGENTS),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate",
        "Connection": "keep-alive",
    }

def polite_sleep():
    time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))

def clean_phone(raw: str) -> str:
    """Normalise phone to +234 format."""
    digits = re.sub(r"[^\d+]", "", raw)
    if digits.startswith("0") and len(digits) >= 11:
        return "+234" + digits[1:]
    if digits.startswith("234") and not digits.startswith("+"):
        return "+" + digits
    return digits if digits.startswith("+234") else raw.strip()

def domain_of(url: str) -> str:
    return urlparse(url).netloc.replace("www.", "")

def should_skip(url: str) -> bool:
    d = domain_of(url)
    return any(skip in d for skip in SKIP_DOMAINS)

def company_name_from(soup: BeautifulSoup, url: str) -> str:
    """Best-effort company name from page."""
    # Try OG title
    og = soup.find("meta", property="og:site_name")
    if og and og.get("content"):
        return og["content"].strip()
    # Try <title>
    title = soup.find("title")
    if title and title.text:
        t = title.text.split("|")[0].split("-")[0].strip()
        return t[:80]
    return domain_of(url).split(".")[0].title()

def linkedin_from(soup: BeautifulSoup) -> str:
    """Find LinkedIn link in page."""
    for a in soup.find_all("a", href=True):
        h = a["href"]
        if "linkedin.com/company" in h:
            return h.split("?")[0]
    return ""

# ══════════════════════════════════════════════════════════════════════════════
# SCRAPER CORE
# ══════════════════════════════════════════════════════════════════════════════

def fetch_page(url: str) -> BeautifulSoup | None:
    try:
        polite_sleep()
        r = requests.get(url, headers=rand_headers(), timeout=REQUEST_TIMEOUT,
                         allow_redirects=True)
        r.raise_for_status()
        return BeautifulSoup(r.text, "lxml")
    except Exception as e:
        log.debug(f"Fetch failed {url}: {e}")
        return None

def extract_from_page(soup: BeautifulSoup, url: str) -> dict:
    text = soup.get_text(separator=" ", strip=True)

    emails  = list(set(EMAIL_RE.findall(text)))
    phones  = list(set(clean_phone(p) for p in PHONE_RE.findall(text)))
    addrs   = ADDRESS_RE.findall(text)
    states  = NG_STATES.findall(text)

    # Filter junk emails (images, assets, etc.)
    emails = [e for e in emails if not re.search(
        r"\.(png|jpg|gif|svg|css|js|woff|ttf)$", e, re.I)]

    # Prefer info@ / contact@ / hello@ / admin@
    primary_email = ""
    for pref in ("info@", "contact@", "hello@", "admin@", "enquir", "enqui"):
        match = next((e for e in emails if e.lower().startswith(pref)), None)
        if match:
            primary_email = match
            break
    if not primary_email and emails:
        primary_email = emails[0]

    primary_phone = phones[0] if phones else ""
    primary_addr  = addrs[0].strip() if addrs else ""
    primary_state = states[0] if states else ""

    # Try to detect city from address line
    city = ""
    if primary_state:
        city_match = re.search(rf"([A-Za-z\s]+),?\s*{primary_state}", primary_addr, re.I)
        city = city_match.group(1).strip() if city_match else primary_state

    return {
        "email":   primary_email,
        "phone":   primary_phone,
        "address": primary_addr,
        "city":    city,
        "state":   primary_state,
        "linkedin": linkedin_from(soup),
        "name":    company_name_from(soup, url),
        "website": url,
        "all_emails": emails,
        "all_phones": phones,
    }

def scrape_company(base_url: str) -> dict | None:
    """Scrape homepage + contact page and merge results."""
    if should_skip(base_url):
        return None

    soup = fetch_page(base_url)
    if not soup:
        return None

    data = extract_from_page(soup, base_url)

    # If homepage lacked contact info, try contact sub-pages
    if not data["email"] or not data["phone"]:
        for path in CONTACT_PATHS[:MAX_PAGES_PER_DOMAIN]:
            contact_url = urljoin(base_url, path)
            csoup = fetch_page(contact_url)
            if not csoup:
                continue
            extra = extract_from_page(csoup, base_url)
            if not data["email"] and extra["email"]:
                data["email"] = extra["email"]
            if not data["phone"] and extra["phone"]:
                data["phone"] = extra["phone"]
            if not data["address"] and extra["address"]:
                data["address"] = extra["address"]
            if not data["state"] and extra["state"]:
                data["state"] = extra["state"]
            if not data["linkedin"] and extra["linkedin"]:
                data["linkedin"] = extra["linkedin"]
            if data["email"] and data["phone"]:
                break  # got what we need

    return data

def google_search_urls(query: str, num: int) -> list[str]:
    """Return URLs from Google search. Falls back to Bing if unavailable."""
    urls = []
    if GOOGLE_AVAILABLE:
        try:
            for url in google_search(query, num_results=num, lang="en",
                                      sleep_interval=3, safe="off"):
                urls.append(url)
        except Exception as e:
            log.warning(f"Google search error: {e}")
    else:
        urls = bing_search_urls(query, num)
    return [u for u in urls if not should_skip(u)]

def bing_search_urls(query: str, num: int) -> list[str]:
    """Scrape Bing search results as fallback."""
    urls = []
    try:
        polite_sleep()
        params = {"q": query, "count": num, "form": "QBLH"}
        r = requests.get("https://www.bing.com/search", params=params,
                         headers=rand_headers(), timeout=REQUEST_TIMEOUT)
        soup = BeautifulSoup(r.text, "lxml")
        for a in soup.select("li.b_algo h2 a"):
            href = a.get("href", "")
            if href.startswith("http"):
                urls.append(href)
    except Exception as e:
        log.warning(f"Bing search error: {e}")
    return urls[:num]

# ══════════════════════════════════════════════════════════════════════════════
# MAIN PIPELINE
# ══════════════════════════════════════════════════════════════════════════════

def run_scraper() -> list[dict]:
    all_results = []
    seen_domains = set()
    seen_emails  = set()

    for sector, queries in SEARCH_QUERIES.items():
        log.info(f"\n{'='*60}\nSector: {sector}\n{'='*60}")

        for query in tqdm(queries, desc=sector):
            log.info(f"  Searching: {query}")
            urls = google_search_urls(query, RESULTS_PER_QUERY)

            for url in urls:
                d = domain_of(url)
                if d in seen_domains:
                    continue
                seen_domains.add(d)

                log.info(f"    Scraping: {url}")
                data = scrape_company(url)
                if not data:
                    continue

                # Must have at least one contact method
                if not data["email"] and not data["phone"]:
                    log.debug(f"    No contact found: {url}")
                    continue

                # Deduplicate by email
                if data["email"] and data["email"] in seen_emails:
                    continue
                if data["email"]:
                    seen_emails.add(data["email"])

                row = {
                    "Company Name":  data["name"],
                    "Industry":      sector,
                    "Services":      "",           # user can fill in
                    "Address":       data["address"],
                    "City":          data["city"],
                    "State":         data["state"],
                    "Country":       "Nigeria",
                    "Email":         data["email"],
                    "Phone":         data["phone"],
                    "Website":       data["website"],
                    "LinkedIn":      data["linkedin"],
                }
                all_results.append(row)
                log.info(f"    ✓ Added: {data['name']} | {data['email']} | {data['phone']}")

    return all_results

# ══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ══════════════════════════════════════════════════════════════════════════════

COLUMNS = ["Company Name", "Industry", "Services", "Address",
           "City", "State", "Country", "Email", "Phone", "Website", "LinkedIn"]

HDR_FILL  = PatternFill("solid", start_color="1F4E79")
HDR_FONT  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
ROW_FONT  = Font(name="Arial", size=9)
ALT_FILL  = PatternFill("solid", start_color="DCE6F1")
THIN      = Side(style="thin", color="B8CCE4")
BRD       = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP      = Alignment(wrap_text=True, vertical="top")
CTR       = Alignment(horizontal="center", vertical="top")

COL_WIDTHS = [32, 22, 28, 40, 16, 16, 10, 36, 22, 36, 42]

def to_excel(rows: list[dict], filepath: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Nigeria Leads"

    # Header
    for ci, (col, w) in enumerate(zip(COLUMNS, COL_WIDTHS), 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.font = HDR_FILL and HDR_FONT
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = CTR
        c.border = BRD
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 22

    # Rows
    for ri, row in enumerate(rows, 2):
        for ci, col in enumerate(COLUMNS, 1):
            c = ws.cell(row=ri, column=ci, value=row.get(col, ""))
            c.font = ROW_FONT
            c.border = BRD
            c.alignment = WRAP
            if ri % 2 == 0:
                c.fill = ALT_FILL

    # Auto-filter
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"

    # Per-sector summary sheet
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Sector"
    ws2["B1"] = "Count"
    for ci, h in enumerate(["A1", "B1"]):
        c = ws2[h]
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.border = BRD
    from collections import Counter
    counts = Counter(r["Industry"] for r in rows)
    for ri, (sector, cnt) in enumerate(counts.items(), 2):
        ws2.cell(row=ri, column=1, value=sector).border = BRD
        ws2.cell(row=ri, column=2, value=cnt).border = BRD
    ws2["A1"].alignment = CTR
    ws2["B1"].alignment = CTR
    ws2.column_dimensions["A"].width = 26
    ws2.column_dimensions["B"].width = 10

    wb.save(filepath)
    log.info(f"\nExcel saved → {filepath}")

def to_csv(rows: list[dict], filepath: str):
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(rows)
    log.info(f"CSV saved   → {filepath}")

# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("""
╔══════════════════════════════════════════════════════════════╗
║       Nigeria Corporate Lead Generation Scraper             ║
║   Sectors: Construction | Agriculture | Oil&Gas | GIS       ║
╚══════════════════════════════════════════════════════════════╝
    """)

    ts       = datetime.now().strftime("%Y%m%d_%H%M")
    xl_path  = f"nigeria_leads_{ts}.xlsx"
    csv_path = f"nigeria_leads_{ts}.csv"

    results = run_scraper()

    if results:
        to_excel(results, xl_path)
        to_csv(results,   csv_path)
        print(f"\n✅  Done!  {len(results)} companies found.")
        print(f"   Excel → {xl_path}")
        print(f"   CSV   → {csv_path}")
    else:
        print("\n⚠  No results found. Check your internet connection "
              "or install googlesearch-python.")