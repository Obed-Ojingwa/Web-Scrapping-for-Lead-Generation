"""
Nigeria Company Leads Scraper
Author: Senior Data Engineer Script
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import time
import random
from urllib.parse import urljoin
from concurrent.futures import ThreadPoolExecutor, as_completed

# Optional (install if needed)
try:
    from googlesearch import search
except:
    print("Install googlesearch-python if not installed.")

# ---------------------------
# CONFIGURATION
# ---------------------------

SEARCH_KEYWORDS = [
    "construction companies in Nigeria",
    "civil engineering firms Nigeria",
    "agriculture companies Nigeria",
    "agritech startups Nigeria",
    "oil and gas companies Nigeria",
    "oilfield services Nigeria",
    "surveying companies Nigeria",
    "GIS mapping companies Nigeria"
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}

MAX_RESULTS_PER_QUERY = 20
MAX_WORKERS = 10
REQUEST_DELAY = (1, 3)

OUTPUT_FILE = "nigeria_company_leads.xlsx"

# ---------------------------
# REGEX PATTERNS
# ---------------------------

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")
PHONE_REGEX = re.compile(r"(\+?234|0)[789][01]\d{8}")

# ---------------------------
# HELPERS
# ---------------------------

def normalize_phone(phone):
    phone = re.sub(r"\D", "", phone)
    if phone.startswith("0"):
        return "+234" + phone[1:]
    if phone.startswith("234"):
        return "+" + phone
    if phone.startswith("+234"):
        return phone
    return None


def valid_email(email):
    return re.match(EMAIL_REGEX, email) is not None


def fetch_page(url):
    try:
        time.sleep(random.uniform(*REQUEST_DELAY))
        response = requests.get(url, headers=HEADERS, timeout=10)
        if response.status_code == 200:
            return response.text
    except:
        return None


def extract_contacts(html):
    emails = set(EMAIL_REGEX.findall(html))
    phones = set(PHONE_REGEX.findall(html))

    phones = {normalize_phone(p) for p in phones if normalize_phone(p)}

    return list(emails), list(phones)


def extract_text(soup):
    return soup.get_text(" ", strip=True)


def extract_company_name(soup, url):
    if soup.title:
        return soup.title.string.strip()
    return url.split("//")[-1].split("/")[0]


def find_linkedin(soup):
    for a in soup.find_all("a", href=True):
        if "linkedin.com/company" in a["href"]:
            return a["href"]
    return ""


def extract_address(text):
    # Simple heuristic for Nigerian addresses
    keywords = ["Lagos", "Abuja", "Port Harcourt", "Kano", "Nigeria"]
    for line in text.split("."):
        if any(k in line for k in keywords):
            return line.strip()
    return ""


# ---------------------------
# SCRAPE SINGLE WEBSITE
# ---------------------------

def scrape_website(url, industry):
    html = fetch_page(url)
    if not html:
        return None

    soup = BeautifulSoup(html, "html.parser")

    text = extract_text(soup)
    emails, phones = extract_contacts(html)

    if not emails and not phones:
        return None

    company_name = extract_company_name(soup, url)
    linkedin = find_linkedin(soup)
    address = extract_address(text)

    return {
        "Company Name": company_name,
        "Industry": industry,
        "Services": "",
        "Address": address,
        "City": "",
        "State": "",
        "Country": "Nigeria",
        "Email": ", ".join(emails),
        "Phone": ", ".join(phones),
        "Website": url,
        "LinkedIn": linkedin
    }


# ---------------------------
# SEARCH ENGINE SCRAPER
# ---------------------------

def search_companies(query):
    urls = []
    try:
        for result in search(query, num_results=MAX_RESULTS_PER_QUERY):
            urls.append(result)
    except:
        pass
    return urls


# ---------------------------
# MAIN SCRAPER
# ---------------------------

def run_scraper():
    all_results = []
    seen_websites = set()

    print("🔍 Searching for companies...")

    search_results = []

    for keyword in SEARCH_KEYWORDS:
        urls = search_companies(keyword)
        for url in urls:
            if url not in seen_websites:
                search_results.append((url, keyword))
                seen_websites.add(url)

    print(f"Found {len(search_results)} websites.")

    print("🚀 Scraping websites...")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [
            executor.submit(scrape_website, url, industry)
            for url, industry in search_results
        ]

        for future in as_completed(futures):
            result = future.result()
            if result:
                all_results.append(result)

    print(f"Collected {len(all_results)} raw records.")

    if not all_results:
        print("❌ No data collected. Check your search module or internet connection.")
        return

    df = pd.DataFrame(all_results)

    # ---------------------------
    # DATA CLEANING
    # ---------------------------

    df.drop_duplicates(subset=["Company Name", "Website"], inplace=True)

    df = df[
        (df["Email"].str.strip() != "") |
        (df["Phone"].str.strip() != "")
    ]

    df.reset_index(drop=True, inplace=True)

    # ---------------------------
    # EXPORT
    # ---------------------------

    df.to_excel(OUTPUT_FILE, index=False)

    print(f"✅ Data saved to {OUTPUT_FILE}")
    print(f"Final dataset size: {len(df)} records")


# ---------------------------
# ENTRY POINT
# ---------------------------

if __name__ == "__main__":
    run_scraper()