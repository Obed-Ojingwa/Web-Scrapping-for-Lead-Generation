"""
Nigeria Company Leads Scraper (DuckDuckGo Optimized)
Author: Senior Data Engineer Upgrade
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import time
import random
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse

# DuckDuckGo search
from duckduckgo_search import DDGS

# ---------------------------
# CONFIGURATION
# ---------------------------

SEARCH_KEYWORDS = [
    "construction companies in Nigeria website contact",
    "civil engineering firms Nigeria email",
    "agriculture companies Nigeria contact email",
    "agritech startups Nigeria website",
    "oil and gas companies Nigeria contact",
    "oilfield services Nigeria company website",
    "surveying companies Nigeria contact",
    "GIS mapping companies Nigeria website"
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}

MAX_RESULTS_PER_QUERY = 30
MAX_WORKERS = 10
REQUEST_DELAY = (1, 2)

OUTPUT_FILE = "nigeria_company_leads.xlsx"

# ---------------------------
# REGEX PATTERNS
# ---------------------------

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")
PHONE_REGEX = re.compile(r"(?:\+234|0)[789][01]\d{8}")

# ---------------------------
# HELPERS
# ---------------------------

def normalize_phone(phone):
    phone = re.sub(r"\D", "", phone)

    if phone.startswith("0"):
        return "+234" + phone[1:]
    elif phone.startswith("234"):
        return "+" + phone
    elif phone.startswith("+234"):
        return phone
    return None


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

    cleaned_phones = set()
    for p in phones:
        norm = normalize_phone(p)
        if norm:
            cleaned_phones.add(norm)

    return list(emails), list(cleaned_phones)


def extract_company_name(soup, url):
    if soup.title and soup.title.string:
        return soup.title.string.strip()
    return urlparse(url).netloc


def find_linkedin(soup):
    for a in soup.find_all("a", href=True):
        if "linkedin.com/company" in a["href"]:
            return a["href"]
    return ""


def extract_address(text):
    keywords = ["Lagos", "Abuja", "Port Harcourt", "Kano", "Nigeria"]

    for line in text.split("\n"):
        if any(k in line for k in keywords) and len(line) < 200:
            return line.strip()

    return ""


def is_valid_website(url):
    blocked = ["facebook.com", "twitter.com", "instagram.com", "youtube.com"]
    return not any(b in url for b in blocked)


# ---------------------------
# SEARCH FUNCTION (DuckDuckGo)
# ---------------------------

def search_companies(query):
    urls = []

    try:
        with DDGS() as ddgs:
            results = ddgs.text(query, max_results=MAX_RESULTS_PER_QUERY)

            for r in results:
                url = r.get("href")
                if url and is_valid_website(url):
                    urls.append(url)

    except Exception as e:
        print(f"Search error: {e}")

    return urls


# ---------------------------
# SCRAPE WEBSITE
# ---------------------------

def scrape_website(url, industry):
    html = fetch_page(url)

    if not html:
        return None

    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text(" ", strip=True)

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
# MAIN SCRAPER
# ---------------------------

def run_scraper():
    all_results = []
    seen_websites = set()

    print("🔍 Searching using DuckDuckGo...")

    search_results = []

    for keyword in SEARCH_KEYWORDS:
        urls = search_companies(keyword)

        for url in urls:
            if url not in seen_websites:
                seen_websites.add(url)
                search_results.append((url, keyword))

    print(f"🌐 Found {len(search_results)} unique websites.")

    print("🚀 Scraping websites...")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [
            executor.submit(scrape_website, url, industry)
            for url, industry in search_results
        ]

        for future in as_completed(futures):
            try:
                result = future.result()
                if result:
                    all_results.append(result)
            except:
                continue

    print(f"📊 Collected {len(all_results)} raw records.")

    if not all_results:
        print("❌ No data collected. Try reducing filters or checking internet.")
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
    print(f"📈 Final dataset size: {len(df)} records")


# ---------------------------
# ENTRY POINT
# ---------------------------

if __name__ == "__main__":
    run_scraper()