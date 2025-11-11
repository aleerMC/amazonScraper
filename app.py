import re
import os
import uuid
import time
import random
from datetime import datetime, timezone
from urllib.parse import urljoin, quote_plus
import pandas as pd
import requests
from bs4 import BeautifulSoup
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ======================== SETUP ========================
st.set_page_config(
    page_title="Amazon ⇄ Micro Center Matcher",
    layout="wide",
    initial_sidebar_state="expanded",
)
SAVED_DIR = ".saved_searches"
os.makedirs(SAVED_DIR, exist_ok=True)

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
]

# ======================== HELPERS ========================
def _session():
    s = requests.Session()
    s.headers.update({"Accept-Language": "en-US,en;q=0.9"})
    return s

def get_soup(url, session=None, timeout=15):
    session = session or _session()
    headers = {"User-Agent": random.choice(USER_AGENTS)}
    r = session.get(url, headers=headers, timeout=timeout)
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser"), r.url

# ======================== AMAZON SCRAPER ========================
ASIN_REGEXES = [
    re.compile(r"/dp/([A-Z0-9]{10})(?:[/?]|$)"),
    re.compile(r"/gp/product/([A-Z0-9]{10})(?:[/?]|$)"),
    re.compile(r"[?&]ASIN=([A-Z0-9]{10})(?:&|$)", re.IGNORECASE),
]

def extract_asin_from_url(url: str):
    for rx in ASIN_REGEXES:
        m = rx.search(url)
        if m:
            return m.group(1)
    return None

def parse_top20_from_category_page(url, session=None):
    session = session or _session()
    soup, final_url = get_soup(url, session)
    anchors = soup.find_all("a", href=True)
    seen_asins, items = set(), []
    for a in anchors:
        href = a.get("href", "")
        asin = extract_asin_from_url(href)
        if not asin or asin in seen_asins:
            continue
        title = (a.get_text(strip=True) or "").strip()
        if not title:
            img = a.find("img", alt=True)
            if img and img.get("alt"):
                title = img["alt"].strip()
        if not title and a.get("title"):
            title = a["title"].strip()
        item_url = urljoin(final_url, href) if href.startswith("/") else href
        try:
            r = session.get(item_url, allow_redirects=True, timeout=10)
            item_url = r.url
        except Exception:
            pass
        seen_asins.add(asin)
        items.append({"ASIN": asin, "Title": title or "", "URL": item_url})
        if len(items) >= 20:
            break
    return items

def extract_price_from_soup_amzn(soup):
    for off in soup.select("span.a-offscreen"):
        val = off.get_text(strip=True)
        if re.match(r"^\$", val):
            return val
    return ""

def extract_image_from_soup_amzn(soup):
    og = soup.find("meta", attrs={"property": "og:image"})
    if og and og.get("content"):
        return og["content"]
    img = soup.find("img", id="landingImage")
    if img:
        for a in ("data-old-hires", "src"):
            if img.get(a):
                return img[a]
    return ""

def fetch_item_details_amzn(item_url, session=None, retries=2, delay_range=(1.5, 3.0)):
    session = session or _session()
    for attempt in range(retries + 1):
        try:
            soup, _ = get_soup(item_url, session, timeout=20)
            price = extract_price_from_soup_amzn(soup)
            image = extract_image_from_soup_amzn(soup)
            if image:
                return price, image
            time.sleep(random.uniform(*delay_range))
        except Exception:
            if attempt >= retries:
                return "", ""
            time.sleep(random.uniform(*delay_range))
    return "", ""

@st.cache_data(ttl=3600)
def load_image_bytes(url):
    try:
        if not url:
            return None
        resp = requests.get(url, headers={"User-Agent": random.choice(USER_AGENTS)}, timeout=10)
        resp.raise_for_status()
        return resp.content
    except Exception:
        return None

# ======================== MICRO CENTER SCRAPER ========================
def fetch_microcenter_info(sku):
    base_url = f"https://www.microcenter.com/search/search_results.aspx?Ntt={quote_plus(str(sku))}"
    try:
        soup, _ = get_soup(base_url)
        links = soup.select("a.ProductLink") or soup.select("a[href*='/product/']")
        for link in links:
            href = link.get("href", "")
            if "battery" in href.lower():
                continue
            full_url = urljoin(base_url, href)
            psoup, _ = get_soup(full_url)
            title = psoup.find("h1")
            price_el = psoup.find("span", {"itemprop": "price"})
            img_el = psoup.find("img", {"id": "productImage"})
            desc = psoup.find("div", {"class": "specs"})
            return {
                "Title": title.get_text(strip=True) if title else "",
                "Price": price_el.get_text(strip=True) if price_el else "",
                "Image": img_el["src"] if img_el else "",
                "Description": desc.get_text(strip=True) if desc else "",
                "URL": full_url,
            }
    except Exception:
        return {}
    return {}

# ======================== EXCEL EXPORT ========================
def build_excel(df_amzn):
    wb = Workbook()
    ws = wb.active
    ws.title = "Top 20"
    ws.append([
        "Amazon Image", "Amazon Title", "Amazon Price", "Amazon Link",
        "MC Image", "MC SKU", "MC Description", "MC Price",
        "Attributes", "Notes"
    ])
    ws.freeze_panes = "A2"
    for i, row in enumerate(df_amzn.itertuples(), start=2):
        band = "FFFFFF" if i % 2 else "F7F7F7"
        ws.append(["", row.Title, row.Price, row.URL, "", "", "", "", "", ""])
        for j in range(1, 11):
            ws.cell(i, j).fill = PatternFill("solid", fgColor=band)
            ws.cell(i, j).alignment = Alignment(wrap_text=True, vertical="top")
    for col in range(1, 11):
        ws.column_dimensions[get_column_letter(col)].width = 25
    return wb

# ======================== SIDEBAR SETTINGS ========================
with st.sidebar.expander("⚙️ Settings", expanded=False):
    st.markdown("### Display & Speed Settings")
    theme = st.radio("Theme", ["Light", "Dark", "Flannel"], horizontal=True)
    delay = st.slider("Scrape Delay (sec)", 0.5, 5.0, 1.5, 0.5)
    retries = st.slider("Retry Count", 0, 4, 2, 1)

# ======================== THEME CSS ========================
if theme == "Flannel":
    st.markdown(
        """
        <style>
        .stApp {
            background-color: #1e1e1e;
            position: relative;
            overflow: hidden;
        }
        /* translucent plaid overlay */
        .stApp::before {
            content: "";
            position: absolute;
            top: 0; left: 0;
            width: 100%; height: 100%;
            background-image: url('rhythmicRed.png');
            background-size: 400px auto;
            background-repeat: repeat;
            opacity: 0.25;
            z-index: -1;
        }
        section[data-testid="stSidebar"] > div:first-child {
            background: linear-gradient(180deg,#2c0000 0%,#1e1e1e 100%);
            border-right: 1px solid #440000;
        }
        h1,h2,h3,h4,h5,h6 { color: #ffdddd !important; }
        .stButton button {
            background: #660000; color:white;
            border:1px solid #992222; border-radius:8px;
        }
        .stButton button:hover { background:#992222; }
        </style>
        """,
        unsafe_allow_html=True,
    )
elif theme == "Dark":
    st.markdown("<style>body{background-color:#0E1117;color:#FAFAFA;}</style>", unsafe_allow_html=True)
else:
    st.markdown("<style>body{background-color:white;color:black;}</style>", unsafe_allow_html=True)

# ======================== MAIN APP ========================
st.title("🧭 Amazon → Micro Center Matcher")
st.write("Fetch Amazon Top-20 items and compare side-by-side with Micro Center SKUs.")

url = st.text_input("Amazon Best Seller Category URL")
if st.button("Fetch Top 20"):
    with st.spinner("Fetching Top 20 items..."):
        items = parse_top20_from_category_page(url)
        session = _session()
        data = []
        for item in items:
            price, img = fetch_item_details_amzn(
                item["URL"], session, retries=retries, delay_range=(delay, delay + 1.0)
            )
            data.append({
                "ASIN": item["ASIN"],
                "Title": item["Title"],
                "Price": price,
                "URL": item["URL"],
                "Image": img,
            })
        st.session_state["df_amzn"] = pd.DataFrame(data)
        st.success("✅ Amazon data fetched successfully!")

if "df_amzn" in st.session_state:
    df = st.session_state["df_amzn"]
    for i, row in df.iterrows():
        st.markdown(f"**#{i+1}. {row['Title']}** — {row['Price']}")
        img_data = load_image_bytes(row["Image"])
        if img_data:
            st.image(img_data, width=150)
        mc_sku = st.text_input(f"Micro Center SKU for {row['ASIN']}", key=f"sku_{i}")
        if mc_sku:
            info = fetch_microcenter_info(mc_sku)
            if info:
                st.image(load_image_bytes(info["Image"]), width=120)
                st.write(f"**{info['Title']}** — {info['Price']}")
                st.write(info["Description"])
        st.text_area("Attributes", key=f"attr_{i}")
        st.text_area("Notes", key=f"note_{i}")
        st.markdown("---")

    if st.button("Export Excel"):
        wb = build_excel(df)
        filename = f"top20_export_{uuid.uuid4().hex[:6]}.xlsx"
        wb.save(filename)
        with open(filename, "rb") as f:
            st.download_button("⬇️ Download Excel", f, file_name=filename)
