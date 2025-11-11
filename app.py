import os
import re
import uuid
import time
import random
from io import BytesIO
from datetime import datetime, timezone
from urllib.parse import urljoin
import pandas as pd
import requests
from bs4 import BeautifulSoup
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

# ---------------- Config ----------------
st.set_page_config(page_title="Amazon Top 20 Scraper", layout="wide")
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
]

# ---------------- Helpers ----------------
def utc_now_str():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

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

# ---------------- Amazon Scraping ----------------
def extract_asin_from_url(url):
    for pattern in [
        r"/dp/([A-Z0-9]{10})(?:[/?]|$)",
        r"/gp/product/([A-Z0-9]{10})(?:[/?]|$)",
    ]:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    return None

def parse_top20_from_category_page(url, session=None):
    session = session or _session()
    soup, final_url = get_soup(url, session)
    anchors = soup.find_all("a", href=True)
    seen, items = set(), []
    for a in anchors:
        href = a["href"]
        asin = extract_asin_from_url(href)
        if not asin or asin in seen:
            continue
        title = (a.get_text(strip=True) or "").strip()
        if not title:
            img = a.find("img", alt=True)
            if img and img.get("alt"):
                title = img["alt"].strip()
        if not title and a.get("title"):
            title = a["title"].strip()
        item_url = urljoin(final_url, href)
        seen.add(asin)
        items.append({"ASIN": asin, "Title": title or "", "URL": item_url})
        if len(items) >= 20:
            break
    return items

def extract_price_from_soup(soup):
    for el in soup.select("span.a-offscreen"):
        txt = el.get_text(strip=True)
        if txt.startswith("$"):
            return txt
    return ""

def extract_image_from_soup(soup):
    og = soup.find("meta", attrs={"property": "og:image"})
    if og and og.get("content"):
        return og["content"]
    img = soup.find("img", id="landingImage")
    if img:
        for k in ("src", "data-old-hires"):
            if img.get(k):
                return img[k]
    return ""

def fetch_item_details(url, session=None):
    session = session or _session()
    try:
        soup, _ = get_soup(url, session)
        return extract_price_from_soup(soup), extract_image_from_soup(soup)
    except Exception:
        return "", ""

# ---------------- Image Handling ----------------
def download_and_resize_image(url, max_px=180):
    try:
        r = requests.get(url, headers={"User-Agent": random.choice(USER_AGENTS)}, timeout=10)
        r.raise_for_status()
        img = PILImage.open(BytesIO(r.content))
        img.thumbnail((max_px, max_px), PILImage.LANCZOS)
        bio = BytesIO()
        img.save(bio, format="PNG")
        bio.seek(0)
        return bio
    except Exception:
        return None

# ---------------- Excel Export ----------------
def build_excel(df):
    wb = Workbook()
    ws_top = wb.active
    ws_top.title = "Top 20"
    ws_data = wb.create_sheet("Data")

    # Data sheet columns (new order)
    cols = [
        "MC SKU", "MC Title", "MC Retail", "MC Cost", "1-4 Avg",
        "Attributes", "Notes", "ASIN", "Amazon Title", "Amazon Price", "Amazon URL", "Image URL"
    ]
    for i, c in enumerate(cols, start=1):
        ws_data.cell(row=1, column=i, value=c)
    for i, row in enumerate(df.itertuples(), start=2):
        ws_data.cell(row=i, column=8, value=row.ASIN)
        ws_data.cell(row=i, column=9, value=row.Title)
        ws_data.cell(row=i, column=10, value=row.Price)
        ws_data.cell(row=i, column=11, value=row.URL)
        ws_data.cell(row=i, column=12, value=row.Image)
    for j, w in enumerate([14, 50, 14, 12, 12, 20, 20, 12, 60, 12, 40, 50], start=1):
        ws_data.column_dimensions[get_column_letter(j)].width = w

    # Top 20 layout
    items_per_row = 5
    block_rows = 10
    max_px = 160
    thin = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    align_left = Alignment(horizontal="left", vertical="top", wrap_text=True)

    for c in range(1, items_per_row + 1):
        ws_top.column_dimensions[get_column_letter(c)].width = 28

    for idx, row in enumerate(df.itertuples(), start=0):
        group = idx // items_per_row
        col = 1 + (idx % items_per_row)
        base = 1 + group * block_rows

        img_buf = download_and_resize_image(row.Image, max_px)
        if img_buf:
            xl_img = XLImage(img_buf)
            xl_img.anchor = f"{get_column_letter(col)}{base}"
            ws_top.add_image(xl_img)
            ws_top.row_dimensions[base].height = max_px * 0.75  # dynamic fit

        ws_top.cell(row=base + 1, column=col, value=f"Rank #{idx+1}").alignment = align_left
        ws_top.cell(row=base + 2, column=col, value=row.Title).alignment = align_left
        ws_top.cell(row=base + 3, column=col, value=row.Price).alignment = align_left
        ws_top.cell(row=base + 4, column=col, value="MC SKU:").alignment = align_left
        ws_top.cell(row=base + 5, column=col, value="MC Title:").alignment = align_left
        ws_top.cell(row=base + 6, column=col, value="MC Retail:").alignment = align_left
        ws_top.cell(row=base + 7, column=col, value="MC Cost:").alignment = align_left
        ws_top.cell(row=base + 8, column=col, value="1-4 Avg:").alignment = align_left
        ws_top.cell(row=base + 9, column=col, value="Attributes:").alignment = align_left
        ws_top.cell(row=base + 10, column=col, value="Notes:").alignment = align_left

        for r in range(base, base + block_rows + 1):
            ws_top.cell(row=r, column=col).border = border

    return wb

# ---------------- Streamlit UI ----------------
st.title("🧭 Amazon Top 20 Scraper")

url = st.text_input("Amazon Best Sellers URL", placeholder="https://www.amazon.com/gp/bestsellers/...")
delay = st.slider("Delay between requests (seconds)", 0.5, 3.0, 1.0)

if st.button("Fetch Top 20", type="primary"):
    if not url.startswith("http"):
        st.error("Please enter a valid Amazon URL starting with https://")
    else:
        session = _session()
        st.info("Fetching top 20 items from Amazon...")
        items = parse_top20_from_category_page(url, session)
        rows = []
        for i, item in enumerate(items):
            p, img = fetch_item_details(item["URL"], session)
            rows.append({
                "Rank": i + 1,
                "ASIN": item["ASIN"],
                "Title": item["Title"],
                "Price": p,
                "URL": item["URL"],
                "Image": img,
            })
            time.sleep(delay)
        df = pd.DataFrame(rows)
        st.session_state.results = df
        st.success(f"Fetched {len(df)} items!")

if "results" in st.session_state and st.session_state.results is not None:
    df = st.session_state.results
    for i, row in df.iterrows():
        c1, c2 = st.columns([1, 3])
        with c1:
            if row["Image"]:
                try:
                    img_data = requests.get(row["Image"], timeout=10).content
                    st.image(img_data, width=120)
                except Exception:
                    st.write("🖼️")
        with c2:
            st.markdown(f"**#{i+1}: {row['Title']}**")
            st.write(row["Price"])
            st.caption(row["URL"])
        st.divider()

    if st.button("Export to Excel"):
        wb = build_excel(df)
        fn = f"Amazon_Top20_{uuid.uuid4().hex[:6]}.xlsx"
        wb.save(fn)
        with open(fn, "rb") as f:
            st.download_button("⬇️ Download Spreadsheet", f, file_name=fn)
