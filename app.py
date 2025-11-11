import re
import os
import json
import uuid
import time
import random
from io import BytesIO
from datetime import datetime, timezone
from urllib.parse import urljoin, quote_plus, urlparse
import pandas as pd
import requests
from bs4 import BeautifulSoup
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

# ---------------- Global Config ----------------
st.set_page_config(page_title="Amazon → Micro Center Matcher", layout="wide")

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
]

SAVED_DIR = ".saved_searches"
os.makedirs(SAVED_DIR, exist_ok=True)

# ---------------- Utility Functions ----------------
def _session():
    s = requests.Session()
    s.headers.update({"Accept-Language": "en-US,en;q=0.9"})
    return s

def utc_now_str():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

def get_soup(url, session=None, timeout=15):
    session = session or _session()
    headers = {"User-Agent": random.choice(USER_AGENTS)}
    r = session.get(url, headers=headers, timeout=timeout)
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser"), r.url

# ---------------- Amazon Scraping ----------------
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

def fetch_item_details_amzn(item_url, session=None, retries=2, delay_range=(1.2, 2.4)):
    session = session or _session()
    for attempt in range(retries + 1):
        try:
            soup, _ = get_soup(item_url, session, timeout=20)
            price = extract_price_from_soup_amzn(soup)
            image = extract_image_from_soup_amzn(soup)
            return price, image
        except Exception:
            if attempt >= retries:
                return "", ""
            time.sleep(random.uniform(*delay_range))
    return "", ""

# ---------------- Micro Center Scraping ----------------
@st.cache_data(ttl=1800)
def fetch_microcenter_candidates(q: str, limit: int = 8):
    q = (q or "").strip()
    if not q:
        return []
    session = _session()
    headers = {"User-Agent": random.choice(USER_AGENTS)}
    try:
        search_url = f"https://www.microcenter.com/search/search_results.aspx?Ntt={quote_plus(q)}"
        r = session.get(search_url, headers=headers, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        links = []
        for a in soup.select("a[href*='/product/']"):
            href = a["href"]
            if any(x in href.lower() for x in ["service", "battery", "repair"]):
                continue
            links.append("https://www.microcenter.com" + href)
        links = list(dict.fromkeys(links))
        results = []
        for l in links[:limit]:
            psoup, _ = get_soup(l, session)
            sku = ""
            sku_meta = psoup.find(attrs={"itemprop": "sku"})
            if sku_meta:
                sku = sku_meta.get("content") or sku_meta.get_text(strip=True)
            title = psoup.find("h1")
            price = psoup.find(attrs={"itemprop": "price"})
            img = psoup.find("meta", {"property": "og:image"})
            results.append({
                "MCSKU": sku,
                "MCTitle": title.get_text(strip=True) if title else "",
                "MCPrice": price.get("content") if price and price.get("content") else "",
                "MCImageURL": img["content"] if img and img.get("content") else "",
                "MCURL": l
            })
        return results
    except Exception:
        return []

# ---------------- Persistence ----------------
def _run_path(run_id): return os.path.join(SAVED_DIR, run_id)
def _meta_path(run_id): return os.path.join(_run_path(run_id), "meta.json")
def _data_path(run_id): return os.path.join(_run_path(run_id), "data.csv")

def list_saved_runs():
    runs = []
    for rid in os.listdir(SAVED_DIR):
        mp = _meta_path(rid)
        dp = _data_path(rid)
        if os.path.exists(mp) and os.path.exists(dp):
            try:
                with open(mp, "r", encoding="utf-8") as f: meta = json.load(f)
                runs.append({"id": rid, **meta})
            except Exception: pass
    runs.sort(key=lambda r: r.get("updated_at", ""), reverse=True)
    return runs

def save_run(run_id, df, meta):
    os.makedirs(_run_path(run_id), exist_ok=True)
    df.to_csv(_data_path(run_id), index=False, encoding="utf-8")
    meta["updated_at"] = utc_now_str()
    with open(_meta_path(run_id), "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2)

def load_run(run_id):
    df = pd.read_csv(_data_path(run_id), dtype=str).fillna("")
    with open(_meta_path(run_id), "r", encoding="utf-8") as f:
        meta = json.load(f)
    return df, meta

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("Saved Searches")

    saved = list_saved_runs()
    if saved:
        st.markdown("---")
        for s in saved:
            st.markdown(f"**{s['name']}**  \n🕓 *{s.get('updated_at','')}*  \n🔗 *{s.get('category_desc','')}*")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                if st.button("Load", key=f"load_{s['id']}"):
                    df, meta = load_run(s["id"])
                    st.session_state.results = df
                    st.session_state.current_run_id = s["id"]
                    st.session_state.current_meta = meta
                    st.rerun()
            with col2:
                if st.button("Save", key=f"save_{s['id']}"):
                    if st.session_state.results is not None:
                        save_run(s["id"], st.session_state.results, st.session_state.current_meta)
                        st.success("Saved.")
            with col3:
                if st.button("Rename", key=f"rename_{s['id']}"):
                    new_name = st.text_input("New Name", s["name"], key=f"name_{s['id']}")
                    if new_name.strip():
                        df, meta = load_run(s["id"])
                        meta["name"] = new_name.strip()
                        save_run(s["id"], df, meta)
                        st.success("Renamed.")
            with col4:
                if st.button("Delete", key=f"delete_{s['id']}"):
                    import shutil
                    shutil.rmtree(_run_path(s["id"]), ignore_errors=True)
                    st.warning(f"Deleted {s['name']}")
                    st.rerun()
            st.markdown("---")
    else:
        st.caption("No saved searches yet.")

    with st.expander("⚙️ Settings", expanded=False):
        theme = st.radio("Theme", ["Light", "Dark"], horizontal=True)
        delay = st.slider("Scrape Delay (sec)", 0.5, 5.0, 1.5, 0.5)

# ---------------- Theme Styling ----------------
if theme == "Dark":
    st.markdown("<style>body{background-color:#0E1117;color:#FAFAFA;}</style>", unsafe_allow_html=True)
else:
    st.markdown("<style>body{background-color:#FFFFFF;color:#000000;}</style>", unsafe_allow_html=True)

# ---------------- Main App UI ----------------
st.title("🧭 Amazon → Micro Center Matcher")

col_in1, col_in2 = st.columns([3,2])
with col_in1:
    amz_url = st.text_input("Amazon Best Sellers URL", placeholder="https://www.amazon.com/gp/bestsellers/pc/...")
with col_in2:
    category_desc = st.text_input("Category Description", placeholder="e.g., Single Board Computers")

if st.button("Fetch Top 20", type="primary"):
    session = _session()
    st.info("Fetching top 20 items from Amazon...")
    items = parse_top20_from_category_page(amz_url.strip(), session)
    rows = []
    for i, item in enumerate(items):
        p, img = fetch_item_details_amzn(item["URL"], session)
        rows.append({
            "Rank": i+1, "ASIN": item["ASIN"], "Title": item["Title"], "URL": item["URL"],
            "Price": p, "Image": img, "Notes": "", "AttrMatch": "", "Category": category_desc
        })
        time.sleep(delay)
    st.session_state.results = pd.DataFrame(rows)
    st.session_state.current_meta = {"name": category_desc, "updated_at": utc_now_str()}
    st.success("Top 20 fetched!")

# ---------------- Display ----------------
if "results" in st.session_state and st.session_state.results is not None:
    df = st.session_state.results
    for i, r in df.iterrows():
        c1, c2, c3 = st.columns([1,1,1])
        with c1:
            st.image(r["Image"], width=120)
            st.markdown(f"**#{r['Rank']} — {r['Title']}**  \n{r['Price']}")
        with c2:
            sku = st.text_input("MC SKU", key=f"sku_{i}")
            if sku:
                mc = fetch_microcenter_candidates(sku, 3)
                if mc:
                    st.image(mc[0]["MCImageURL"], width=100)
                    st.write(mc[0]["MCTitle"])
                    st.write(f"Price: {mc[0]['MCPrice']}")
        with c3:
            df.at[i, "AttrMatch"] = st.text_input("Attributes", value=r["AttrMatch"], key=f"attr_{i}")
            df.at[i, "Notes"] = st.text_input("Notes", value=r["Notes"], key=f"note_{i}")
        st.markdown("---")

    if st.button("Save Search"):
        run_id = uuid.uuid4().hex[:12]
        save_run(run_id, df, {"name": category_desc, "updated_at": utc_now_str()})
        st.success("Saved!")

    if st.button("Export Excel"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Top 20"
        ws.append(df.columns.tolist())
        for _, row in df.iterrows():
            ws.append(row.tolist())
        filename = f"Top20_{uuid.uuid4().hex[:6]}.xlsx"
        wb.save(filename)
        with open(filename, "rb") as f:
            st.download_button("⬇️ Download", f, file_name=filename)
