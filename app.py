#######################################################################
# FULL APP.PY — Amazon Top-20 Scraper (with Image Proxy)
#######################################################################

import os
import re
import json
import time
import uuid
import random
from io import BytesIO
from datetime import datetime, timezone
from urllib.parse import urljoin
from typing import Optional

import requests
import pandas as pd
from bs4 import BeautifulSoup
import streamlit as st

from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------- STREAMLIT CONFIG ----------------------
st.set_page_config(page_title="Amazon Top-20 Scraper", layout="wide")

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_1) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.3 Safari/605.1.15",
]

SAVED_DIR = ".saved_searches"
os.makedirs(SAVED_DIR, exist_ok=True)

#######################################################################
# IMAGE PROXY SERVER (FastAPI Mounted in Thread)
#######################################################################
from fastapi import FastAPI
from fastapi.responses import Response
import uvicorn
import threading

if "api_started" not in st.session_state:
    api = FastAPI()

    @api.get("/img")
    def proxy_image(url: str):
        """
        Server-side download → convert → return stable PNG.
        Completely stops Streamlit image drops.
        """
        try:
            r = requests.get(url, headers={"User-Agent": random.choice(USER_AGENTS)}, timeout=10)
            r.raise_for_status()
            img = PILImage.open(BytesIO(r.content)).convert("RGB")
            buf = BytesIO()
            img.save(buf, format="PNG")
            return Response(buf.getvalue(), media_type="image/png")
        except Exception:
            return Response(status_code=404)

    def _run_proxy():
        uvicorn.run(api, host="0.0.0.0", port=9000, log_level="error")

    threading.Thread(target=_run_proxy, daemon=True).start()
    st.session_state.api_started = True

#######################################################################
# HELPERS
#######################################################################
def _session():
    s = requests.Session()
    s.headers.update({"Accept-Language": "en-US,en;q=0.9"})
    return s

def utc_now():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

def get_soup(url, s=None, timeout=15):
    s = s or _session()
    r = s.get(url, headers={"User-Agent": random.choice(USER_AGENTS)}, timeout=timeout)
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser")

#######################################################################
# AMAZON SCRAPING
#######################################################################
def extract_asin(url):
    for pat in [r"/dp/([A-Z0-9]{10})", r"/gp/product/([A-Z0-9]{10})"]:
        m = re.search(pat, url)
        if m:
            return m.group(1)

def parse_top20(url, s=None):
    s = s or _session()
    soup = get_soup(url, s)
    items = []
    seen = set()

    for a in soup.find_all("a", href=True):
        asin = extract_asin(a["href"])
        if not asin or asin in seen:
            continue

        title = a.get_text(strip=True)
        if not title:
            img = a.find("img", alt=True)
            if img:
                title = img.get("alt", "").strip()

        items.append({
            "ASIN": asin,
            "Title": title,
            "URL": urljoin(url, a["href"]),
        })
        seen.add(asin)

        if len(items) >= 20:
            break

    return items

def extract_price(soup):
    for sp in soup.select("span.a-offscreen"):
        t = sp.get_text(strip=True)
        if t.startswith("$"):
            return t
    return ""

def extract_sell(soup):
    # “9K+ bought in past month”
    pat = re.compile(r"(\d[\dK\+\,\.]*\s*\+?\s*bought.*past\s+month)", re.I)

    for el in soup.find_all(["div", "span"], string=pat):
        txt = el.get_text(strip=True)
        if txt:
            return txt
    return ""

def extract_image(soup):
    og = soup.find("meta", {"property": "og:image"})
    if og:
        return og.get("content", "")
    img = soup.find("img", id="landingImage")
    if img:
        return img.get("data-old-hires") or img.get("src", "")
    return ""

def fetch_details(url, s=None, retries=2, delay=(1.0, 2.0)):
    s = s or _session()
    for attempt in range(retries + 1):
        try:
            soup = get_soup(url, s)
            price = extract_price(soup)
            img = extract_image(soup)
            sell = extract_sell(soup)
            return price, img, sell
        except Exception:
            if attempt >= retries:
                return "", "", ""
            time.sleep(random.uniform(*delay))
    return "", "", ""

#######################################################################
# SAVED SEARCHES
#######################################################################
def run_path(rid): return os.path.join(SAVED_DIR, rid)
def meta_path(rid): return os.path.join(run_path(rid), "meta.json")
def data_path(rid): return os.path.join(run_path(rid), "data.csv")

def list_saved():
    out = []
    for rid in os.listdir(SAVED_DIR):
        try:
            mp, dp = meta_path(rid), data_path(rid)
            if os.path.exists(mp) and os.path.exists(dp):
                meta = json.load(open(mp))
                out.append({"id": rid, **meta})
        except:
            pass
    return sorted(out, key=lambda r: r.get("updated",""), reverse=True)

def save_run(rid, df, meta):
    os.makedirs(run_path(rid), exist_ok=True)
    df.to_csv(data_path(rid), index=False)
    meta["updated"] = utc_now()
    json.dump(meta, open(meta_path(rid), "w"), indent=2)

def load_run(rid):
    df = pd.read_csv(data_path(rid))
    meta = json.load(open(meta_path(rid)))
    return df, meta

#######################################################################
# SESSION STATE
#######################################################################
if "df" not in st.session_state: st.session_state.df = None
if "meta" not in st.session_state: st.session_state.meta = None
if "delay_min" not in st.session_state: st.session_state.delay_min = 1.0
if "delay_max" not in st.session_state: st.session_state.delay_max = 2.0

#######################################################################
# SIDEBAR — Clean Saved Searches
#######################################################################
st.sidebar.header("Saved Searches")

saved = list_saved()
options = ["(none)"] + [f"{r['name']} — {r.get('updated','')}" for r in saved]
sel = st.sidebar.selectbox("Select saved:", options)

col_l, col_d = st.sidebar.columns(2)

if col_l.button("Load"):
    if sel != "(none)":
        rid = saved[options.index(sel)-1]["id"]
        df, meta = load_run(rid)
        st.session_state.df = df
        st.session_state.meta = meta
        st.success(f"Loaded {meta['name']}")

if col_d.button("Delete"):
    if sel != "(none)":
        rid = saved[options.index(sel)-1]["id"]
        import shutil
        shutil.rmtree(run_path(rid), ignore_errors=True)
        st.warning("Deleted. Refreshing...")
        st.rerun()

#######################################################################
# TOP TOOLBAR
#######################################################################
st.title("Amazon Top-20 Scraper")

tb = st.container()
with tb:
    c1, c2, c3, c4 = st.columns([1.2, 1.4, 1.7, 2.7])
    fetch_btn = c1.button("Fetch Top 20", type="primary")
    download_placeholder = c2.empty()

    view_mode = c3.radio("View", ["List", "Grid", "Compact"], horizontal=True)

    with c4.expander("⚙️ Settings"):
        st.session_state.delay_min = st.slider("Min Delay", 0.3, 5.0, st.session_state.delay_min)
        st.session_state.delay_max = st.slider("Max Delay", 0.3, 6.0, st.session_state.delay_max)

#######################################################################
# INPUTS
#######################################################################
url = st.text_input("Amazon Best Sellers URL")
name = st.text_input("Category Name")

#######################################################################
# FETCH TOP 20
#######################################################################
if fetch_btn:
    if not url.startswith("http"):
        st.error("Please enter a valid Amazon URL.")
    else:
        s = _session()
        items = parse_top20(url, s)

        out = []
        prog = st.progress(0)
        stat = st.empty()

        for i, it in enumerate(items, start=1):
            stat.write(f"Loading item {i}...")
            price, img, sell = fetch_details(
                it["URL"], s,
                delay=(st.session_state.delay_min, st.session_state.delay_max)
            )
            out.append({
                "Rank": i,
                "Title": it["Title"],
                "ASIN": it["ASIN"],
                "URL": it["URL"],
                "Price": price,
                "Image": img,
                "Sell": sell,
                # data sheet fields
                "MCSKU": "", "MCTitle": "", "MCRetail": "",
                "MCCost": "", "Avg1_4": "",
                "Attributes": "", "Notes": "",
            })
            prog.progress(i/20)
            time.sleep(random.uniform(st.session_state.delay_min, st.session_state.delay_max))

        stat.success("Done!")
        st.session_state.df = pd.DataFrame(out)
        st.session_state.meta = {
            "name": name or "Top 20",
            "url": url,
            "created": utc_now(),
            "updated": utc_now()
        }

#######################################################################
# RENDER RESULTS
#######################################################################
def proxy(url):
    if not url:
        return ""
    return f"http://localhost:9000/img?url={url}"  # always stable

# --- List View ---
def render_list(df):
    for _, r in df.iterrows():
        cols = st.columns([1, 4])
        with cols[0]:
            st.markdown(f"<img src='{proxy(r['Image'])}' width='100'/>", unsafe_allow_html=True)
        with cols[1]:
            st.markdown(f"<b>#{r['Rank']}</b> — {r['Title']}")
            if r["Price"]:
                st.markdown(f"<span style='color:#B12704;font-weight:800'>{r['Price']}</span>", unsafe_allow_html=True)
            if r["Sell"]:
                st.caption(r["Sell"])
            st.markdown(f"[Amazon Link]({r['URL']})")
        st.markdown("<hr/>", unsafe_allow_html=True)

# --- Grid / Compact ---
GRID_CSS = """
<style>
.card{
  background:white;
  color:black;
  border:1px solid #e3e3e3;
  border-radius:8px;
  padding:8px;
  margin:5px;
  font-size:0.85rem;
  box-shadow:0 1px 3px rgba(0,0,0,0.05);
  min-height:330px;
}
.card:hover{ box-shadow:0 2px 8px rgba(0,0,0,0.15); }
.rank{
  background:#C45500;
  padding:2px 6px;
  color:white;
  font-weight:700;
  border-radius:4px;
  font-size:0.78rem;
}
.title{ font-weight:600; margin-top:4px; line-height:1.2rem; }
.price{ color:#B12704; font-weight:700; margin-top:6px; }
.sell{ font-size:0.78rem; color:#444; margin-top:2px; }
</style>
"""
st.markdown(GRID_CSS, unsafe_allow_html=True)

def render_grid(df, imgw):
    ncols = 5
    rows = [df.iloc[i:i+ncols] for i in range(0, len(df), ncols)]
    for row in rows:
        cols = st.columns(ncols)
        for c, (_, r) in zip(cols, row.iterrows()):
            html = "<div class='card'>"
            html += f"<div class='rank'>#{r['Rank']}</div>"
            html += f"<img src='{proxy(r['Image'])}' width='{imgw}'/>"
            html += f"<div class='title'>{r['Title']}</div>"
            if r["Price"]:
                html += f"<div class='price'>{r['Price']}</div>"
            if r["Sell"]:
                html += f"<div class='sell'>{r['Sell']}</div>"
            html += f"<div><a href='{r['URL']}' target='_blank'>Amazon Link</a></div>"
            html += "</div>"
            c.markdown(html, unsafe_allow_html=True)

if st.session_state.df is not None:
    df = st.session_state.df.head(20)

    if view_mode == "List":
        render_list(df)
    elif view_mode == "Grid":
        render_grid(df, imgw=110)
    else:
        render_grid(df, imgw=80)

#######################################################################
# EXCEL EXPORT
#######################################################################
def build_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Top 20"

    # Simple data dump
    cols = ["Rank","ASIN","URL","Image","Title","Price","Sell",
            "MCSKU","MCTitle","MCRetail","MCCost","Avg1_4","Attributes","Notes"
    ]
    for j,cname in enumerate(cols,1):
        ws.cell(1,j,cname)

    for i,r in enumerate(df.itertuples(index=False),2):
        ws.cell(i,1,r.Rank)
        ws.cell(i,2,r.ASIN)
        ws.cell(i,3,r.URL)
        ws.cell(i,4,r.Image)
        ws.cell(i,5,r.Title)
        ws.cell(i,6,r.Price)
        ws.cell(i,7,r.Sell)
        ws.cell(i,8,r.MCSKU)
        ws.cell(i,9,r.MCTitle)
        ws.cell(i,10,r.MCRetail)
        ws.cell(i,11,r.MCCost)
        ws.cell(i,12,r.Avg1_4)
        ws.cell(i,13,r.Attributes)
        ws.cell(i,14,r.Notes)

    mem = BytesIO()
    wb.save(mem)
    mem.seek(0)
    return mem.read()

if st.session_state.df is not None:
    download_placeholder.download_button(
        "Download Excel",
        data=build_excel(st.session_state.df.head(20)),
        file_name=f"{st.session_state.meta['name'].replace(' ','_')}_Top20.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
