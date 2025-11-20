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

# ========================= App Setup =========================
st.set_page_config(page_title="Amazon Top-20 → Excel Exporter", layout="wide")

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_0) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.2 Safari/605.1.15",
]

SAVED_DIR = ".saved_searches"
os.makedirs(SAVED_DIR, exist_ok=True)

# Excel layout knobs
EXCEL_COL_WIDTH = 24
EXCEL_IMG_MAX_PX = 200
EXCEL_IMG_ROW_HEIGHT = 82
RANK_FILL = PatternFill("solid", fgColor="C45500")
WHITE_FONT = Font(color="FFFFFF", bold=True)
THIN_SIDE = Side(style="thin", color="DDDDDD")
BORDER_BOX = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

CARD_CSS = """
<style>
.card{
  background:#ffffff;
  color:#000000;
  border:1px solid #e0e0e0;
  border-radius:8px;
  padding:8px;
  margin:3px;
  box-shadow:0 1px 3px rgba(0,0,0,0.05);
  font-size:0.85rem;
}
.card:hover{ box-shadow:0 2px 6px rgba(0,0,0,0.12); }
.title{
  font-weight:600;
  font-size:0.90rem;
  line-height:1.2rem;
  margin:4px 0 2px 0;
}
.price{ color:#B12704; font-weight:700; margin:2px 0; }
.sell{ font-size:0.78rem; color:#444; margin:2px 0; }
.rankbar{
  background:#C45500;
  color:#fff;
  padding:1px 6px;
  border-radius:6px;
  display:inline-block;
  font-weight:800;
  font-size:0.78rem;
  margin-bottom:4px;
}
.meta a{ text-decoration:none; font-size:0.80rem; }
.toolbar .stButton>button { height:40px; padding:0 10px; }
</style>
"""

# ========================= Helpers =========================
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

# ========================= Amazon scraping =========================
def extract_asin(url):
    for pattern in [r"/dp/([A-Z0-9]{10})", r"/gp/product/([A-Z0-9]{10})", r"[?&]ASIN=([A-Z0-9]{10})"]:
        m = re.search(pattern, url, re.I)
        if m:
            return m.group(1)

def parse_top20(url, s=None):
    s = s or _session()
    soup = get_soup(url, s)
    items, seen = [], set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        asin = extract_asin(href or "")
        if not asin or asin in seen:
            continue
        title = (a.get_text(strip=True) or "").strip()
        if not title:
            img = a.find("img", alt=True)
            if img and img.get("alt"):
                title = img["alt"].strip()
        if not title and a.get("title"):
            title = a.get("title", "").strip()
        items.append({
            "ASIN": asin,
            "Title": title,
            "URL": urljoin(url, href),
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

def extract_sellthrough(soup):
    # Examples: "9K+ bought in past month", "50+ bought in past month"
    pat = re.compile(r"\b\d[\d,\.K\+]*\s*\+?\s*bought.*past\s+month", re.I)
    for el in soup.find_all(["div", "span"], string=pat):
        txt = el.get_text(" ", strip=True)
        if txt:
            return txt
    # Fallback: any "bought in past month"
    pat2 = re.compile(r"bought.*past\s+month", re.I)
    for el in soup.find_all(["div", "span"], string=pat2):
        txt = el.get_text(" ", strip=True)
        if re.search(r"\d", txt):
            return txt
    return ""

def extract_image(soup):
    og = soup.find("meta", {"property": "og:image"})
    if og and og.get("content"):
        return og["content"]
    img = soup.find("img", id="landingImage")
    if img:
        return img.get("data-old-hires") or img.get("src", "")
    return ""

def fetch_details(url, s=None, retries=2, delay=(1.0, 2.0)):
    s = s or _session()
    for attempt in range(retries + 1):
        try:
            soup = get_soup(url, s, timeout=20)
            return extract_price(soup), extract_image(soup), extract_sellthrough(soup)
        except Exception:
            if attempt >= retries:
                return "", "", ""
            time.sleep(random.uniform(*delay))
    return "", "", ""

# ========================= Persistence =========================
def _run_path(run_id): return os.path.join(SAVED_DIR, run_id)
def _meta_path(run_id): return os.path.join(_run_path(run_id), "meta.json")
def _data_path(run_id): return os.path.join(_run_path(run_id), "data.csv")

def list_saved():
    runs = []
    for rid in os.listdir(SAVED_DIR):
        try:
            mp, dp = _meta_path(rid), _data_path(rid)
            if os.path.exists(mp) and os.path.exists(dp):
                meta = json.load(open(mp, "r"))
                runs.append({"id": rid, **meta})
        except Exception:
            pass
    runs.sort(key=lambda r: r.get("updated",""), reverse=True)
    return runs

def save_run(run_id, df, meta):
    os.makedirs(_run_path(run_id), exist_ok=True)
    df.to_csv(_data_path(run_id), index=False)
    meta["updated"] = utc_now()
    json.dump(meta, open(_meta_path(run_id), "w"), indent=2)

def load_run(run_id):
    df = pd.read_csv(_data_path(run_id))
    meta = json.load(open(_meta_path(run_id), "r"))
    return df, meta

def new_meta(name, url):
    now = utc_now()
    return {"name": name or "Top 20", "url": url or "", "created": now, "updated": now}

# ========================= Session State =========================
if "df" not in st.session_state: st.session_state.df = None
if "meta" not in st.session_state: st.session_state.meta = None
if "delay_min" not in st.session_state: st.session_state.delay_min = 1.0
if "delay_max" not in st.session_state: st.session_state.delay_max = 2.0
if "debug" not in st.session_state: st.session_state.debug = False

# ========================= Sidebar (Saved Searches only) =========================
st.sidebar.header("Saved Searches")
saved = list_saved()
choices = ["(none)"] + [f"{r['name']} — {r.get('updated','')}" for r in saved]
sel = st.sidebar.selectbox("Select a saved search", choices, index=0)

col_sb1, col_sb2 = st.sidebar.columns(2)
if col_sb1.button("Load"):
    if sel != "(none)":
        ridx = choices.index(sel) - 1
        run = saved[ridx]
        df, meta = load_run(run["id"])
        st.session_state.df = df
        st.session_state.meta = meta
        st.success(f"Loaded “{meta.get('name','Top 20')}”")
if col_sb2.button("Delete"):
    if sel != "(none)":
        ridx = choices.index(sel) - 1
        run = saved[ridx]
        import shutil
        shutil.rmtree(_run_path(run["id"]), ignore_errors=True)
        st.warning("Deleted. Refreshing list…")
        st.rerun()

st.sidebar.caption("Dropdown + Load/Delete only, to keep things clean.")

# ========================= Top Toolbar =========================
st.markdown(CARD_CSS, unsafe_allow_html=True)
toolbar = st.container()
with toolbar:
    c1, c2, c3, c4 = st.columns([1.2, 1.4, 1.7, 2.7])

    fetch_clicked = c1.button("Fetch Top 20", type="primary")
    download_placeholder = c2.empty()

    view_mode = c3.radio(
        "View Mode",
        ["List", "Grid", "Compact"],
        horizontal=True,
        label_visibility="visible"
    )

    with c4.expander("⚙️ Settings", expanded=False):
        st.session_state.delay_min = st.slider(
            "Min per-item delay (sec)", 0.3, 5.0,
            st.session_state.delay_min, 0.1,
            key="delay_min_slider"
        )
        st.session_state.delay_max = st.slider(
            "Max per-item delay (sec)", 0.4, 6.0,
            st.session_state.delay_max, 0.1,
            key="delay_max_slider"
        )
        st.session_state.debug = st.checkbox(
            "Debug (log basic events)",
            value=st.session_state.debug
        )

# ========================= Inputs (below toolbar) =========================
url = st.text_input(
    "Amazon Best Sellers URL",
    placeholder="https://www.amazon.com/gp/bestsellers/pc/17441247011"
)
name = st.text_input(
    "Category Name (for saving/export)",
    placeholder="e.g., Single Board Computers"
)

# ========================= Fetch Handler =========================
if fetch_clicked:
    if not re.match(r"^https?://", (url or "").strip(), re.I):
        st.error("Please enter a valid URL starting with http:// or https://")
    else:
        s = _session()
        try:
            items = parse_top20(url.strip(), s)
        except Exception as e:
            st.error(f"Failed to fetch list: {e}")
            items = []

        out = []
        status = st.empty()
        progress = st.progress(0) if items else None

        for i, it in enumerate(items, start=1):
            status.write(f"Fetching item {i} of {len(items)}…")
            price, image, sell = fetch_details(
                it["URL"],
                s,
                delay=(st.session_state.delay_min, st.session_state.delay_max)
            )
            out.append({
                "Rank": i,
                "Title": it["Title"],
                "ASIN": it["ASIN"],
                "URL": it["URL"],
                "Price": price,
                "Image": image,
                "Sell": sell,
                # Fields for Data sheet (MC comparison later)
                "MCSKU": "", "MCTitle": "", "MCRetail": "",
                "MCCost": "", "Avg1_4": "", "Attributes": "", "Notes": "",
            })
            if progress:
                progress.progress(int(i / max(1, len(items)) * 100))
            time.sleep(random.uniform(st.session_state.delay_min, st.session_state.delay_max))

        if progress:
            progress.empty()
        status.success("Top 20 fetched!")
        st.session_state.df = pd.DataFrame(out)
        st.session_state.meta = new_meta(name or "Top 20", url.strip())

# ========================= Display Results (no re-scrape on view change) =========================
def render_list(df: pd.DataFrame):
    for _, r in df.iterrows():
        row = st.container()
        with row:
            cols = st.columns([1, 4])
            img_html = ""
            if isinstance(r["Image"], str) and r["Image"]:
                img_html = f"<img src='{r['Image']}' width='90'/>"
            with cols[0]:
                if img_html:
                    st.markdown(img_html, unsafe_allow_html=True)
                else:
                    st.write("—")
            with cols[1]:
                st.markdown(
                    f"<span class='rankbar'>#{r['Rank']}</span>",
                    unsafe_allow_html=True
                )
                st.markdown(
                    f"<div class='title'>{r['Title']}</div>",
                    unsafe_allow_html=True
                )
                if r["Price"]:
                    st.markdown(
                        f"<div class='price'>{r['Price']}</div>",
                        unsafe_allow_html=True
                    )
                if r["Sell"]:
                    st.markdown(
                        f"<div class='sell'>{r['Sell']}</div>",
                        unsafe_allow_html=True
                    )
                st.markdown(
                    f"<a href='{r['URL']}' target='_blank'>Open on Amazon</a>",
                    unsafe_allow_html=True
                )
        st.markdown("<hr/>", unsafe_allow_html=True)

def render_cards(df: pd.DataFrame, imgw: int):
    ncols = 5
    rows = [df.iloc[i:i + ncols] for i in range(0, len(df), ncols)]
    for row in rows:
        cols = st.columns(ncols)
        for c, (_, r) in zip(cols, row.iterrows()):
            html = "<div class='card'>"
            html += f"<div class='rankbar'>#{r['Rank']}</div>"
            if isinstance(r["Image"], str) and r["Image"]:
                html += f"<br><img src='{r['Image']}' width='{imgw}'/>"
            html += f"<div class='title'>{r['Title']}</div>"
            if r["Price"]:
                html += f"<div class='price'>{r['Price']}</div>"
            if r["Sell"]:
                html += f"<div class='sell'>{r['Sell']}</div>"
            html += f"<div class='meta'><a href='{r['URL']}' target='_blank'>Open on Amazon</a></div>"
            html += "</div>"
            c.markdown(html, unsafe_allow_html=True)
        st.write("")

if st.session_state.df is not None:
    df = st.session_state.df.copy()
    # ensure we never show more than 20
    df = df.head(20)

    if view_mode == "List":
        render_list(df)
    elif view_mode == "Grid":
        render_cards(df, imgw=110)
    else:  # Compact
        render_cards(df, imgw=80)

# ========================= Excel Export =========================
def _download_image_bytes(url: str, max_px: int = EXCEL_IMG_MAX_PX) -> Optional[BytesIO]:
    try:
        if not url or not isinstance(url, str) or not url.lower().startswith("http"):
            return None
        resp = requests.get(url, headers={"User-Agent": random.choice(USER_AGENTS)}, timeout=15)
        resp.raise_for_status()
        img = PILImage.open(BytesIO(resp.content)).convert("RGBA")
        img.thumbnail((max_px, max_px), PILImage.LANCZOS)
        buff = BytesIO()
        img.save(buff, format="PNG")
        buff.seek(0)
        return buff
    except Exception:
        return None

def build_xlsx(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws_top = wb.active
    ws_top.title = "Top 20"
    ws_data = wb.create_sheet("Data")

    # Data columns fixed
    cols = [
        "Rank","ASIN","URL","Image","Title","Price","Sell",
        "MCSKU","MCTitle","MCRetail","MCCost","Avg1_4","Attributes","Notes"
    ]
    for j, c in enumerate(cols, 1):
        ws_data.cell(row=1, column=j, value=c)
    for i, r in enumerate(df.itertuples(index=False), 2):
        ws_data.cell(i, 1, r.Rank)
        ws_data.cell(i, 2, r.ASIN)
        ws_data.cell(i, 3, r.URL)
        ws_data.cell(i, 4, r.Image)
        ws_data.cell(i, 5, r.Title)
        ws_data.cell(i, 6, r.Price)
        ws_data.cell(i, 7, r.Sell)
        ws_data.cell(i, 8, r.MCSKU)
        ws_data.cell(i, 9, r.MCTitle)
        ws_data.cell(i,10, r.MCRetail)
        ws_data.cell(i,11, r.MCCost)
        ws_data.cell(i,12, r.Avg1_4)
        ws_data.cell(i,13, r.Attributes)
        ws_data.cell(i,14, r.Notes)

    for j in range(1, len(cols)+1):
        ws_data.column_dimensions[get_column_letter(j)].width = 22
    ws_data.freeze_panes = "A2"

    # Top sheet layout
    for c in range(1, 6):  # 5 columns
        ws_top.column_dimensions[get_column_letter(c)].width = EXCEL_COL_WIDTH

    wrap_top = Alignment(wrap_text=True, vertical="top")
    center_mid = Alignment(horizontal="center", vertical="center")
    ITEMS_PER_ROW = 5
    ROW_BLOCK = 11  # rank + image + 9 info rows

    def data_ref(col_idx, drow):
        return f"Data!{get_column_letter(col_idx)}{drow}"

    max_items = min(20, len(df))
    for idx in range(max_items):
        group = idx // ITEMS_PER_ROW
        col = (idx % ITEMS_PER_ROW) + 1
        base = group * ROW_BLOCK + 1
        drow = idx + 2  # corresponding row in Data

        # Rank bar
        rank_cell = ws_top.cell(row=base, column=col, value=f"Rank #{df.at[idx, 'Rank']}")
        rank_cell.fill = RANK_FILL
        rank_cell.font = WHITE_FONT
        rank_cell.alignment = center_mid
        rank_cell.border = BORDER_BOX

        # Image row (in-cell placement)
        ws_top.row_dimensions[base + 1].height = EXCEL_IMG_ROW_HEIGHT
        img_buf = _download_image_bytes(df.at[idx, "Image"], EXCEL_IMG_MAX_PX)
        if img_buf:
            xl = XLImage(img_buf)
            xl.anchor = f"{get_column_letter(col)}{base + 1}"
            ws_top.add_image(xl)
        img_cell = ws_top.cell(row=base + 1, column=col, value=" ")
        img_cell.border = BORDER_BOX

        # Info lines
        info = [
            ("Amazon",    5),   # Title
            ("Price",     6),
            ("Sell",      7),
            ("MC SKU",    8),
            ("MC Title",  9),
            ("MC Retail", 10),
            ("MC Cost",   11),
            ("1-4 Avg",   12),
            ("Attributes",13),
            ("Notes",     14),
        ]
        for r_off, (label, dcol) in enumerate(info, start=2):
            cell = ws_top.cell(
                row=base + r_off,
                column=col,
                value=f'=CONCAT("{label}: ", {data_ref(dcol, drow)})'
            )
            cell.alignment = wrap_top
            cell.border = BORDER_BOX

    mem = BytesIO()
    wb.save(mem)
    mem.seek(0)
    return mem.read()

# Top-bar Download button (only when data exists)
if st.session_state.df is not None:
    with toolbar:
        download_placeholder.download_button(
            "Download Excel",
            data=build_xlsx(st.session_state.df.head(20)),
            file_name=f"{(st.session_state.meta or {'name':'Top_20'})['name'].replace(' ','_')}_Top20.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
