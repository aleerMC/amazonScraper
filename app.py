import os
import re
import json
import uuid
import time
import random
from io import BytesIO
from datetime import datetime, timezone
from urllib.parse import urljoin, urlparse
from typing import Optional, List, Dict

import pandas as pd
import requests
from bs4 import BeautifulSoup
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

# ---------------- Excel layout knobs ----------------
EXCEL_COL_WIDTH = 24          # width per item column
EXCEL_IMG_MAX_PX = 220        # max embedded image size (longest side)
EXCEL_IMG_ROW_HEIGHT = 90     # image row height
ITEMS_PER_ROW = 5             # 5 items per "band"
BLOCK_ROWS = 12               # rows per item block (1 image + 11 text lines)
BAND_COLOR_1 = "FFFFFF"       # white
BAND_COLOR_2 = "F7F7F7"       # light gray
RANK_BGCOLOR = "C45500"       # Amazon-style orange for Rank row
RANK_FGCOLOR = "FFFFFF"       # White text for Rank row

# ====================================================
# Basic setup
# ====================================================

st.set_page_config(
    page_title="Amazon Top-20 Scraper",
    layout="wide",
)

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
]

SAVED_DIR = ".saved_searches"
os.makedirs(SAVED_DIR, exist_ok=True)


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"Accept-Language": "en-US,en;q=0.9"})
    return s


def utc_now_str() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")


# ====================================================
# Amazon scraping
# ====================================================

def get_soup(url: str, session: Optional[requests.Session] = None, timeout: int = 15):
    """
    Fetch URL and return (BeautifulSoup, final_url).
    """
    if not url.lower().startswith("http"):
        raise ValueError("URL must start with http:// or https://")
    session = session or _session()
    headers = {"User-Agent": random.choice(USER_AGENTS)}
    r = session.get(url, headers=headers, timeout=timeout)
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser"), r.url


ASIN_REGEX = re.compile(r"/dp/([A-Z0-9]{10})")


def _extract_from_faceouts(soup: BeautifulSoup) -> List[Dict]:
    """
    Preferred parser for Best Sellers pages:
    - Uses <div class="p13n-sc-uncoverable-faceout" id="<ASIN>"> blocks
    - Extracts ASIN, title, URL, image URL, price, and (if present) sell-through text.
    """
    base = "https://www.amazon.com"
    items = []

    faceouts = soup.find_all("div", class_="p13n-sc-uncoverable-faceout")
    for div in faceouts:
        asin = div.get("id") or ""
        if not asin or len(asin) != 10:
            # fallback from href
            for a in div.find_all("a", href=True):
                m = ASIN_REGEX.search(a["href"])
                if m:
                    asin = m.group(1)
                    break

        # Title + URL
        title = ""
        url = ""
        # Prefer anchors with non-price text
        for a in div.find_all("a", href=True):
            text = a.get_text(" ", strip=True)
            if text and "$" not in text:
                url = a["href"]
                title = text
                break
        if not url:
            # fallback: first anchor
            a = div.find("a", href=True)
            if a:
                url = a["href"]
                title = a.get_text(" ", strip=True)

        if url and url.startswith("/"):
            url = base + url

        # Image
        img_url = ""
        img = div.find("img")
        if img:
            img_url = img.get("src") or ""

        # Sell-through text (e.g. "9K+ bought in past month")
        sell = ""
        for span in div.find_all("span"):
            t = span.get_text(" ", strip=True)
            if "bought" in t.lower():
                sell = t
                break

        # Price: pick shortest text containing a $ and extract the $XX.XX
        price_candidates = []
        for el in div.find_all(["span", "a", "div"], recursive=True):
            t = el.get_text(" ", strip=True)
            if "$" in t:
                m = re.search(r"\$\s*\d[\d,]*(?:\.\d{2})?", t)
                if m:
                    price_candidates.append(m.group(0).replace(" ", ""))
        price = ""
        if price_candidates:
            price = min(price_candidates, key=len)

        items.append(
            {
                "ASIN": asin,
                "Title": title,
                "URL": url,
                "ImageURL": img_url,
                "AmazonPrice": price,
                "AmazonSellThru": sell,
            }
        )
        if len(items) >= 20:
            break

    return items


def _extract_from_links(soup: BeautifulSoup, base_url: str) -> List[Dict]:
    """
    Fallback parser: scan all anchors for unique ASINs.
    Used if faceout-based parsing fails for some reason.
    """
    items = []
    seen = set()
    base = base_url

    for a in soup.find_all("a", href=True):
        href = a["href"]
        m = ASIN_REGEX.search(href)
        if not m:
            continue
        asin = m.group(1)
        if asin in seen:
            continue

        title = (a.get_text(strip=True) or "").strip()
        if not title:
            img = a.find("img", alt=True)
            if img and img.get("alt"):
                title = img["alt"].strip()

        url = urljoin(base, href) if href.startswith("/") else href
        items.append(
            {
                "ASIN": asin,
                "Title": title,
                "URL": url,
                "ImageURL": "",
                "AmazonPrice": "",
                "AmazonSellThru": "",
            }
        )
        seen.add(asin)
        if len(items) >= 20:
            break

    return items


def parse_top20_from_category_page(url: str, session: Optional[requests.Session] = None) -> List[Dict]:
    """
    High-level parser: try faceout blocks first; if we get nothing, fall back to link scan.
    """
    session = session or _session()
    soup, final_url = get_soup(url, session)
    items = _extract_from_faceouts(soup)
    if not items:
        items = _extract_from_links(soup, final_url)
    return items


# ====================================================
# Persistence (simple CSV + JSON, no pyarrow)
# ====================================================

def _run_path(run_id: str) -> str:
    return os.path.join(SAVED_DIR, run_id)


def _meta_path(run_id: str) -> str:
    return os.path.join(_run_path(run_id), "meta.json")


def _data_path(run_id: str) -> str:
    return os.path.join(_run_path(run_id), "data.csv")


def list_saved_runs() -> List[Dict]:
    runs = []
    for rid in os.listdir(SAVED_DIR):
        rp = _run_path(rid)
        mp = _meta_path(rid)
        dp = _data_path(rid)
        if os.path.isdir(rp) and os.path.exists(mp) and os.path.exists(dp):
            try:
                with open(mp, "r", encoding="utf-8") as f:
                    meta = json.load(f)
                runs.append({"id": rid, **meta})
            except Exception:
                continue
    runs.sort(key=lambda r: r.get("updated_at", ""), reverse=True)
    return runs


def save_run(run_id: str, df: pd.DataFrame, meta: dict):
    os.makedirs(_run_path(run_id), exist_ok=True)
    df.to_csv(_data_path(run_id), index=False, encoding="utf-8")
    meta["updated_at"] = utc_now_str()
    with open(_meta_path(run_id), "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2)


def load_run(run_id: str):
    df = pd.read_csv(_data_path(run_id), dtype=str).fillna("")
    with open(_meta_path(run_id), "r", encoding="utf-8") as f:
        meta = json.load(f)
    if "Rank" in df.columns:
        with pd.option_context("mode.chained_assignment", None):
            df["Rank"] = pd.to_numeric(df["Rank"], errors="coerce").fillna(0).astype(int)
    return df, meta


def make_meta(category_desc: str, amazon_url: str, fetched_at: str) -> dict:
    now = utc_now_str()
    return {
        "name": category_desc or "Top 20",
        "category_desc": category_desc or "",
        "amazon_url": amazon_url or "",
        "fetched_at": fetched_at or now,
        "created_at": now,
        "updated_at": now,
    }


# ====================================================
# Session state
# ====================================================

if "results" not in st.session_state:
    st.session_state.results = None
if "current_run_id" not in st.session_state:
    st.session_state.current_run_id = None
if "current_meta" not in st.session_state:
    st.session_state.current_meta = None
if "view_mode" not in st.session_state:
    st.session_state.view_mode = "Grid (4x5)"


# ====================================================
# Sidebar: Saved Searches (simple & clean)
# ====================================================

with st.sidebar:
    st.header("Saved Searches")

    saved = list_saved_runs()
    if saved:
        labels = [f"{r['name']}  —  {r.get('updated_at','')}" for r in saved]
        sel_idx = st.selectbox(
            "Saved runs",
            options=list(range(len(saved))),
            format_func=lambda i: labels[i],
        )
        sel = saved[sel_idx]

        c1, c2 = st.columns(2)
        with c1:
            if st.button("Load"):
                df, meta = load_run(sel["id"])
                st.session_state.results = df
                st.session_state.current_run_id = sel["id"]
                st.session_state.current_meta = meta
                st.success(f"Loaded: {meta.get('name')}")
        with c2:
            if st.button("Delete"):
                import shutil
                try:
                    shutil.rmtree(_run_path(sel["id"]), ignore_errors=True)
                    st.warning("Deleted. Refreshing list…")
                    st.experimental_rerun()
                except Exception as e:
                    st.error(f"Delete failed: {e}")
    else:
        st.caption("No saved searches yet.")

    with st.expander("Scrape Settings", expanded=False):
        st.caption("Adjust delays if Amazon gets cranky.")
        dmin = st.slider("Min per-item delay (sec)", 0.5, 5.0, 1.0, 0.1, key="dmin")
        dmax = st.slider("Max per-item delay (sec)", 0.6, 6.0, 2.0, 0.1, key="dmax")
        if dmax < dmin:
            st.session_state.dmax = dmin + 0.1


# ====================================================
# Main layout
# ====================================================

st.title("Amazon Top-20 Scraper → Excel Layout")

# ---------- Top control bar ----------
top_bar = st.container()
with top_bar:
    col_bar = st.columns([1.2, 1.2, 1.6, 2.5, 2.5])
    with col_bar[0]:
        fetch_btn = st.button("Fetch Top 20", type="primary")
    with col_bar[1]:
        save_as_new_btn = st.button("Save as New", disabled=(st.session_state.results is None))
    with col_bar[2]:
        export_placeholder = st.empty()
    with col_bar[3]:
        st.session_state.view_mode = st.radio(
            "View",
            options=["Grid (4x5)", "List"],
            horizontal=True,
            label_visibility="visible",
        )
    with col_bar[4]:
        st.markdown("")  # spacer

# ---------- Inputs under top bar ----------
col_in1, col_in2 = st.columns([3, 2])
with col_in1:
    amz_url = st.text_input(
        "Amazon Best Sellers URL",
        placeholder="https://www.amazon.com/gp/bestsellers/pc/17441247011",
    )
with col_in2:
    category_desc = st.text_input(
        "Short internal name (for saved searches)",
        placeholder="Single Board Computers",
    )

# ====================================================
# Fetch handler
# ====================================================

if fetch_btn:
    if not amz_url.strip():
        st.error("Please enter a full Amazon Best Sellers URL.")
    else:
        session = _session()
        st.info("Fetching top 20 items from Amazon...")
        try:
            items = parse_top20_from_category_page(amz_url.strip(), session=session)
        except Exception as e:
            st.error(f"Failed to fetch/parse Amazon page: {e}")
            items = []

        if not items:
            st.warning("No items parsed. Amazon might have changed the layout or blocked the request.")
        else:
            rows = []
            ts = utc_now_str()
            for idx, it in enumerate(items):
                rows.append(
                    {
                        "ImageURL": it.get("ImageURL", ""),
                        "Rank": idx + 1,
                        "Title": it.get("Title", ""),
                        "ASIN": it.get("ASIN", ""),
                        "AmazonURL": it.get("URL", ""),
                        "AmazonPrice": it.get("AmazonPrice", ""),
                        "AmazonSellThru": it.get("AmazonSellThru", ""),
                        # MC fields for later manual use in Excel:
                        "MCSKU": "",
                        "MCTitle": "",
                        "MCRetail": "",
                        "MCCost": "",
                        "Avg1_4": "",
                        "AttrMatch": "",
                        "Notes": "",
                        # Meta:
                        "FetchedAt": ts,
                        "CategoryDesc": category_desc.strip(),
                        "AmazonBestURL": amz_url.strip(),
                    }
                )
                time.sleep(random.uniform(st.session_state.dmin, st.session_state.dmax))

            df = pd.DataFrame(
                rows,
                columns=[
                    "ImageURL",
                    "Rank",
                    "Title",
                    "ASIN",
                    "AmazonURL",
                    "AmazonPrice",
                    "AmazonSellThru",
                    "MCSKU",
                    "MCTitle",
                    "MCRetail",
                    "MCCost",
                    "Avg1_4",
                    "AttrMatch",
                    "Notes",
                    "FetchedAt",
                    "CategoryDesc",
                    "AmazonBestURL",
                ],
            )
            st.session_state.results = df
            st.session_state.current_run_id = None
            st.session_state.current_meta = None
            st.success("Top 20 fetched!")

# ====================================================
# Save-as-new handler
# ====================================================

if save_as_new_btn and st.session_state.results is not None:
    df = st.session_state.results
    cat = df.iloc[0].get("CategoryDesc", "") if not df.empty else category_desc
    url_meta = df.iloc[0].get("AmazonBestURL", "") if not df.empty else amz_url
    fetched_at = df.iloc[0].get("FetchedAt", utc_now_str()) if not df.empty else utc_now_str()
    meta = make_meta(cat, url_meta, fetched_at)
    run_id = uuid.uuid4().hex[:12]
    save_run(run_id, df, meta)
    st.session_state.current_run_id = run_id
    st.session_state.current_meta = meta
    st.success(f"Saved as new: {meta['name']}")
    st.experimental_rerun()

# ====================================================
# Display results (List / Grid) — no re-scrape on view change
# ====================================================

df = st.session_state.results

if df is not None and not df.empty:
    meta = st.session_state.current_meta or {
        "name": df.iloc[0].get("CategoryDesc") or "Top 20",
        "fetched_at": df.iloc[0].get("FetchedAt"),
        "amazon_url": df.iloc[0].get("AmazonBestURL"),
    }

    # Header info
    hdr_cols = st.columns([3, 2, 2])
    with hdr_cols[0]:
        st.subheader(meta.get("name", "Top 20"))
    with hdr_cols[1]:
        st.caption(f"Pulled at: {meta.get('fetched_at', '')}")
    with hdr_cols[2]:
        if meta.get("amazon_url"):
            st.caption(f"[Source: Amazon Best Sellers]({meta['amazon_url']})")

    st.markdown(
        """
        <style>
        .tight p { margin: 0.1rem 0 !important; }
        .tiny { font-size: 0.8rem; color: #6b7280; }
        .price { font-weight: 600; }
        .card {
            background-color: #ffffff;
            border-radius: 8px;
            border: 1px solid #e5e7eb;
            padding: 0.5rem;
            box-shadow: 0 1px 2px rgba(15,23,42,0.05);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # ---------- List view ----------
    if st.session_state.view_mode == "List":
        for _, row in df.iterrows():
            c1, c2 = st.columns([1, 5])
            with c1:
                if isinstance(row["ImageURL"], str) and row["ImageURL"]:
                    st.image(row["ImageURL"], width=120)
                else:
                    st.write("—")
            with c2:
                st.markdown(
                    f"<div class='card tight'>"
                    f"<b>#{int(row['Rank'])}</b> &nbsp; {row['Title'] or '(no title)'}<br>"
                    f"<span class='price'>{row['AmazonPrice'] or ''}</span>"
                    f"{' &nbsp; • &nbsp; ' + row['AmazonSellThru'] if row['AmazonSellThru'] else ''}<br>"
                    f"<span class='tiny'>ASIN: {row['ASIN']} &nbsp;|&nbsp; "
                    f"<a href='{row['AmazonURL']}' target='_blank'>Amazon Link</a></span>"
                    f"</div>",
                    unsafe_allow_html=True,
                )
            st.markdown("---")

    # ---------- Grid view (4 rows x 5 columns) ----------
    else:
        top20 = df.head(20).reset_index(drop=True)
        num_items = len(top20)
        # 4 rows of cards, 5 per row (20 total)
        for row_idx in range(0, num_items, ITEMS_PER_ROW):
            cols = st.columns(ITEMS_PER_ROW)
            for col_offset in range(ITEMS_PER_ROW):
                idx = row_idx + col_offset
                if idx >= num_items:
                    continue
                row = top20.iloc[idx]
                with cols[col_offset]:
                    st.markdown("<div class='card'>", unsafe_allow_html=True)
                    if isinstance(row["ImageURL"], str) and row["ImageURL"]:
                        st.image(row["ImageURL"], use_column_width="auto")
                    st.markdown(
                        f"<div class='tight'>"
                        f"<b># {int(row['Rank'])}</b><br>"
                        f"{row['Title'] or '(no title)'}<br>"
                        f"<span class='price'>{row['AmazonPrice'] or ''}</span><br>"
                        f"{row['AmazonSellThru'] if row['AmazonSellThru'] else ''}<br>"
                        f"<span class='tiny'>ASIN: {row['ASIN']}</span><br>"
                        f"<span class='tiny'><a href='{row['AmazonURL']}' target='_blank'>Amazon Link</a></span>"
                        f"</div>",
                        unsafe_allow_html=True,
                    )
                    st.markdown("</div>", unsafe_allow_html=True)

    # ====================================================
    # Excel export builder (Data + Top 20 layout)
    # ====================================================

    def _download_image_bytes(url: str, max_px: int = EXCEL_IMG_MAX_PX) -> Optional[BytesIO]:
        """
        Fetch & resize image; return PNG bytes or None if failed.
        Small retry to reduce random drops.
        """
        if not url or not isinstance(url, str):
            return None
        for attempt in range(2):
            try:
                resp = requests.get(url, headers={"User-Agent": random.choice(USER_AGENTS)}, timeout=15)
                resp.raise_for_status()
                img = PILImage.open(BytesIO(resp.content)).convert("RGBA")
                img.thumbnail((max_px, max_px), PILImage.LANCZOS)
                buff = BytesIO()
                img.save(buff, format="PNG")
                buff.seek(0)
                return buff
            except Exception:
                if attempt == 1:
                    return None
                time.sleep(0.5)
        return None

    def build_xlsx(df_src: pd.DataFrame) -> bytes:
        wb = Workbook()
        ws_top = wb.active
        ws_top.title = "Top 20"
        ws_data = wb.create_sheet("Data")

        # ----- Data sheet -----
        cols = [
            "Rank",            # 1
            "ASIN",            # 2
            "AmazonURL",       # 3
            "ImageURL",        # 4
            "Title",           # 5
            "AmazonPrice",     # 6
            "AmazonSellThru",  # 7
            "MCSKU",           # 8
            "MCTitle",         # 9
            "MCRetail",        #10
            "MCCost",          #11
            "Avg1_4",          #12
            "AttrMatch",       #13
            "Notes",           #14
            "FetchedAt",       #15
            "CategoryDesc",    #16
            "AmazonBestURL",   #17
        ]
        for j, c in enumerate(cols, start=1):
            ws_data.cell(row=1, column=j, value=c)

        top20 = df_src.head(20).reset_index(drop=True)
        for i, r in top20.iterrows():
            values = [
                int(r.get("Rank", i + 1)),
                r.get("ASIN", ""),
                r.get("AmazonURL", ""),
                r.get("ImageURL", ""),
                r.get("Title", ""),
                r.get("AmazonPrice", ""),
                r.get("AmazonSellThru", ""),
                r.get("MCSKU", ""),
                r.get("MCTitle", ""),
                r.get("MCRetail", ""),
                r.get("MCCost", ""),
                r.get("Avg1_4", ""),
                r.get("AttrMatch", ""),
                r.get("Notes", ""),
                r.get("FetchedAt", ""),
                r.get("CategoryDesc", ""),
                r.get("AmazonBestURL", ""),
            ]
            for j, v in enumerate(values, start=1):
                ws_data.cell(row=2 + i, column=j, value=v)

        data_widths = [8, 12, 30, 30, 50, 12, 18, 12, 50, 12, 12, 12, 20, 30, 20, 24, 32]
        for j, w in enumerate(data_widths, start=1):
            ws_data.column_dimensions[get_column_letter(j)].width = w
        ws_data.freeze_panes = "A2"

        # ----- Top 20 layout -----
        for c in range(1, ITEMS_PER_ROW + 1):
            ws_top.column_dimensions[get_column_letter(c)].width = EXCEL_COL_WIDTH

        left_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)
        left_mid = Alignment(horizontal="left", vertical="center", wrap_text=True)
        center_mid = Alignment(horizontal="center", vertical="center", wrap_text=True)

        thin = Side(style="thin", color="DDDDDD")
        box = Border(left=thin, right=thin, top=thin, bottom=thin)
        band1 = PatternFill("solid", fgColor=BAND_COLOR_1)
        band2 = PatternFill("solid", fgColor=BAND_COLOR_2)
        rank_fill = PatternFill("solid", fgColor=RANK_BGCOLOR)
        rank_font = Font(bold=True, color=RANK_FGCOLOR)
        bold = Font(bold=True)

        def DREF(col_idx: int, data_row: int) -> str:
            return f"Data!{get_column_letter(col_idx)}{data_row}"

        for idx in range(len(top20)):
            group = idx // ITEMS_PER_ROW
            col = 1 + (idx % ITEMS_PER_ROW)
            base = 1 + group * BLOCK_ROWS
            data_row = 2 + idx

            band_fill = band1 if (group % 2 == 0) else band2

            # Image row
            img_url = ws_data.cell(row=data_row, column=4).value
            amz_url = ws_data.cell(row=data_row, column=3).value

            img_buf = _download_image_bytes(img_url, max_px=EXCEL_IMG_MAX_PX)
            if img_buf:
                xl_img = XLImage(img_buf)
                xl_img.anchor = f"{get_column_letter(col)}{base}"
                ws_top.add_image(xl_img)

            cell = ws_top.cell(row=base, column=col, value=" ")
            if amz_url:
                cell.hyperlink = amz_url
            cell.alignment = center_mid
            cell.border = box
            cell.fill = band_fill
            ws_top.row_dimensions[base].height = EXCEL_IMG_ROW_HEIGHT

            # Rank row (with colored background)
            c = ws_top.cell(row=base + 1, column=col, value=f'= "Rank: #" & {DREF(1, data_row)}')
            c.alignment = left_mid
            c.font = rank_font
            c.border = box
            c.fill = rank_fill

            # Amazon title
            c = ws_top.cell(row=base + 2, column=col, value=f'= "Amazon: " & {DREF(5, data_row)}')
            c.alignment = left_wrap
            c.border = box
            c.fill = band_fill

            # Amazon price
            c = ws_top.cell(row=base + 3, column=col, value=f'= "Amazon Price: " & {DREF(6, data_row)}')
            c.alignment = left_mid
            c.font = bold
            c.border = box
            c.fill = band_fill

            # Sell-through
            c = ws_top.cell(row=base + 4, column=col, value=f'= "Sell-through: " & {DREF(7, data_row)}')
            c.alignment = left_mid
            c.border = box
            c.fill = band_fill

            # MC SKU
            c = ws_top.cell(row=base + 5, column=col, value=f'= "MC SKU: " & {DREF(8, data_row)}')
            c.alignment = left_mid
            c.border = box
            c.fill = band_fill

            # MC Title
            c = ws_top.cell(row=base + 6, column=col, value=f'= "MC Title: " & {DREF(9, data_row)}')
            c.alignment = left_wrap
            c.border = box
            c.fill = band_fill

            # MC Retail
            c = ws_top.cell(row=base + 7, column=col, value=f'= "MC Retail: " & {DREF(10, data_row)}')
            c.alignment = left_mid
            c.font = bold
            c.border = box
            c.fill = band_fill

            # MC Cost
            c = ws_top.cell(row=base + 8, column=col, value=f'= "MC Cost: " & {DREF(11, data_row)}')
            c.alignment = left_mid
            c.border = box
            c.fill = band_fill

            # 1–4 Avg
            c = ws_top.cell(row=base + 9, column=col, value=f'= "1-4 Avg: " & {DREF(12, data_row)}')
            c.alignment = left_mid
            c.border = box
            c.fill = band_fill

            # Attributes
            c = ws_top.cell(row=base + 10, column=col, value=f'= "Attributes: " & {DREF(13, data_row)}')
            c.alignment = left_wrap
            c.border = box
            c.fill = band_fill

            # Notes
            c = ws_top.cell(row=base + 11, column=col, value=f'= "Notes: " & {DREF(14, data_row)}')
            c.alignment = left_wrap
            c.border = box
            c.fill = band_fill

        ws_top.freeze_panes = None  # no freeze on Top 20 sheet

        bio = BytesIO()
        wb.save(bio)
        bio.seek(0)
        return bio.read()

    # Single download button in the top bar
    export_placeholder.download_button(
        "Download Excel",
        data=build_xlsx(df),
        file_name=f"{(meta.get('name') or 'Top20').replace(' ', '_')}_Top20.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

else:
    st.info("Enter an Amazon Best Sellers URL and click **Fetch Top 20** to begin.")
