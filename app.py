# app.py
# Amazon Top-20 scraper → elegant Excel (5-per-row) with Sell-Through
# - Streamlit preview (image/title/price/sell-through)
# - Excel export with:
#     • Top 20 sheet: image row, Rank (orange), Sell-Through, Title, Price,
#                     MC SKU, MC Title, MC Retail, MC Cost, 1-4 Avg, Attributes, Notes
#       (images embedded; everything else references the Data sheet)
#     • Data sheet: MC fields first (for later manual fill), then Amazon fields
#
# Notes:
# - Amazon markup varies; sell-through is parsed via robust text search
# - If sell-through not found, inserts "—"
# - If image download fails, leaves the image cell empty

import os, re, uuid, time, random
from io import BytesIO
from datetime import datetime, timezone
from urllib.parse import urljoin
from typing import Optional, Tuple, List

import pandas as pd
import requests
from bs4 import BeautifulSoup
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

# ---------------- Streamlit config ----------------
st.set_page_config(page_title="Amazon Top 20 → Excel (Compact)", layout="wide")

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
]

def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"Accept-Language": "en-US,en;q=0.9"})
    return s

def get_soup(url: str, session: Optional[requests.Session] = None, timeout: int = 15) -> Tuple[BeautifulSoup, str]:
    session = session or _session()
    headers = {"User-Agent": random.choice(USER_AGENTS)}
    r = session.get(url, headers=headers, timeout=timeout)
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser"), r.url

# ---------------- Amazon scraping ----------------

ASIN_REGEXES = [
    re.compile(r"/dp/([A-Z0-9]{10})(?:[/?]|$)"),
    re.compile(r"/gp/product/([A-Z0-9]{10})(?:[/?]|$)"),
    re.compile(r"[?&]ASIN=([A-Z0-9]{10})(?:&|$)", re.IGNORECASE),
]

def extract_asin_from_url(url: str) -> Optional[str]:
    for rx in ASIN_REGEXES:
        m = rx.search(url)
        if m:
            return m.group(1)
    return None

def parse_top20_from_category_page(url: str, session: Optional[requests.Session] = None) -> List[dict]:
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

def extract_amazon_price(soup: BeautifulSoup) -> str:
    # Standard offscreen spans hold formatted price
    for off in soup.select("span.a-offscreen"):
        val = off.get_text(strip=True)
        if re.match(r"^\$\s*\d", val):
            return val.replace(" ", "")
    # Fallback meta tags
    meta = soup.find("meta", attrs={"itemprop": "price"})
    if meta and meta.get("content"):
        c = meta["content"]
        return c if c.startswith("$") else f"${c}"
    og = soup.find("meta", attrs={"property": "og:price:amount"})
    if og and og.get("content"):
        c = og["content"]
        return c if c.startswith("$") else f"${c}"
    return ""

def extract_amazon_image(soup: BeautifulSoup) -> str:
    og = soup.find("meta", attrs={"property": "og:image"})
    if og and og.get("content"):
        return og["content"]
    landing = soup.find("img", id="landingImage")
    if landing:
        for attr in ("data-old-hires", "src", "data-a-dynamic-image"):
            val = landing.get(attr)
            if val and isinstance(val, str) and val.strip():
                if attr == "data-a-dynamic-image":
                    m = re.search(r'"(https:[^"]+)"\s*:', val)
                    if m:
                        return m.group(1)
                else:
                    return val
    img = soup.select_one("#imgTagWrapperId img")
    if img:
        for attr in ("data-old-hires", "src"):
            if img.get(attr):
                return img.get(attr)
    return ""

def extract_sell_through(soup: BeautifulSoup) -> str:
    """
    Parse lines like "9K+ bought in past month" / "200+ bought in past month".
    Amazon varies classes, so we scan text.
    If missing or < 50, often omitted; we return '—' if nothing obvious is found.
    """
    txt = soup.get_text(" ", strip=True)
    # Common phrasing: "bought in past month", sometimes "bought in past week"
    m = re.search(r"([0-9,.Kk\+]+\s*(?:\+)?\s*bought\s+in\s+past\s+(?:month|week))", txt, re.I)
    if m:
        return m.group(1).strip()
    # Another pattern sometimes: "X bought last month"
    m2 = re.search(r"([0-9,.Kk\+]+\s*(?:\+)?\s*bought\s+last\s+month)", txt, re.I)
    if m2:
        return m2.group(1).strip()
    return "—"

def fetch_item_details_amzn(item_url: str, session: Optional[requests.Session] = None,
                            retries: int = 2, delay_range=(1.0, 2.0)) -> Tuple[str, str, str]:
    """
    Returns: (price, image_url, sell_through)
    """
    session = session or _session()
    for attempt in range(retries + 1):
        try:
            soup, _ = get_soup(item_url, session, timeout=20)
            price = extract_amazon_price(soup)
            image = extract_amazon_image(soup)
            sell = extract_sell_through(soup)
            return price, image, sell
        except Exception:
            if attempt >= retries:
                return "", "", "—"
            time.sleep(random.uniform(*delay_range))
    return "", "", "—"

# ---------------- Image helper for Excel ----------------

def download_and_resize_image(url: str, max_px: int = 120) -> Optional[BytesIO]:
    try:
        if not url or not isinstance(url, str):
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

# ---------------- Excel builder (compact, 5-per-row) ----------------

# Layout knobs
EXCEL_ITEM_COL_WIDTH = 26     # per product column width (Top 20)
EXCEL_IMG_MAX_PX = 120        # longest side for embedded image
EXCEL_IMG_ROW_HEIGHT = 90     # height for the image row so it doesn't overlap
ITEMS_PER_ROW = 5
# Block rows (per product, Top 20 sheet):
# 1: Image
# 2: Rank (orange)
# 3: Sell-Through
# 4: Amazon Title
# 5: Amazon Price
# 6: MC SKU
# 7: MC Title
# 8: MC Retail
# 9: MC Cost
# 10: 1-4 Avg
# 11: Attributes
# 12: Notes
BLOCK_ROWS = 12

def build_xlsx_with_links(df: pd.DataFrame) -> bytes:
    """
    Builds a two-sheet workbook:
    - Data: MC fields first, then Amazon fields (including Sell Through)
    - Top 20: embedded images + formula links back to Data sheet
    """
    wb = Workbook()
    ws_top = wb.active
    ws_top.title = "Top 20"
    ws_data = wb.create_sheet("Data")

    # ----- Data sheet columns (MC-first order)
    data_cols = [
        "MC SKU", "MC Title", "MC Retail", "MC Cost", "1-4 Avg", "Attributes", "Notes",
        "Rank", "ASIN", "Amazon Title", "Amazon Price", "Sell Through", "Amazon URL", "Image URL"
    ]
    for j, c in enumerate(data_cols, start=1):
        ws_data.cell(row=1, column=j, value=c)

    # Write Data rows (Amazon fields filled; MC fields blank for later)
    top = df.head(20).reset_index(drop=True)
    for i, r in top.iterrows():
        row = 2 + i
        ws_data.cell(row=row, column=1,  value="")                # MC SKU
        ws_data.cell(row=row, column=2,  value="")                # MC Title
        ws_data.cell(row=row, column=3,  value="")                # MC Retail
        ws_data.cell(row=row, column=4,  value="")                # MC Cost
        ws_data.cell(row=row, column=5,  value="")                # 1-4 Avg
        ws_data.cell(row=row, column=6,  value="")                # Attributes
        ws_data.cell(row=row, column=7,  value="")                # Notes
        ws_data.cell(row=row, column=8,  value=int(r.get("Rank", i+1)))  # Rank
        ws_data.cell(row=row, column=9,  value=r.get("ASIN",""))         # ASIN
        ws_data.cell(row=row, column=10, value=r.get("Title",""))        # Amazon Title
        ws_data.cell(row=row, column=11, value=r.get("Price",""))        # Amazon Price
        ws_data.cell(row=row, column=12, value=r.get("SellThrough","—")) # Sell Through
        ws_data.cell(row=row, column=13, value=r.get("URL",""))          # Amazon URL
        ws_data.cell(row=row, column=14, value=r.get("Image",""))        # Image URL

    # Column widths (Data)
    data_widths = [14, 50, 12, 12, 12, 24, 28, 8, 14, 56, 12, 20, 40, 50]
    for j, w in enumerate(data_widths, start=1):
        ws_data.column_dimensions[get_column_letter(j)].width = w
    ws_data.freeze_panes = "A2"

    # ----- Top 20 sheet styling
    thin = Side(style="thin", color="CCCCCC")
    border_box = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    font_bold = Font(bold=True, size=10)
    font_norm = Font(size=9)
    font_rank = Font(bold=True, size=11, color="FFFFFF")
    fill_rank = PatternFill("solid", fgColor="C45500")  # Amazon orange
    fill_even = PatternFill("solid", fgColor="F9F9F9")
    fill_sep = PatternFill("solid", fgColor="E0E0E0")

    # Column widths (Top 20)
    for c in range(1, ITEMS_PER_ROW + 1):
        ws_top.column_dimensions[get_column_letter(c)].width = EXCEL_ITEM_COL_WIDTH

    # Helper to build Data!A1 refs
    def DREF(col_idx: int, data_row: int) -> str:
        return f"Data!{get_column_letter(col_idx)}{data_row}"

    # Render each item block
    for idx in range(len(top)):
        group = idx // ITEMS_PER_ROW
        col = 1 + (idx % ITEMS_PER_ROW)
        base = 1 + group * (BLOCK_ROWS + 1)   # +1 to leave a separator row after each block
        data_row = 2 + idx                    # row in Data sheet
        band_fill = fill_even if (group % 2 == 0) else None

        # ----- Row 1: Embedded image
        img_url = ws_data.cell(row=data_row, column=14).value
        img_buf = download_and_resize_image(img_url, max_px=EXCEL_IMG_MAX_PX)
        if img_buf:
            xl_img = XLImage(img_buf)
            xl_img.anchor = f"{get_column_letter(col)}{base}"
            ws_top.add_image(xl_img)
        ws_top.row_dimensions[base].height = EXCEL_IMG_ROW_HEIGHT
        # Empty bordered cell behind image for uniform look / hyperlink to Amazon
        cell_img = ws_top.cell(row=base, column=col, value=" ")
        cell_img.alignment = center
        cell_img.border = border_box
        if band_fill: cell_img.fill = band_fill
        amazon_url = ws_data.cell(row=data_row, column=13).value
        if amazon_url:
            cell_img.hyperlink = amazon_url

        # ----- Row 2: Rank (orange)
        c = ws_top.cell(row=base + 1, column=col, value=f'= "Rank # " & {DREF(8, data_row)}')
        c.alignment = center
        c.font = font_rank
        c.fill = fill_rank
        c.border = border_box

        # ----- Row 3: Sell-Through (subtle)
        c = ws_top.cell(row=base + 2, column=col, value=f'= {DREF(12, data_row)}')
        c.alignment = left
        c.font = font_norm
        c.border = border_box
        if band_fill: c.fill = band_fill

        # ----- Row 4: Amazon Title
        c = ws_top.cell(row=base + 3, column=col, value=f'= {DREF(10, data_row)}')
        c.alignment = left
        c.font = font_norm
        c.border = border_box
        if band_fill: c.fill = band_fill

        # ----- Row 5: Amazon Price (bold)
        c = ws_top.cell(row=base + 4, column=col, value=f'= {DREF(11, data_row)}')
        c.alignment = left
        c.font = font_bold
        c.border = border_box
        if band_fill: c.fill = band_fill

        # ----- Row 6: MC SKU
        c = ws_top.cell(row=base + 5, column=col, value=f'= "MC SKU: " & {DREF(1, data_row)}')
        c.alignment = left
        c.font = font_norm
        c.border = border_box
        if band_fill: c.fill = band_fill

        # ----- Row 7: MC Title
        c = ws_top.cell(row=base + 6, column=col, value=f'= "MC Title: " & {DREF(2, data_row)}')
        c.alignment = left
        c.font = font_norm
        c.border = border_box
        if band_fill: c.fill = band_fill

        # ----- Row 8: MC Retail (bold)
        c = ws_top.cell(row=base + 7, column=col, value=f'= "MC Retail: " & {DREF(3, data_row)}')
        c.alignment = left
        c.font = font_bold
        c.border = border_box
        if band_fill: c.fill = band_fill

        # ----- Row 9: MC Cost
        c = ws_top.cell(row=base + 8, column=col, value=f'= "MC Cost: " & {DREF(4, data_row)}')
        c.alignment = left
        c.font = font_norm
        c.border = border_box
        if band_fill: c.fill = band_fill

        # ----- Row 10: 1-4 Avg
        c = ws_top.cell(row=base + 9, column=col, value=f'= "1-4 Avg: " & {DREF(5, data_row)}')
        c.alignment = left
        c.font = font_norm
        c.border = border_box
        if band_fill: c.fill = band_fill

        # ----- Row 11: Attributes
        c = ws_top.cell(row=base + 10, column=col, value=f'= "Attributes: " & {DREF(6, data_row)}')
        c.alignment = left
        c.font = font_norm
        c.border = border_box
        if band_fill: c.fill = band_fill

        # ----- Row 12: Notes
        c = ws_top.cell(row=base + 11, column=col, value=f'= "Notes: " & {DREF(7, data_row)}')
        c.alignment = left
        c.font = font_norm
        c.border = border_box
        if band_fill: c.fill = band_fill

        # ----- Separator row after each 5-item row (visual grouping)
        # Only draw once per group under the last column
        if (idx % ITEMS_PER_ROW) == (ITEMS_PER_ROW - 1):
            sep_row = base + BLOCK_ROWS
            for ccol in range(1, ITEMS_PER_ROW + 1):
                ws_top.cell(row=sep_row, column=ccol).fill = fill_sep

    # No freeze on Top 20; wrapped text everywhere by default
    return_bytes = BytesIO()
    wb.save(return_bytes)
    return_bytes.seek(0)
    return return_bytes.read()

# ---------------- Streamlit UI ----------------

st.title("🧭 Amazon Top 20 → Excel (Compact, Linked)")

col1, col2 = st.columns([3, 1])
with col1:
    amz_url = st.text_input("Amazon Best Sellers URL", placeholder="https://www.amazon.com/gp/bestsellers/pc/17441247011")
with col2:
    delay = st.slider("Delay (sec)", 0.5, 3.0, 1.0, 0.1)

fetch_btn = st.button("Fetch Top 20", type="primary")
export_slot = st.empty()

if fetch_btn:
    if not amz_url.strip().startswith("http"):
        st.error("Please enter a valid URL that starts with http(s)://")
    else:
        s = _session()
        st.info("Fetching top 20 items from Amazon…")
        try:
            items = parse_top20_from_category_page(amz_url.strip(), s)
        except Exception as e:
            st.error(f"Failed to fetch category page: {e}")
            items = []

        rows = []
        for i, it in enumerate(items):
            price, image, sell = fetch_item_details_amzn(it["URL"], s)
            rows.append({
                "Rank": i + 1,
                "ASIN": it["ASIN"],
                "Title": it["Title"],
                "Price": price,
                "SellThrough": sell if sell else "—",
                "URL": it["URL"],
                "Image": image
            })
            time.sleep(delay)

        df = pd.DataFrame(rows)
        st.session_state.results = df
        st.success(f"Top 20 fetched!")

# Preview list
if st.session_state.get("results") is not None:
    df = st.session_state["results"]
    for i, r in df.iterrows():
        c1, c2 = st.columns([1, 4])
        with c1:
            try:
                if r["Image"]:
                    resp = requests.get(r["Image"], timeout=10)
                    resp.raise_for_status()
                    st.image(resp.content, width=110)
                else:
                    st.write("🖼️ (no image)")
            except Exception:
                st.write("🖼️ (no image)")
        with c2:
            st.markdown(f"**#{int(r['Rank'])}: {r['Title'] or '(no title)'}**")
            if r.get("SellThrough") and r["SellThrough"] != "—":
                st.caption(r["SellThrough"])
            st.write(r.get("Price", ""))
            st.caption(r.get("URL", ""))
        st.divider()

    # Export button
    if export_slot.button("⬇️ Download Excel (Top 20 + Data)", type="primary"):
        xlsx = build_xlsx_with_links(df)
        fname = f"Amazon_Top20_{uuid.uuid4().hex[:6]}.xlsx"
        st.download_button("Download Spreadsheet", data=xlsx, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
