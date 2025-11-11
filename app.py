import os, re, uuid, time, random
from io import BytesIO
from datetime import datetime, timezone
from urllib.parse import urljoin
import pandas as pd, requests
from bs4 import BeautifulSoup
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

# ---------- Streamlit Config ----------
st.set_page_config(page_title="Amazon Top 20 Scraper", layout="wide")
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
]

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

def extract_asin_from_url(url):
    for pattern in [r"/dp/([A-Z0-9]{10})(?:[/?]|$)",
                    r"/gp/product/([A-Z0-9]{10})(?:[/?]|$)"]:
        m = re.search(pattern, url)
        if m:
            return m.group(1)
    return None

def parse_top20_from_category_page(url, session=None):
    session = session or _session()
    soup, final_url = get_soup(url, session)
    seen, items = set(), []
    for a in soup.find_all("a", href=True):
        asin = extract_asin_from_url(a["href"])
        if not asin or asin in seen:
            continue
        title = (a.get_text(strip=True) or "").strip()
        if not title:
            img = a.find("img", alt=True)
            if img and img.get("alt"):
                title = img["alt"].strip()
        if not title and a.get("title"):
            title = a["title"].strip()
        items.append({
            "ASIN": asin,
            "Title": title,
            "URL": urljoin(final_url, a["href"]),
        })
        seen.add(asin)
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

def download_and_resize_image(url, max_px=140):
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

# ---------- Elegant Excel builder ----------
def build_excel(df):
    wb = Workbook()
    ws_top = wb.active
    ws_top.title = "Top 20"
    ws_data = wb.create_sheet("Data")

    # --- Data tab ---
    cols = ["MC SKU","MC Title","MC Retail","MC Cost","1-4 Avg","Attributes","Notes",
            "ASIN","Amazon Title","Amazon Price","Amazon URL","Image URL"]
    for i, c in enumerate(cols, 1):
        ws_data.cell(row=1, column=i, value=c)
    for i, row in enumerate(df.itertuples(), 2):
        ws_data.cell(row=i, column=8, value=row.ASIN)
        ws_data.cell(row=i, column=9, value=row.Title)
        ws_data.cell(row=i, column=10, value=row.Price)
        ws_data.cell(row=i, column=11, value=row.URL)
        ws_data.cell(row=i, column=12, value=row.Image)

    for j, w in enumerate([14,40,14,12,12,20,20,12,50,12,40,50],1):
        ws_data.column_dimensions[get_column_letter(j)].width=w

    # --- Top 20 tab styling ---
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin,right=thin,top=thin,bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    font_title = Font(bold=True, size=10)
    font_norm = Font(size=9)
    font_rank = Font(bold=True, size=11, color="FFFFFF")
    fill_rank = PatternFill("solid", fgColor="C45500")
    fill_even = PatternFill("solid", fgColor="F9F9F9")
    fill_sep  = PatternFill("solid", fgColor="E0E0E0")

    items_per_row = 5
    block_rows = 11  # 1 image + 10 info lines
    img_px = 110

    for c in range(1, items_per_row+1):
        ws_top.column_dimensions[get_column_letter(c)].width = 26

    for idx,row in enumerate(df.itertuples(),0):
        group = idx // items_per_row
        col = 1 + (idx % items_per_row)
        base = 1 + group * block_rows

        fill = fill_even if group % 2 == 0 else None

        # image row
        img_buf = download_and_resize_image(row.Image, img_px)
        if img_buf:
            xl_img = XLImage(img_buf)
            xl_img.anchor = f"{get_column_letter(col)}{base}"
            ws_top.add_image(xl_img)
        ws_top.row_dimensions[base].height = img_px * 0.75

        def write(line, val, bold=False, align=None, fill_override=None, font_override=None):
            cell = ws_top.cell(row=base+line, column=col, value=val)
            cell.alignment = align or left
            cell.font = font_override or (font_title if bold else font_norm)
            cell.border = border
            if fill_override:
                cell.fill = fill_override
            elif fill:
                cell.fill = fill

        # rank header row
        write(1, None)  # image occupies row 1
        write(2, f"Rank #{idx+1}", bold=True, align=center,
              fill_override=fill_rank, font_override=font_rank)
        write(3, row.Title)
        write(4, row.Price, bold=True)
        write(5, "MC SKU:")
        write(6, "MC Title:")
        write(7, "MC Retail:")
        write(8, "MC Cost:")
        write(9, "1-4 Avg:")
        write(10, "Attributes:")
        write(11, "Notes:")

        # separator below each block
        for c2 in range(1, items_per_row+1):
            ws_top.cell(row=base+block_rows, column=c2).fill = fill_sep

    return wb

# ---------- Streamlit UI ----------
st.title("🧭 Amazon Top 20 Scraper — Compact Report")

url = st.text_input("Amazon Best Sellers URL", placeholder="https://www.amazon.com/gp/bestsellers/...")
delay = st.slider("Delay between requests (seconds)", 0.5, 3.0, 1.0)

if st.button("Fetch Top 20", type="primary"):
    if not url.startswith("http"):
        st.error("Enter a valid Amazon URL starting with https://")
    else:
        s = _session()
        st.info("Fetching top 20 items from Amazon …")
        items = parse_top20_from_category_page(url, s)
        rows=[]
        for i,it in enumerate(items):
            price,img=fetch_item_details(it["URL"], s)
            rows.append({"Rank":i+1,"ASIN":it["ASIN"],"Title":it["Title"],
                         "Price":price,"URL":it["URL"],"Image":img})
            time.sleep(delay)
        df=pd.DataFrame(rows)
        st.session_state.results=df
        st.success(f"Fetched {len(df)} items!")

if "results" in st.session_state and st.session_state.results is not None:
    df=st.session_state.results
    for i,row in df.iterrows():
        c1,c2=st.columns([1,3])
        with c1:
            try:
                if row["Image"]:
                    img=requests.get(row["Image"],timeout=10).content
                    st.image(img,width=100)
            except Exception:
                st.write("🖼️ (no image)")
        with c2:
            st.markdown(f"**#{i+1}: {row['Title']}**")
            st.write(row["Price"])
            st.caption(row["URL"])
        st.divider()

    if st.button("Export to Excel"):
        wb=build_excel(df)
        fn=f"Amazon_Top20_{uuid.uuid4().hex[:6]}.xlsx"
        wb.save(fn)
        with open(fn,"rb") as f:
            st.download_button("⬇️ Download Spreadsheet", f, file_name=fn)
