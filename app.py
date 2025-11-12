import re, os, json, uuid, time, random
from io import BytesIO
from datetime import datetime, timezone
from urllib.parse import urljoin
from typing import Optional
import pandas as pd, requests
from bs4 import BeautifulSoup
import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

# ----------------- CONFIG -----------------
st.set_page_config(page_title="Amazon Top-20 → Excel Export", layout="wide")
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
]
EXCEL_COL_WIDTH = 24
EXCEL_IMG_MAX_PX = 200          # max image pixel size (longest side)
EXCEL_IMG_ROW_HEIGHT = 82       # row height to fit the image row nicely
RANK_FILL = PatternFill("solid", fgColor="C45500")
WHITE_FONT = Font(color="FFFFFF", bold=True)
SAVED_DIR = ".saved_searches"; os.makedirs(SAVED_DIR, exist_ok=True)

def _session():
    s = requests.Session()
    s.headers.update({"Accept-Language": "en-US,en;q=0.9"})
    return s

def get_soup(url, s=None):
    s = s or _session()
    r = s.get(url, headers={"User-Agent": random.choice(USER_AGENTS)}, timeout=15)
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser")

def utc_now(): 
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

# -------- AMAZON SCRAPING ----------
def extract_asin(url):
    for r in [r"/dp/([A-Z0-9]{10})", r"/gp/product/([A-Z0-9]{10})", r"[?&]ASIN=([A-Z0-9]{10})"]:
        m = re.search(r, url, re.I)
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
            "URL": urljoin(url, href)
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
    # Look for "... bought in past month"
    pat = re.compile(r"bought.*past\s+month", re.I)
    for el in soup.find_all(["div", "span"], string=pat):
        txt = el.get_text(" ", strip=True)
        if re.search(r"\b\d", txt):
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

def fetch_details(url, s=None):
    s = s or _session()
    try:
        soup = get_soup(url, s)
        return extract_price(soup), extract_image(soup), extract_sellthrough(soup)
    except Exception:
        return "", "", ""

# ---------- SAVE / LOAD ----------
def _p(id): return os.path.join(SAVED_DIR, id)
def _d(id): return os.path.join(_p(id), "data.csv")
def _m(id): return os.path.join(_p(id), "meta.json")
def list_saved():
    runs=[]
    for rid in os.listdir(SAVED_DIR):
        try:
            mp=_m(rid); dp=_d(rid)
            if os.path.exists(mp) and os.path.exists(dp):
                meta=json.load(open(mp))
                runs.append({"id":rid, **meta})
        except: 
            pass
    return sorted(runs, key=lambda r: r.get("updated",""), reverse=True)
def save_run(id, df, meta):
    os.makedirs(_p(id), exist_ok=True)
    df.to_csv(_d(id), index=False)
    meta["updated"] = utc_now()
    json.dump(meta, open(_m(id), "w"), indent=2)
def load_run(id):
    return pd.read_csv(_d(id)), json.load(open(_m(id)))
def meta_new(name, url):
    now = utc_now()
    return {"name": name, "url": url, "created": now, "updated": now}

# -------- STREAMLIT UI ------------
if "df" not in st.session_state: st.session_state.df=None
if "meta" not in st.session_state: st.session_state.meta=None

st.sidebar.header("Saved Searches")
runs = list_saved()
for r in runs:
    st.sidebar.markdown(f"**{r['name']}**  \n<small>{r.get('updated','')}</small>", unsafe_allow_html=True)
    c1, c2 = st.sidebar.columns(2)
    if c1.button("Load", key=f"load{r['id']}"):
        df, m = load_run(r["id"])
        st.session_state.df = df
        st.session_state.meta = m
        st.success("Loaded.")
    if c2.button("Delete", key=f"del{r['id']}"):
        import shutil
        shutil.rmtree(_p(r["id"]), ignore_errors=True)
        st.rerun()
st.sidebar.divider()
st.sidebar.caption("Searches are saved locally.")

with st.sidebar.expander("⚙️ Settings", expanded=False):
    delay_min = st.slider("Min delay (sec)", 0.5, 3.0, 1.0, 0.1)
    delay_max = st.slider("Max delay (sec)", 0.6, 4.0, 2.0, 0.1)

st.title("Amazon Top-20 → Excel Exporter")

url = st.text_input("Amazon Best Sellers URL")
desc = st.text_input("Category Description", placeholder="e.g. Single Board Computers")
mode = st.radio("View Mode:", ["List","Grid","Compact"], horizontal=True)

c1, c2 = st.columns([1,3])
if c1.button("Fetch Top 20", type="primary"):
    if not re.match(r"^https?://", (url or "").strip(), re.I):
        st.error("Please enter a valid URL starting with http:// or https://")
    else:
        s = _session()
        try:
            items = parse_top20(url.strip(), s)
        except Exception as e:
            st.error(f"Failed: {e}")
            items = []
        out=[]
        for i, it in enumerate(items, 1):
            price, img, sell = fetch_details(it["URL"], s)
            out.append({
                "Rank": i, "Title": it["Title"], "ASIN": it["ASIN"], "URL": it["URL"],
                "Price": price, "Image": img, "Sell": sell,
                "MCSKU":"", "MCTitle":"", "MCRetail":"", "MCCost":"", "Avg1_4":"", "Attributes":"", "Notes":""
            })
            time.sleep(random.uniform(delay_min, delay_max))
        df = pd.DataFrame(out)
        st.session_state.df = df
        st.session_state.meta = meta_new(desc or "Top 20", url)
        st.success("Top 20 fetched!")

# ====== RENDER MODES ======
if st.session_state.df is not None:
    df = st.session_state.df.copy()
    if mode == "List":
        for _, r in df.iterrows():
            st.markdown(f"### #{r['Rank']}  {r['Title']}")
            cols = st.columns([1,3])
            with cols[0]:
                if r["Image"]:
                    st.image(r["Image"], width=120)
            with cols[1]:
                st.write(r["Price"])
                if r["Sell"]: st.caption(r["Sell"])
                st.write(f"[Open on Amazon]({r['URL']})")
            st.divider()
    else:
        ncols=5; rows=[df.iloc[i:i+ncols] for i in range(0, len(df), ncols)]
        card_css = """
        <style>
        .card{border:1px solid #ddd;padding:6px;border-radius:8px;margin:3px;box-shadow:0 0 4px rgba(0,0,0,0.1);}
        .card:hover{box-shadow:0 0 8px rgba(0,0,0,0.2);}
        .title{font-weight:600;font-size:0.9rem}
        .price{color:#B12704;font-weight:600;}
        .sell{font-size:0.8rem;color:#555;}
        .rankbar{background:#C45500;color:#fff;padding:2px 6px;border-radius:4px;display:inline-block;font-weight:700;}
        </style>
        """
        st.markdown(card_css, unsafe_allow_html=True)
        imgw = 120 if mode == "Grid" else 80
        for row in rows:
            cols = st.columns(ncols)
            for c, (_, r) in zip(cols, row.iterrows()):
                html = f"<div class='card'><div class='rankbar'>#{r['Rank']}</div><br>"
                if r["Image"]: html += f"<img src='{r['Image']}' width='{imgw}'/><br>"
                html += f"<div class='title'>{r['Title']}</div>"
                html += f"<div class='price'>{r['Price']}</div>"
                if r["Sell"]: html += f"<div class='sell'>{r['Sell']}</div>"
                html += f"<a href='{r['URL']}' target='_blank'>Open on Amazon</a></div>"
                c.markdown(html, unsafe_allow_html=True)
            st.write("")

    # ====== EXPORT (Top 20 + Data) ======
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

        # ---- Data sheet columns (fixed order)
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

        # ---- Top 20 layout
        wrap_top = Alignment(wrap_text=True, vertical="top")
        thin = Side(style="thin", color="DDDDDD")
        box = Border(left=thin, right=thin, top=thin, bottom=thin)

        ITEMS_PER_ROW = 5
        ROW_BLOCK = 11  # rank + image + 8 info rows + spacer if needed

        # column widths
        for c in range(1, ITEMS_PER_ROW+1):
            ws_top.column_dimensions[get_column_letter(c)].width = EXCEL_COL_WIDTH

        def data_ref(col_idx, drow):
            return f"Data!{get_column_letter(col_idx)}{drow}"

        for idx in range(min(20, len(df))):
            group = idx // ITEMS_PER_ROW
            col = (idx % ITEMS_PER_ROW) + 1
            base = group * ROW_BLOCK + 1
            drow = idx + 2  # corresponding row in Data

            # Rank bar
            rank_cell = ws_top.cell(row=base, column=col, value=f"Rank #{df.at[idx, 'Rank']}")
            rank_cell.fill = RANK_FILL
            rank_cell.font = WHITE_FONT
            rank_cell.alignment = Alignment(horizontal="center", vertical="center")
            rank_cell.border = box

            # Image row
            ws_top.row_dimensions[base+1].height = EXCEL_IMG_ROW_HEIGHT
            img_buf = _download_image_bytes(df.at[idx, "Image"], EXCEL_IMG_MAX_PX)
            if img_buf:
                img = XLImage(img_buf)
                # Anchor at the top-left of the image row cell
                img.anchor = f"{get_column_letter(col)}{base+1}"
                ws_top.add_image(img)
            # put a space so the cell exists and keeps border/background consistent
            img_cell = ws_top.cell(row=base+1, column=col, value=" ")
            img_cell.border = box

            # Info lines (single cell with formula per line)
            info_lines = [
                ("Amazon",    data_ref(5, drow)),    # Title
                ("Price",     data_ref(6, drow)),
                ("Sell",      data_ref(7, drow)),
                ("MC SKU",    data_ref(8, drow)),
                ("MC Title",  data_ref(9, drow)),
                ("MC Retail", data_ref(10, drow)),
                ("MC Cost",   data_ref(11, drow)),
                ("1-4 Avg",   data_ref(12, drow)),
                ("Attributes",data_ref(13, drow)),
                ("Notes",     data_ref(14, drow)),
            ]
            for r_off, (label, ref) in enumerate(info_lines, start=2):  # start at base+2
                cell = ws_top.cell(row=base + r_off, column=col)
                cell.value = f'=CONCAT("{label}: ", {ref})'
                cell.alignment = wrap_top
                cell.border = box

        bio = BytesIO()
        wb.save(bio)
        bio.seek(0)
        return bio.read()

    st.download_button(
        "📥 Download Excel",
        data=build_xlsx(st.session_state.df),
        file_name=f"{(st.session_state.meta['name']).replace(' ','_')}_Top20.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
