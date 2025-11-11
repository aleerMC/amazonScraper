import re
import os
import json
import uuid
import time
import random
from io import BytesIO
from datetime import datetime, timezone
from urllib.parse import urljoin, quote_plus, urlparse
from typing import Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup
import streamlit as st

# ====== Excel builder (embed images; compact) ======
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

# ---------------- Excel layout knobs (compact) ----------------
EXCEL_COL_WIDTH = 24          # width per item column
EXCEL_IMG_MAX_PX = 220        # max embedded image size (longest side)
EXCEL_IMG_ROW_HEIGHT = 90     # image row height (compact)
BAND_COLOR_1 = "FFFFFF"       # white
BAND_COLOR_2 = "F7F7F7"       # light gray
LABEL_BOLD = True             # bold for label lines

# ========================= Setup / Helpers =========================

st.set_page_config(page_title="Amazon Top-20 → Micro Center Matcher (Persistent)", layout="wide")

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
]

SAVED_DIR = ".saved_searches"
os.makedirs(SAVED_DIR, exist_ok=True)

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

def utc_now_str():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

# ========================= Amazon Scraping =========================

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
    """Collect the first 20 unique ASINs on the Best Sellers page.
       Also resolve intermediary redirect links to the final product URL."""
    session = session or _session()
    soup, final_url = get_soup(url, session)
    anchors = soup.find_all("a", href=True)
    seen_asins, items = set(), []
    for a in anchors:
        href = a.get("href", "")
        asin = extract_asin_from_url(href)
        if not asin or asin in seen_asins:
            continue

        # best-effort title
        title = (a.get_text(strip=True) or "").strip()
        if not title:
            img = a.find("img", alt=True)
            if img and img.get("alt"):
                title = img["alt"].strip()
        if not title and a.get("title"):
            title = a["title"].strip()

        # absolute URL
        item_url = urljoin(final_url, href) if href.startswith("/") else href

        # --- Resolve redirects (Amazon shortlinks) ---
        try:
            r = session.get(item_url, allow_redirects=True, timeout=10,
                            headers={"User-Agent": random.choice(USER_AGENTS)})
            item_url = r.url
        except Exception:
            pass

        seen_asins.add(asin)
        items.append({"ASIN": asin, "Title": title or "", "URL": item_url})
        if len(items) >= 20:
            break
    return items

def extract_price_from_soup_amzn(soup):
    candidate_ids = [
        "priceblock_ourprice", "priceblock_dealprice", "priceblock_saleprice",
        "priceblock_pospromoprice", "sns-base-price", "corePrice_feature_div",
        "apex_desktop", "priceToPay"
    ]
    for cid in candidate_ids:
        el = soup.find(id=cid)
        if el:
            off = el.find("span", class_="a-offscreen")
            if off and off.get_text(strip=True):
                return off.get_text(strip=True)
            txt = el.get_text(" ", strip=True)
            m = re.search(r"\$\s*\d[\d,]*(?:\.\d{2})?", txt or "")
            if m:
                return m.group(0).replace(" ", "")
    for off in soup.select("span.a-offscreen"):
        val = off.get_text(strip=True)
        if re.match(r"^\$\s*\d", val):
            return val.replace(" ", "")
    meta = soup.find("meta", attrs={"itemprop": "price"})
    if meta and meta.get("content"):
        c = meta.get("content")
        return c if c.startswith("$") else f"${c}"
    og = soup.find("meta", attrs={"property": "og:price:amount"})
    if og and og.get("content"):
        c = og.get("content")
        return c if c.startswith("$") else f"${c}"
    return ""

def extract_image_from_soup_amzn(soup):
    og = soup.find("meta", attrs={"property": "og:image"})
    if og and og.get("content"):
        return og["content"]
    link_img = soup.find("link", rel=lambda v: v and "image_src" in v)
    if link_img and link_img.get("href"):
        return link_img.get("href")
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

# --- improved: add random delay and retries for more stable images ---
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

# --- cache image bytes to reduce dropped images on Streamlit Cloud ---
@st.cache_data(ttl=3600)
def load_image_bytes(url: str):
    try:
        if not url:
            return None
        resp = requests.get(url, headers={"User-Agent": random.choice(USER_AGENTS)}, timeout=10)
        resp.raise_for_status()
        return resp.content
    except Exception:
        return None

# ========================= Micro Center (robust candidates & pricing) =========================

_SERVICE_PAT = re.compile(
    r"/(service|services|repair|repairs|battery|batteries|appointment|warranty|diagnostic|trade-in|protection|tech-support)\b",
    re.I
)
def _is_service_like_url(url: str) -> bool:
    try:
        p = urlparse(url)
        path = (p.path or "").lower()
        if _SERVICE_PAT.search(path):
            return True
        if "/in-store-service" in path or "/data-recovery" in path or "/apple-" in path:
            return True
        return False
    except Exception:
        return False

def _normalize_tokens(s: str):
    s = (s or "").lower()
    s = re.sub(r"[^a-z0-9\s\-\+_/\.]", " ", s)
    return [t for t in re.split(r"\s+", s) if t]

def _score_candidate(query: str, pdict: dict) -> float:
    q_tokens = set(_normalize_tokens(query))
    fields = " ".join([
        pdict.get("MCTitle",""),
        pdict.get("MCBrand",""),
        pdict.get("MCModel",""),
        pdict.get("MCDescription","")
    ])
    c_tokens = set(_normalize_tokens(fields))
    if not q_tokens or not c_tokens:
        return 0.0
    inter = len(q_tokens & c_tokens)
    union = len(q_tokens | c_tokens)
    score = inter / max(union, 1)
    model = (pdict.get("MCModel","") or "").lower()
    if model and model in " ".join(q_tokens):
        score += 0.15
    return float(min(score, 1.0))

PRICE_CTX_CURRENT = re.compile(r"\b(your|current|now|sale|today)\b", re.I)
PRICE_CTX_RETAIL  = re.compile(r"\b(was|reg|regular|compare|strike|list)\b", re.I)
PRICE_BAD_CTX     = re.compile(r"\b(per\s*month|plan|protection|service|replacement|battery|warranty)\b", re.I)

def _clean_money(s: str) -> str:
    if not isinstance(s, str):
        return ""
    m = re.search(r"\$?\s*(\d[\d,]*(?:\.\d{2})?)", s.replace("\xa0"," ").strip())
    if not m:
        return ""
    amt = m.group(1).replace(",", "")
    return f"${amt}"

def _collect_price_candidates(psoup: BeautifulSoup):
    cands = []
    for script in psoup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "{}")
            objs = data if isinstance(data, list) else [data]
            for obj in objs:
                if isinstance(obj, dict) and str(obj.get("@type","")).lower() == "product":
                    offers = obj.get("offers")
                    if isinstance(offers, dict):
                        p = offers.get("price")
                        if p:
                            cands.append((_clean_money(str(p)), "jsonld offers", "jsonld"))
        except Exception:
            pass
    for el in psoup.find_all(attrs={"itemprop": "price"}):
        txt = el.get("content") or el.get_text(" ", strip=True)
        val = _clean_money(txt or "")
        if val:
            ctx = el.get_text(" ", strip=True)
            if not PRICE_BAD_CTX.search(ctx):
                cands.append((val, ctx, "itemprop=price"))
    for el in psoup.find_all(["span","div"], class_=re.compile(r"(price|yourprice|current|was|reg|strike)", re.I)):
        ctx = el.get_text(" ", strip=True)
        if not ctx or PRICE_BAD_CTX.search(ctx):
            continue
        m = re.search(r"\$\s*\d[\d,]*(?:\.\d{2})?", ctx)
        if m:
            val = _clean_money(m.group(0))
            cands.append((val, ctx, "class=price*"))
    body_txt = psoup.get_text(" ", strip=True)
    if not cands:
        for m in re.finditer(r"\$\s*\d[\d,]*(?:\.\d{2})?", body_txt):
            snip = body_txt[max(0, m.start()-25): m.end()+25]
            if PRICE_BAD_CTX.search(snip):
                continue
            cands.append((_clean_money(m.group(0)), snip, "body-scan"))
            if len(cands) >= 3:
                break
    return cands

def _pick_prices(cands):
    if not cands:
        return "", ""
    def score_current(ctx: str) -> int:
        return 2 if PRICE_CTX_CURRENT.search(ctx or "") else 0
    def score_retail(ctx: str) -> int:
        return 2 if PRICE_CTX_RETAIL.search(ctx or "") else 0
    def to_float(val: str) -> float:
        try:
            return float(val.replace("$",""))
        except Exception:
            return -1.0

    current_cands, retail_cands, unknown_cands = [], [], []
    for (val, ctx, src) in cands:
        sc_c = score_current(ctx)
        sc_r = score_retail(ctx)
        fval = to_float(val)
        if sc_c > sc_r:
            current_cands.append((val, sc_c, fval, ctx, src))
        elif sc_r > sc_c:
            retail_cands.append((val, sc_r, fval, ctx, src))
        else:
            unknown_cands.append((val, 1, fval, ctx, src))

    if not current_cands and unknown_cands:
        current_cands = unknown_cands

    current_cands.sort(key=lambda x: (x[1], x[2]), reverse=True)
    retail_cands.sort(key=lambda x: (x[1], x[2]), reverse=True)

    cur = current_cands[0][0] if current_cands else ""
    ret = retail_cands[0][0]  if retail_cands  else ""

    def to_num(s):
        try: return float(s.replace("$",""))
        except: return None
    n_cur, n_ret = to_num(cur), to_num(ret)
    if n_cur is not None and n_ret is not None and n_cur > n_ret:
        cur, ret = ret, cur

    return cur, ret

def _extract_mc_prices(psoup: BeautifulSoup, debug: bool=False):
    cands = _collect_price_candidates(psoup)
    cur, ret = _pick_prices(cands)
    if debug:
        st.caption(f"MC price debug → candidates={len(cands)}, picked current={cur or '—'}, retail={ret or '—'}")
    return cur, ret

def _parse_mc_product_page(url, session=None, debug: bool=False):
    session = session or _session()
    headers = {"User-Agent": random.choice(USER_AGENTS)}
    r = session.get(url, headers=headers, timeout=20)
    r.raise_for_status()
    psoup = BeautifulSoup(r.text, "html.parser")

    detected_sku = ""
    sku_meta = psoup.find(attrs={"itemprop": "sku"})
    if sku_meta:
        detected_sku = sku_meta.get("content") or sku_meta.get_text(strip=True) or ""

    json_ld_sku = ""
    product_title = ""
    product_image = ""
    product_brand = ""
    product_model = ""
    product_desc = ""
    product_price = ""
    product_retail = ""
    is_product_schema = False

    for script in psoup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "{}")
            objs = data if isinstance(data, list) else [data]
            for obj in objs:
                if isinstance(obj, dict):
                    ty = str(obj.get("@type", "")).lower()
                    if ty == "product":
                        is_product_schema = True
                        json_ld_sku = obj.get("sku") or json_ld_sku
                        product_title = obj.get("name") or product_title
                        product_desc = obj.get("description") or product_desc
                        if isinstance(obj.get("brand"), dict):
                            product_brand = obj["brand"].get("name") or product_brand
                        if isinstance(obj.get("image"), str):
                            product_image = obj.get("image") or product_image
                        break
        except Exception:
            pass

    if not (is_product_schema or detected_sku or json_ld_sku):
        title_txt = (psoup.find("title").get_text(strip=True) if psoup.find("title") else "").lower()
        if "service" in title_txt or "repair" in title_txt or "battery" in title_txt:
            raise ValueError("Service page, not product")

    if not detected_sku:
        text = psoup.get_text(" ", strip=True)
        m = re.search(r"\bSKU\s*[:#]?\s*(\d+)\b", text, re.I)
        if m:
            detected_sku = m.group(1)

    if not product_title:
        h1 = psoup.find("h1")
        if h1:
            product_title = h1.get_text(strip=True)

    if not product_image:
        og = psoup.find("meta", {"property": "og:image"})
        if og and og.get("content"):
            product_image = og["content"]

    cur, reg = _extract_mc_prices(psoup, debug=debug)
    product_price  = product_price or cur
    product_retail = product_retail or reg

    text = psoup.get_text(" ", strip=True)
    if not product_brand:
        bb = re.search(r"\bBrand\s*[:#]?\s*([A-Za-z0-9\-_/\. ]+)", text, re.I)
        if bb:
            product_brand = bb.group(1).strip()
    if not product_model:
        mm = re.search(r"\bModel\s*[:#]?\s*([A-Za-z0-9\-_/\.]+)", text, re.I)
        if mm:
            product_model = mm.group(1)

    return {
        "detected_sku": (json_ld_sku or detected_sku).strip(),
        "MCTitle": product_title or "",
        "MCPrice": product_price or "",
        "MCRetail": product_retail or "",
        "MCImageURL": product_image or "",
        "MCDescription": product_desc or "",
        "MCModel": product_model or "",
        "MCBrand": product_brand or "",
        "MCURL": url,
    }

@st.cache_data(ttl=1800, show_spinner=False)
def fetch_microcenter_candidates(q: str, limit: int = 8, debug: bool=False):
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

        cand_links, seen = [], set()
        for a in soup.find_all("a", href=True):
            href = a["href"] or ""
            href_abs = href if href.startswith("http") else ("https://www.microcenter.com" + href)
            low = href_abs.lower()
            if "/product/" not in low:
                continue
            if _is_service_like_url(href_abs):
                continue
            if href_abs not in seen:
                seen.add(href_abs)
                cand_links.append(href_abs)

        parsed = []
        for link in cand_links:
            try:
                pdict = _parse_mc_product_page(link, session, debug=debug)
                if not (pdict.get("detected_sku") or pdict.get("MCTitle")):
                    continue
                pdict["_score"] = _score_candidate(q, pdict)
                parsed.append(pdict)
            except Exception:
                continue

        num_q = re.fullmatch(r"\d{4,}", q) is not None
        exact, others = [], []
        for p in parsed:
            if num_q and p.get("detected_sku") == q:
                exact.append(p)
            else:
                others.append(p)
        others = [p for p in others if p.get("_score", 0.12) >= 0.12]
        others.sort(key=lambda x: x.get("_score", 0.0), reverse=True)
        results = exact + others
        return results[:limit]
    except Exception:
        return []

# ========================= Persistence (CSV, no pyarrow) =========================

def _run_path(run_id: str) -> str:
    return os.path.join(SAVED_DIR, run_id)

def _meta_path(run_id: str) -> str:
    return os.path.join(_run_path(run_id), "meta.json")

def _data_path(run_id: str) -> str:
    return os.path.join(_run_path(run_id), "data.csv")

def list_saved_runs():
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

def new_run_meta(category_desc: str, amazon_url: str, fetched_at: str, name: Optional[str] = None):
    now = utc_now_str()
    return {
        "name": name or (category_desc or "Top 20"),
        "category_desc": category_desc or "",
        "amazon_url": amazon_url or "",
        "fetched_at": fetched_at or now,
        "created_at": now,
        "updated_at": now
    }

# ========================= Session State =========================

if "results" not in st.session_state:
    st.session_state.results = None
if "mc_cands" not in st.session_state:
    st.session_state.mc_cands = {}
if "current_run_id" not in st.session_state:
    st.session_state.current_run_id = None
if "current_meta" not in st.session_state:
    st.session_state.current_meta = None
if "autosave" not in st.session_state:
    st.session_state.autosave = True

# ========================= Sidebar =========================

with st.sidebar:
    st.header("Saved Searches")

    st.session_state.autosave = st.checkbox("Autosave edits", value=st.session_state.autosave)

    saved = list_saved_runs()
    if saved:
        labels = [f"{r['name']}  —  {r.get('updated_at','')}" for r in saved]
        idx = st.selectbox("Select a saved search", options=list(range(len(saved))),
                           format_func=lambda i: labels[i])
        sel = saved[idx]

        c1, c2, c3, c4 = st.columns([1,1,1,1])
        with c1:
            if st.button("Load"):
                df, meta = load_run(sel["id"])
                st.session_state.results = df
                st.session_state.current_run_id = sel["id"]
                st.session_state.current_meta = meta
                st.success(f"Loaded: {meta.get('name')}")

        with c2:
            if st.button("Save", disabled=st.session_state.results is None):
                if st.session_state.current_run_id and st.session_state.results is not None:
                    save_run(st.session_state.current_run_id, st.session_state.results, st.session_state.current_meta)
                    st.success("Saved.")

        with c3:
            new_name = st.text_input("Rename", value=sel["name"], label_visibility="collapsed")
            if st.button("Apply"):
                df, meta = load_run(sel["id"])
                meta["name"] = new_name.strip() or meta["name"]
                save_run(sel["id"], df, meta)
                st.success("Renamed. Refreshing list…")
                st.rerun()

        with c4:
            if st.button("Delete"):
                try:
                    import shutil
                    shutil.rmtree(_run_path(sel["id"]), ignore_errors=True)
                    st.warning("Deleted. Refreshing list…")
                    st.rerun()
                except Exception as e:
                    st.error(f"Delete failed: {e}")
    else:
        st.caption("No saved searches yet.")

    with st.expander("Scrape Options", expanded=False):
        delay_min = st.slider("Min per-item delay (sec)", 0.5, 5.0, 1.0, 0.1, key="dmin")
        delay_max = st.slider("Max per-item delay (sec)", 0.6, 6.0, 2.0, 0.1, key="dmax")
        mc_price_debug = st.checkbox("Debug MC price parse", value=False, help="Show how MC prices were detected")

# ========================= Top Inputs & Actions =========================

st.title("Amazon Top-20 → Micro Center Matcher (Persistent)")

col_in1, col_in2 = st.columns([3, 2])
with col_in1:
    amz_url = st.text_input("Amazon Best Sellers URL", placeholder="https://www.amazon.com/gp/bestsellers/pc/....")
with col_in2:
    category_desc = st.text_input("Short Description (for your reference)", placeholder="e.g., Single Board Computers")

col_btn = st.columns([1, 2, 2, 2])
with col_btn[0]:
    fetch_btn = st.button("Fetch Top 20", type="primary")
with col_btn[1]:
    save_new_btn = st.button("Save as New", disabled=(st.session_state.results is None))
with col_btn[2]:
    save_btn = st.button("Save Changes", disabled=(st.session_state.results is None or st.session_state.current_run_id is None))
with col_btn[3]:
    export_placeholder = st.empty()

# ========================= Fetch Handler =========================

if fetch_btn:
    if not amz_url.strip():
        st.error("Please enter an Amazon Best Sellers URL.")
    else:
        session = _session()
        st.info(f"Fetching top 20 from: {amz_url}")
        try:
            items = parse_top20_from_category_page(amz_url.strip(), session)
        except Exception as e:
            st.error(f"Failed to fetch category page: {e}")
            items = []

        out_rows = []
        ts = utc_now_str()
        total = len(items)
        progress = st.progress(0) if total else None
        status = st.empty()

        for idx, it in enumerate(items):
            rank = idx + 1
            status.write(f"Scraping {rank} of {total}…")
            price, image = fetch_item_details_amzn(it.get("URL", ""), session)
            out_rows.append({
                "ImageURL": image,
                "Rank": rank,
                "Title": it.get("Title", ""),
                "ASIN": it.get("ASIN", ""),
                "AmazonURL": it.get("URL", ""),
                "AmazonPrice": price,
                "MCSKU": "",
                "MCPrice": "",
                "MCRetail": "",
                "MCTitle": "",
                "MCImageURL": "",
                "MCDescription": "",
                "MCModel": "",
                "MCBrand": "",
                "MCURL": "",
                "AttrMatch": "",
                "Notes": "",
                "MCCost": "",
                "Avg1_4": "",
                "FetchedAt": ts,
                "CategoryDesc": category_desc.strip(),
                "AmazonBestURL": amz_url.strip(),
            })
            if progress:
                progress.progress(int(((idx+1)/max(total,1))*100))
            time.sleep(random.uniform(st.session_state.dmin, st.session_state.dmax))

        if progress:
            progress.empty()
        status.success(f"Done! Fetched {len(out_rows)} items.")

        df = pd.DataFrame(out_rows, columns=[
            "ImageURL","Rank","Title","ASIN","AmazonURL","AmazonPrice",
            "MCSKU","MCPrice","MCRetail","MCTitle","MCImageURL","MCDescription","MCModel","MCBrand","MCURL",
            "AttrMatch","Notes","MCCost","Avg1_4","FetchedAt","CategoryDesc","AmazonBestURL"
        ])
        st.session_state.results = df
        st.session_state.current_run_id = None
        st.session_state.current_meta = None

# ========================= Save Handlers =========================

if save_new_btn and st.session_state.results is not None:
    run_id = uuid.uuid4().hex[:12]
    meta = new_run_meta(
        category_desc=st.session_state.results.iloc[0].get("CategoryDesc",""),
        amazon_url=st.session_state.results.iloc[0].get("AmazonBestURL",""),
        fetched_at=st.session_state.results.iloc[0].get("FetchedAt",""),
        name=(st.session_state.results.iloc[0].get("CategoryDesc") or "Top 20")
    )
    save_run(run_id, st.session_state.results, meta)
    st.session_state.current_run_id = run_id
    st.session_state.current_meta = meta
    st.success(f"Saved as new: {meta['name']}")
    st.rerun()

if save_btn and st.session_state.results is not None and st.session_state.current_run_id:
    save_run(st.session_state.current_run_id, st.session_state.results, st.session_state.current_meta)
    st.success("Changes saved.")

def maybe_autosave():
    if st.session_state.autosave and st.session_state.current_run_id and st.session_state.results is not None:
        save_run(st.session_state.current_run_id, st.session_state.results, st.session_state.current_meta)

# ========================= Results UI (3 columns per item) =========================

if st.session_state.results is not None and not st.session_state.results.empty:
    meta = st.session_state.current_meta or {
        "name": st.session_state.results.iloc[0].get("CategoryDesc") or "Top 20",
        "fetched_at": st.session_state.results.iloc[0].get("FetchedAt"),
        "amazon_url": st.session_state.results.iloc[0].get("AmazonBestURL")
    }
    header_cols = st.columns([3,2,2])
    with header_cols[0]:
        st.subheader(f"{meta.get('name','Top 20')}")
    with header_cols[1]:
        st.caption(f"Pulled at: {meta.get('fetched_at','')}")
    with header_cols[2]:
        if meta.get("amazon_url"):
            st.caption(f"[Source: Amazon Best Sellers]({meta['amazon_url']})")

    st.markdown("""
        <style>
        .tight p { margin: 0.1rem 0 !important; }
        .tight .stMarkdown { line-height: 1.1; }
        .tiny { font-size: 0.85rem; color: #6b7280; }
        .price { font-weight: 600; }
        .rowrule { border-top: 1px solid #e5e7eb; margin: .25rem 0 .75rem 0; }
        </style>
    """, unsafe_allow_html=True)

    data = st.session_state.results.copy()

    for i, row in data.iterrows():
        colA, colB, colC = st.columns([1, 1, 1])

        # === Column 1: Amazon info ===
        with colA:
            a_cols = st.columns([1, 3])
            with a_cols[0]:
                if isinstance(row["ImageURL"], str) and row["ImageURL"]:
                    img_data = load_image_bytes(row["ImageURL"])
                    if img_data:
                        st.image(img_data, width=120)
                    else:
                        st.write("—")
                else:
                    st.write("—")
            with a_cols[1]:
                st.markdown(f"**#{int(row['Rank'])} Amazon**")
                st.markdown(
                    f"<div class='tight' style='text-align:left'>{row['Title'] or '(no title)'}<br>"
                    f"<span class='price'>{row['AmazonPrice'] or ''}</span><br>"
                    f"<span class='tiny'>ASIN: {row['ASIN']}</span></div>",
                    unsafe_allow_html=True
                )
                if row.get("AmazonURL"):
                    st.markdown(f"[Amazon Link]({row['AmazonURL']})")

        # === Column 2: Micro Center info ===
        with colB:
            disp = data.loc[i]
            b_cols = st.columns([1, 3])
            with b_cols[0]:
                mc_img_url = disp.get("MCImageURL", "")
                if isinstance(mc_img_url, str) and mc_img_url:
                    mc_img = load_image_bytes(mc_img_url)
                    if mc_img:
                        st.image(mc_img, width=120)
                    else:
                        st.write("—")
                else:
                    st.write("—")
            with b_cols[1]:
                st.markdown("**Micro Center**")
                sku_disp = disp.get("MCSKU","") or "—"
                title_disp = disp.get("MCTitle","") or "—"
                current_price = disp.get("MCPrice","") or ""
                retail_price = disp.get("MCRetail","") or ""
                mcurl = disp.get("MCURL","")

                st.markdown(f"<div class='tight' style='text-align:left'>SKU: <b>{sku_disp}</b></div>", unsafe_allow_html=True)
                st.markdown(f"<div class='tight' style='text-align:left'>{title_disp}</div>", unsafe_allow_html=True)
                if current_price:
                    st.markdown(f"<div class='tight' style='text-align:left'>Current: <span class='price'>{current_price}</span></div>", unsafe_allow_html=True)
                if retail_price:
                    st.markdown(f"<div class='tiny' style='text-align:left'>Reg: {retail_price}</div>", unsafe_allow_html=True)
                if mcurl:
                    st.markdown(f"[MC Link]({mcurl})")

                row_key = f"{i}"
                attr_key = f"attr_{row_key}"
                notes_key = f"notes_{row_key}"
                attr_val = st.text_input("Attribute Match", value=disp.get("AttrMatch",""), key=attr_key,
                                         placeholder="e.g., 8GB / 128GB / BT5.3 / IPX7")
                note_val = st.text_input("Notes", value=disp.get("Notes",""), key=notes_key,
                                         placeholder="e.g., don't carry / in process")
                data.at[i, "AttrMatch"] = attr_val
                data.at[i, "Notes"] = note_val

        # === Column 3: Search & Submit ===
        with colC:
            st.markdown("**Find & Submit**")
            row_key = f"{i}"
            search_key = f"search_{row_key}"
            dosearch_key = f"dosearch_{row_key}"
            pick_key = f"pick_{row_key}"
            submit_key = f"submit_{row_key}"
            price_override_key = f"price_{row_key}"

            default_q = row.get("MCSKU","") or ""
            query = st.text_input("Search MC (SKU or words)", value=default_q, key=search_key, placeholder="e.g., 123456 or 'Raspberry Pi 5'")
            do_search = st.button("Search", key=dosearch_key)

            cands_store = st.session_state.mc_cands
            prev = cands_store.get(row_key, {})
            need_fetch = do_search or (query.strip() and query.strip() != (prev.get("sku","") or ""))

            if need_fetch:
                cand_list = fetch_microcenter_candidates(query.strip(), limit=8, debug=mc_price_debug) if query.strip() else []
                def label_of(c):
                    sku = c.get("detected_sku") or "—"
                    title = c.get("MCTitle") or "(no title)"
                    price_lbl = c.get("MCRetail") or c.get("MCPrice") or ""
                    return f"{sku} • {title[:60]}{'…' if len(title)>60 else ''} {('— ' + price_lbl) if price_lbl else ''}"
                labels = [label_of(c) for c in cand_list]
                selected_idx = 0 if cand_list else -1
                cands_store[row_key] = {"sku": query.strip(), "cands": cand_list, "labels": labels, "selected": selected_idx}

            store = st.session_state.mc_cands.get(row_key, {})
            cand_list = store.get("cands", [])
            labels = store.get("labels", [])
            selected_idx = store.get("selected", 0 if cand_list else -1)

            if cand_list:
                new_idx = st.selectbox("Candidates", options=list(range(len(cand_list))),
                                       index=max(0, selected_idx), format_func=lambda k: labels[k], key=pick_key)
                st.session_state.mc_cands[row_key]["selected"] = new_idx

                cur_price_val = data.at[i, "MCPrice"] if isinstance(data.at[i, "MCPrice"], str) else ""
                new_price_val = st.text_input("Override Current Price", value=cur_price_val, key=price_override_key)

                if st.button("Submit to Micro Center column", type="primary", key=submit_key):
                    chosen = cand_list[new_idx]
                    data.at[i, "MCSKU"] = (chosen.get("detected_sku","") or "").strip()
                    for k in ["MCTitle","MCPrice","MCRetail","MCImageURL","MCDescription","MCModel","MCBrand","MCURL"]:
                        data.at[i, k] = chosen.get(k, "")
                    if new_price_val:
                        data.at[i, "MCPrice"] = new_price_val

                    st.session_state.results = data
                    if st.session_state.current_run_id and st.session_state.autosave:
                        save_run(st.session_state.current_run_id, st.session_state.results, st.session_state.current_meta)
                    st.success("Submitted to Micro Center column.")
                    st.rerun()

        st.markdown("<div class='rowrule'></div>", unsafe_allow_html=True)

    st.session_state.results = data
    maybe_autosave()

    # ===== Export (Top 20 + Data) with EMBEDDED IMAGES (compact, banded, left-justified) =====

    def _download_image_bytes(url: str, max_px: int = EXCEL_IMG_MAX_PX) -> Optional[BytesIO]:
        """Fetch & resize image; return PNG bytes or None."""
        try:
            if not url or not isinstance(url, str):
                return None
            resp = requests.get(url, headers={"User-Agent": random.choice(USER_AGENTS)}, timeout=15)
            resp.raise_for_status()
            img = PILImage.open(BytesIO(resp.content)).convert("RGBA")
            img.load()  # force-load to avoid truncated JPEGs in some environments
            img.thumbnail((max_px, max_px), PILImage.LANCZOS)
            buff = BytesIO()
            img.save(buff, format="PNG")
            buff.seek(0)
            return buff
        except Exception:
            return None

    def build_xlsx_two_sheets(df: pd.DataFrame) -> bytes:
        wb = Workbook()
        ws_top = wb.active
        ws_top.title = "Top 20"
        ws_data = wb.create_sheet("Data")

        # ----- Data sheet
        cols = [
            "Rank","ASIN","AmazonURL","ImageURL","Title","AmazonPrice",
            "MCSKU","MCTitle","MCPrice","MCRetail","MCImageURL","MCURL",
            "AttrMatch","Notes","FetchedAt","CategoryDesc","AmazonBestURL",
            "MCCost","Avg1_4"
        ]
        for j, c in enumerate(cols, start=1):
            ws_data.cell(row=1, column=j, value=c)

        top20 = df.head(20).reset_index(drop=True)
        for i, r in top20.iterrows():
            values = [
                int(r.get("Rank", i+1)),
                r.get("ASIN",""),
                r.get("AmazonURL",""),
                r.get("ImageURL",""),
                r.get("Title",""),
                r.get("AmazonPrice",""),
                r.get("MCSKU",""),
                r.get("MCTitle",""),
                r.get("MCPrice",""),
                r.get("MCRetail",""),
                r.get("MCImageURL",""),
                r.get("MCURL",""),
                r.get("AttrMatch",""),
                r.get("Notes",""),
                r.get("FetchedAt",""),
                r.get("CategoryDesc",""),
                r.get("AmazonBestURL",""),
                r.get("MCCost",""),
                r.get("Avg1_4",""),
            ]
            for j, v in enumerate(values, start=1):
                ws_data.cell(row=2+i, column=j, value=v)

        data_widths = [8,12,30,30,50,12,14,50,12,12,30,30,20,30,20,24,32,12,12]
        for j, w in enumerate(data_widths, start=1):
            ws_data.column_dimensions[get_column_letter(j)].width = w
        ws_data.freeze_panes = "A2"

        # ----- Top 20 layout (compact & banded)
        ITEMS_PER_ROW = 5
        BLOCK_ROWS = 11
        START_ROW = 1
        START_COL = 1

        left_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)
        left_mid = Alignment(horizontal="left", vertical="center", wrap_text=True)
        center_mid = Alignment(horizontal="center", vertical="center", wrap_text=True)
        bold = Font(bold=LABEL_BOLD)
        thin = Side(style="thin", color="DDDDDD")
        box = Border(left=thin, right=thin, top=thin, bottom=thin)
        band1 = PatternFill("solid", fgColor=BAND_COLOR_1)
        band2 = PatternFill("solid", fgColor=BAND_COLOR_2)

        for c in range(START_COL, START_COL + ITEMS_PER_ROW):
            ws_top.column_dimensions[get_column_letter(c)].width = EXCEL_COL_WIDTH

        def DREF(col_idx, data_row):
            return f"Data!{get_column_letter(col_idx)}{data_row}"

        for idx in range(len(top20)):
            group = idx // ITEMS_PER_ROW
            col = START_COL + (idx % ITEMS_PER_ROW)
            base = START_ROW + group * BLOCK_ROWS
            data_row = 2 + idx
            band_fill = band1 if (group % 2 == 0) else band2

            img_url = ws_data.cell(row=data_row, column=4).value
            amazon_url = ws_data.cell(row=data_row, column=3).value

            img_buf = _download_image_bytes(img_url, max_px=EXCEL_IMG_MAX_PX)
            if img_buf:
                xl_img = XLImage(img_buf)
                xl_img.anchor = f"{get_column_letter(col)}{base + 0}"
                ws_top.add_image(xl_img)

            cell = ws_top.cell(row=base + 0, column=col, value=" ")
            if amazon_url:
                cell.hyperlink = amazon_url
            cell.alignment = center_mid
            cell.border = box
            cell.fill = band_fill
            ws_top.row_dimensions[base + 0].height = EXCEL_IMG_ROW_HEIGHT

            c = ws_top.cell(row=base + 1, column=col, value=f'= "Rank: #" & {DREF(1, data_row)}')
            c.alignment = left_mid; c.font = bold; c.border = box; c.fill = band_fill

            c = ws_top.cell(row=base + 2, column=col, value=f'= "Amazon: " & {DREF(5, data_row)}')
            c.alignment = left_wrap; c.border = box; c.fill = band_fill

            c = ws_top.cell(row=base + 3, column=col, value=f'= "Amazon Price: " & {DREF(6, data_row)}')
            c.alignment = left_mid; c.font = bold; c.border = box; c.fill = band_fill

            c = ws_top.cell(row=base + 4, column=col, value=f'= "MC SKU: " & {DREF(7, data_row)}')
            c.alignment = left_mid; c.border = box; c.fill = band_fill

            c = ws_top.cell(row=base + 5, column=col, value=f'= "MC Title: " & {DREF(8, data_row)}')
            c.alignment = left_wrap; c.border = box; c.fill = band_fill

            c = ws_top.cell(row=base + 6, column=col, value=f'= "MC Retail: " & {DREF(10, data_row)}')
            c.alignment = left_mid; c.font = bold; c.border = box; c.fill = band_fill

            c = ws_top.cell(row=base + 7, column=col, value=f'= "MC Cost: " & {DREF(18, data_row)}')
            c.alignment = left_mid; c.border = box; c.fill = band_fill

            c = ws_top.cell(row=base + 8, column=col, value=f'= "1-4 Avg: " & {DREF(19, data_row)}')
            c.alignment = left_mid; c.border = box; c.fill = band_fill

            c = ws_top.cell(row=base + 9, column=col, value=f'= "Attributes: " & {DREF(13, data_row)}')
            c.alignment = left_wrap; c.border = box; c.fill = band_fill

            c = ws_top.cell(row=base + 10, column=col, value=f'= "Notes: " & {DREF(14, data_row)}')
            c.alignment = left_wrap; c.border = box; c.fill = band_fill

        ws_top.freeze_panes = None  # no freeze on Top 20

        bio = BytesIO()
        wb.save(bio)
        bio.seek(0)
        return bio.read()

    export_placeholder.download_button(
        "Download Spreadsheet (.xlsx)",
        data=build_xlsx_two_sheets(st.session_state.results),
        file_name=f"{(meta.get('name') or 'Top20').replace(' ','_')}_Top20.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
