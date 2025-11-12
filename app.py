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
EXCEL_IMG_MAX_PX = 200
EXCEL_IMG_ROW_HEIGHT = 80
RANK_FILL = PatternFill("solid", fgColor="C45500")
WHITE_FONT = Font(color="FFFFFF", bold=True)
SAVED_DIR = ".saved_searches"; os.makedirs(SAVED_DIR, exist_ok=True)

def _session():
    s = requests.Session(); s.headers.update({"Accept-Language": "en-US,en;q=0.9"}); return s
def get_soup(url, s=None): s=s or _session(); r=s.get(url, headers={"User-Agent":random.choice(USER_AGENTS)},timeout=15); r.raise_for_status(); return BeautifulSoup(r.text,"html.parser")

def utc_now(): return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

# -------- AMAZON SCRAPING ----------
def extract_asin(url):
    for r in [r"/dp/([A-Z0-9]{10})", r"/gp/product/([A-Z0-9]{10})"]:
        m=re.search(r,url); 
        if m:return m.group(1)
def parse_top20(url,s=None):
    s=s or _session(); soup=get_soup(url,s)
    items=[]; seen=set()
    for a in soup.find_all("a",href=True):
        href=a["href"]; asin=extract_asin(href)
        if not asin or asin in seen:continue
        title=a.get_text(strip=True) or (a.find("img",alt=True) or {}).get("alt","")
        if not title:title=a.get("title","")
        seen.add(asin)
        items.append({"ASIN":asin,"Title":title,"URL":urljoin(url,href)})
        if len(items)>=20:break
    return items

def extract_price(soup):
    for sp in soup.select("span.a-offscreen"):
        t=sp.get_text(strip=True)
        if t.startswith("$"):return t
    return ""
def extract_sellthrough(soup):
    for div in soup.find_all("div",string=re.compile(r"bought",re.I)):
        txt=div.get_text(" ",strip=True)
        if re.search(r"\b\d",txt):return txt
    for span in soup.find_all("span",string=re.compile(r"bought",re.I)):
        return span.get_text(" ",strip=True)
    return ""
def extract_image(soup):
    og=soup.find("meta",{"property":"og:image"})
    if og:return og.get("content","")
    img=soup.find("img",id="landingImage")
    if img:return img.get("data-old-hires") or img.get("src","")
    return ""
def fetch_details(url,s=None):
    s=s or _session()
    try:
        soup=get_soup(url,s)
        return extract_price(soup),extract_image(soup),extract_sellthrough(soup)
    except Exception:return"","",""

# ---------- SAVE / LOAD ----------
def _p(id):return os.path.join(SAVED_DIR,id)
def _d(id):return os.path.join(_p(id),"data.csv")
def _m(id):return os.path.join(_p(id),"meta.json")
def list_saved():
    runs=[]
    for rid in os.listdir(SAVED_DIR):
        try:
            mp=_m(rid); dp=_d(rid)
            if os.path.exists(mp)and os.path.exists(dp):
                meta=json.load(open(mp))
                runs.append({"id":rid,**meta})
        except:pass
    return sorted(runs,key=lambda r:r.get("updated",""),reverse=True)
def save_run(id,df,meta):
    os.makedirs(_p(id),exist_ok=True)
    df.to_csv(_d(id),index=False)
    meta["updated"]=utc_now()
    json.dump(meta,open(_m(id),"w"),indent=2)
def load_run(id):
    return pd.read_csv(_d(id)),json.load(open(_m(id)))
def meta_new(name,url):return{"name":name,"url":url,"created":utc_now(),"updated":utc_now()}

# -------- STREAMLIT UI ------------
if "df" not in st.session_state: st.session_state.df=None
if "meta" not in st.session_state: st.session_state.meta=None

st.sidebar.header("Saved Searches")
runs=list_saved()
for r in runs:
    st.sidebar.markdown(f"**{r['name']}**  \n<small>{r.get('updated','')}</small>",unsafe_allow_html=True)
    c1,c2=st.sidebar.columns(2)
    if c1.button("Load",key=f"load{r['id']}"):
        df,m=load_run(r["id"]); st.session_state.df=df; st.session_state.meta=m; st.success("Loaded.")
    if c2.button("Delete",key=f"del{r['id']}"):
        import shutil; shutil.rmtree(_p(r["id"]),ignore_errors=True); st.rerun()
st.sidebar.divider()
st.sidebar.caption("Searches are saved locally.")

with st.sidebar.expander("⚙️ Settings",expanded=False):
    delay_min=st.slider("Min delay",0.5,3.0,1.0,0.1)
    delay_max=st.slider("Max delay",0.6,4.0,2.0,0.1)

st.title("Amazon Top-20 → Excel Exporter")

url=st.text_input("Amazon Best Sellers URL")
desc=st.text_input("Category Description",placeholder="e.g. Single Board Computers")
mode=st.radio("View Mode:",["List","Grid","Compact"],horizontal=True)

c1,c2=st.columns([1,3])
if c1.button("Fetch Top 20",type="primary"):
    s=_session()
    try:items=parse_top20(url,s)
    except Exception as e:st.error(f"Failed: {e}"); items=[]
    out=[]
    for i,it in enumerate(items,1):
        price,img,sell=fetch_details(it["URL"],s)
        out.append({"Rank":i,"Title":it["Title"],"ASIN":it["ASIN"],"URL":it["URL"],
                    "Price":price,"Image":img,"Sell":sell,"MCSKU":"","MCTitle":"",
                    "MCRetail":"","MCCost":"","Avg1_4":"","Attributes":"","Notes":""})
        time.sleep(random.uniform(delay_min,delay_max))
    df=pd.DataFrame(out); st.session_state.df=df
    st.session_state.meta=meta_new(desc or "Top 20",url)
    st.success("Top 20 fetched!")

if st.session_state.df is not None:
    df=st.session_state.df.copy()
    # ====== RENDER MODES ======
    if mode=="List":
        for _,r in df.iterrows():
            st.markdown(f"### #{r['Rank']}  {r['Title']}")
            cols=st.columns([1,3])
            with cols[0]:
                if r["Image"]:
                    st.image(r["Image"],width=120)
            with cols[1]:
                st.write(r["Price"])
                if r["Sell"]: st.caption(r["Sell"])
                st.write(f"[Open on Amazon]({r['URL']})")
            st.divider()
    else:
        ncols=5; rows=[df.iloc[i:i+ncols] for i in range(0,len(df),ncols)]
        card_css="""
        <style>
        .card{border:1px solid #ddd;padding:6px;border-radius:8px;margin:3px;box-shadow:0 0 4px rgba(0,0,0,0.1);}
        .card:hover{box-shadow:0 0 8px rgba(0,0,0,0.2);}
        .title{font-weight:600;font-size:0.9rem}
        .price{color:#B12704;font-weight:600;}
        .sell{font-size:0.8rem;color:#555;}
        </style>"""
        st.markdown(card_css,unsafe_allow_html=True)
        imgw=120 if mode=="Grid" else 80
        for row in rows:
            cols=st.columns(ncols)
            for c,(i,r) in zip(cols,row.iterrows()):
                html=f"<div class='card'><div style='background:#C45500;color:white;padding:2px 6px;border-radius:4px;display:inline-block;'>#{r['Rank']}</div><br>"
                if r["Image"]: html+=f"<img src='{r['Image']}' width='{imgw}'/><br>"
                html+=f"<div class='title'>{r['Title']}</div>"
                html+=f"<div class='price'>{r['Price']}</div>"
                if r["Sell"]: html+=f"<div class='sell'>{r['Sell']}</div>"
                html+=f"<a href='{r['URL']}'>Open on Amazon</a></div>"
                st.markdown(html,unsafe_allow_html=True)
            st.write("")

    # ====== EXPORT ======
    def build_xlsx(df):
        wb=Workbook(); ws_top=wb.active; ws_top.title="Top 20"; ws_data=wb.create_sheet("Data")
        cols=["Rank","ASIN","URL","Image","Title","Price","Sell","MCSKU","MCTitle","MCRetail","MCCost","Avg1_4","Attributes","Notes"]
        for j,c in enumerate(cols,1): ws_data.cell(1,j,c)
        for i,r in enumerate(df.itertuples(),2):
            ws_data.append([r.Rank,r.ASIN,r.URL,r.Image,r.Title,r.Price,r.Sell,r.MCSKU,r.MCTitle,r.MCRetail,r.MCCost,r.Avg1_4,r.Attributes,r.Notes])
        for j in range(1,len(cols)+1): ws_data.column_dimensions[get_column_letter(j)].width=22
        thin=Side(style="thin",color="DDDDDD"); box=Border(left=thin,right=thin,top=thin,bottom=thin)
        wrap=Alignment(wrap_text=True,vertical="top")
        ITEMS_PER_ROW=5; ROW_BLOCK=10
        for c in range(1,ITEMS_PER_ROW+1): ws_top.column_dimensions[get_column_letter(c)].width=EXCEL_COL_WIDTH
        for idx in range(len(df)):
            g=idx//ITEMS_PER_ROW; col=(idx%ITEMS_PER_ROW)+1; base=g*ROW_BLOCK+1; dr=idx+2
            ws_top.cell(base,col).value=f"Rank #{df.at[idx,'Rank']}"; ws_top.cell(base,col).fill=RANK_FILL; ws_top.cell(base,col).font=WHITE_FONT
            img_url=df.at[idx,"Image"]
            if isinstance(img_url,str) and img_url.startswith("http"):
                try:
                    im=requests.get(img_url,timeout=10).content
                    pil=PILImage.open(BytesIO(im)); pil.thumbnail((EXCEL_IMG_MAX_PX,EXCEL_IMG_MAX_PX))
                    fn=BytesIO(); pil.save(fn,"PNG"); fn.seek(0)
                    img=XLImage(fn); img.width=EXCEL_IMG_MAX_PX; img.height=EXCEL_IMG_MAX_PX
                    ws_top.add_image(img,f"{get_column_letter(col)}{base+1}")
                except:pass
            lines=[
                ("Amazon",f'=Data!E{dr}'),
                ("Price",f'=Data!F{dr}'),
                ("Sell",f'=Data!G{dr}'),
                ("MC SKU",f'=Data!H{dr}'),
                ("MC Title",f'=Data!I{dr}'),
                ("MC Retail",f'=Data!J{dr}'),
                ("MC Cost",f'=Data!K{dr}'),
                ("1-4 Avg",f'=Data!L{dr}'),
                ("Attributes",f'=Data!M{dr}'),
                ("Notes",f'=Data!N{dr}')
            ]
            for r_off,(lab,ref) in enumerate(lines,3):
                c=ws_top.cell(base+r_off,col); c.value=f"{lab}: "; c.alignment=wrap; c.border=box
                c2=ws_top.cell(base+r_off,col); c.value=f"=CONCAT(\"{lab}: \",{ref})"; c.alignment=wrap; c.border=box
        bio=BytesIO(); wb.save(bio); bio.seek(0); return bio.read()

    st.download_button("📥 Download Excel",data=build_xlsx(df),
        file_name=f"{(st.session_state.meta['name']).replace(' ','_')}_Top20.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
