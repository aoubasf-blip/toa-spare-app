from pathlib import Path
import re
import html
import base64
import pandas as pd
import streamlit as st

# ============================================================
# CONFIG
# ============================================================
BASE_DIR = Path(__file__).parent
DATA_FILE = BASE_DIR / "master_data_DATA.xlsx"
RAW_DATA_FILE = BASE_DIR / "master_data.xlsx"
IMAGE_DIR = BASE_DIR / "images"
SPARE_IMAGE_DIR = IMAGE_DIR / "spare"
PRODUCT_IMAGE_DIR = IMAGE_DIR / "product"
ASSET_DIR = BASE_DIR / "assets"
LOGO_FILE = ASSET_DIR / "after_sale_logo.jpg"

APP_TITLE = "TOA | JOMOO After Sale Service"

# ============================================================
# PAGE SETUP
# ============================================================
st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🧰",
    layout="wide",
)

# ============================================================
# STYLE
# ============================================================
st.markdown(
    """
    <style>
    :root {
        --toa-ink: #071427;
        --toa-navy: #0B1F3A;
        --toa-blue: #123C7C;
        --toa-red: #D71920;
        --toa-silver: #D8E1EA;
        --toa-line: #E5EAF1;
        --toa-bg: #F6F8FB;
        --toa-card: #FFFFFF;
        --toa-text: #1F2937;
        --toa-muted: #687589;
    }

    header[data-testid="stHeader"] {display: none;}
    [data-testid="stToolbar"] {display: none;}
    [data-testid="stDecoration"] {display: none;}

    .block-container {
        padding-top: 1.25rem;
        padding-bottom: 2.5rem;
        max-width: 1520px;
    }

    .stApp {
        background:
            radial-gradient(circle at top left, rgba(18,60,124,0.08), transparent 28%),
            linear-gradient(180deg, #FBFCFE 0%, #F3F6FA 42%, #EEF3F8 100%);
        font-family: "Segoe UI", system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
        color: var(--toa-text);
    }

    /* ===== Brand header: clean white premium ===== */
    .app-hero {
        position: relative;
        overflow: hidden;
        border-radius: 30px;
        padding: 1.05rem 1.6rem;
        margin-bottom: 1.35rem;
        background: #FFFFFF;
        border: 1px solid rgba(225,232,240,0.95);
        box-shadow: 0 22px 60px rgba(15, 35, 65, 0.08);
        isolation: isolate;
    }
    .app-hero:before {
        content: "";
        position: absolute;
        right: -140px;
        top: -150px;
        width: 460px;
        height: 460px;
        border-radius: 999px;
        background: radial-gradient(circle, rgba(18,60,124,0.075), rgba(18,60,124,0.025) 48%, transparent 72%);
        z-index: -1;
    }
    .app-hero:after {
        content: "";
        position: absolute;
        left: 30px;
        bottom: 0;
        width: 38%;
        height: 5px;
        border-radius: 999px;
        background: linear-gradient(90deg, var(--toa-red), var(--toa-blue), transparent);
        opacity: 0.95;
    }
    .hero-grid {
        display: grid;
        grid-template-columns: 255px minmax(0, 1fr);
        gap: 1.35rem;
        align-items: center;
    }
    .logo-stage {
        height: 235px;
        background: #FFFFFF;
        border: 0;
        box-shadow: none;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 0;
    }
    .brand-logo {
        width: 230px;
        height: 230px;
        object-fit: contain;
        border-radius: 999px;
        background: #FFFFFF;
        box-shadow: none;
        border: 0;
        filter: none;
        mix-blend-mode: normal;
        display: block;
    }
    .logo-fallback {
        color: #64748b;
        text-align: center;
        font-weight: 800;
        line-height: 1.6;
        font-size: 0.85rem;
    }
    .hero-content {padding: 0.2rem 0.25rem;}
    .eyebrow {
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
        padding: 0.34rem 0.78rem;
        border-radius: 999px;
        background: #F3F6FA;
        border: 1px solid #E3EAF2;
        color: var(--toa-blue);
        font-weight: 900;
        font-size: 0.76rem;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        margin-bottom: 0.75rem;
    }
    .eyebrow-dot {
        width: 8px;
        height: 8px;
        border-radius: 999px;
        background: var(--toa-red);
        box-shadow: 0 0 0 4px rgba(215,25,32,0.12);
    }
    .hero-title {
        font-size: clamp(2rem, 3.8vw, 3.6rem);
        line-height: 1.02;
        font-weight: 950;
        letter-spacing: -0.052em;
        color: var(--toa-navy);
        margin: 0 0 0.6rem 0;
    }
    .hero-title span {
        display: block;
        color: #4B5F78;
        font-size: 0.63em;
        letter-spacing: -0.03em;
        font-weight: 850;
    }
    .hero-sub {
        max-width: 980px;
        color: #4B5563;
        font-size: 1.01rem;
        line-height: 1.68;
        font-weight: 560;
        margin-bottom: 0.95rem;
    }
    .hero-actions {
        display: flex;
        flex-wrap: wrap;
        gap: 0.5rem;
        margin-bottom: 0.2rem;
    }
    .hero-chip {
        padding: 0.44rem 0.76rem;
        border-radius: 999px;
        background: #FFFFFF;
        border: 1px solid #DDE6F0;
        color: #334155;
        font-size: 0.82rem;
        font-weight: 850;
        box-shadow: 0 6px 16px rgba(15,23,42,0.045);
    }
    .hero-metrics {display: none;}

    /* ===== Search ===== */
    .search-card-title {
        padding: 1rem 1.12rem;
        border-radius: 24px;
        background: #FFFFFF;
        box-shadow: 0 14px 34px rgba(15,23,42,0.07);
        border: 1px solid #E2E8F0;
        color: var(--toa-navy);
        margin-bottom: 0.85rem;
    }
    .search-card-title .search-title {
        font-size: 1.14rem;
        font-weight: 920;
        margin-bottom: 0.18rem;
    }
    .search-card-title .search-caption {
        color: #64748B;
        font-size: 0.86rem;
        line-height: 1.45;
    }

    /* ===== Detail card: less boxy, more professional ===== */
    .detail-card {
        padding: 1.05rem 1.2rem;
        margin-bottom: 1rem;
        background: rgba(255,255,255,0.97);
        border: 1px solid #E3EAF2;
        border-radius: 28px;
        box-shadow: 0 20px 52px rgba(15,23,42,0.075);
    }
    .code-title {
        font-size: 1.75rem;
        line-height: 1.08;
        font-weight: 950;
        color: var(--toa-navy);
        letter-spacing: -0.035em;
        margin-bottom: 0.18rem;
    }
    .subheading {
        font-size: 0.96rem;
        color: var(--toa-muted);
        margin-bottom: 0.48rem;
        font-weight: 650;
    }
    .toa-highlight {
        padding: 0.7rem 0.85rem;
        border-radius: 18px;
        background: linear-gradient(90deg, #FFF7F7 0%, #F6FAFF 100%);
        border: 1px solid #E6EDF5;
        border-left: 5px solid var(--toa-red);
        margin: 0.45rem 0 0.62rem 0;
        color: var(--toa-navy);
    }
    .toa-highlight-label {
        font-size: 0.72rem;
        font-weight: 950;
        color: var(--toa-red);
        letter-spacing: 0.13em;
        text-transform: uppercase;
        margin-bottom: 0.24rem;
    }
    .toa-highlight-name {font-size: 1.06rem; font-weight: 900; line-height: 1.45;}

    .info-grid {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 0.28rem 1.05rem;
        margin-top: 0.42rem;
    }
    .info-grid.two-col {grid-template-columns: repeat(2, minmax(0, 1fr));}
    .field-box {
        background: transparent;
        border: 0;
        border-bottom: 1px solid #EDF1F6;
        border-radius: 0;
        padding: 0.28rem 0 0.34rem 0;
        min-height: 0;
        box-shadow: none;
    }
    .field-box small {
        display: block;
        font-size: 0.66rem;
        color: #7A8798;
        font-weight: 900;
        letter-spacing: 0.045em;
        margin-bottom: 0.04rem;
        text-transform: uppercase;
    }
    .field-box div {
        font-weight: 740;
        color: #1F2937;
        line-height: 1.28;
        overflow-wrap: anywhere;
        font-size: 0.92rem;
    }

    .visual-stack {display: flex; flex-direction: column; gap: 0.65rem;}
    .visual-panel {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 18px;
        padding: 0.55rem;
        box-shadow: 0 10px 24px rgba(15,23,42,0.045);
    }
    .visual-title {
        display: flex; align-items: center; justify-content: space-between;
        font-size: 0.72rem;
        letter-spacing: 0.09em;
        text-transform: uppercase;
        font-weight: 950;
        color: var(--toa-blue);
        margin-bottom: 0.35rem;
    }
    .visual-title.spare-title {color: var(--toa-red);}
    .visual-title span {
        height: 1px; flex: 1; margin-left: 0.5rem;
        background: linear-gradient(90deg, currentColor, transparent); opacity: 0.22;
    }
    .no-image {
        width: 100%; min-height: 118px; border-radius: 14px;
        background: linear-gradient(135deg, #F8FAFC, #EEF3F8);
        border: 1px dashed #C9D3DF;
        display: flex; align-items: center; justify-content: center; color: #64748b; font-weight: 800; text-align: center; padding: 0.8rem;
        font-size: 0.86rem;
    }

    .section-title {
        margin-top: 0.72rem;
        padding-top: 0.68rem;
        border-top: 1px solid #E7EDF4;
        color: var(--toa-navy);
        font-weight: 950;
        font-size: 0.96rem;
    }
    .summary-note {
        padding: 0.8rem 0.95rem;
        border-radius: 18px;
        background: rgba(255,255,255,0.96);
        border: 1px solid #E2E8F0;
        box-shadow: 0 10px 24px rgba(15,23,42,0.045);
        margin-bottom: 0.9rem;
        color: #334155;
        font-weight: 850;
    }
    .stDataFrame {border-radius: 16px; overflow: hidden;}
    div[data-testid="stMetric"] {
        background: rgba(255,255,255,0.88);
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 0.7rem 0.9rem;
        box-shadow: 0 8px 20px rgba(15,23,42,0.04);
    }
    hr {border-color: #E7EDF4 !important;}
    @media (max-width: 980px) {
        .hero-grid {grid-template-columns: 1fr;}
        .logo-stage {height: 190px; justify-content: center;}
        .brand-logo {height: 185px; width: 185px;}
        .info-grid {grid-template-columns: 1fr;}
    }


    /* ============================================================
       MOBILE OPTIMIZATION ONLY
       ใช้เฉพาะหน้าจอโทรศัพท์ ไม่กระทบ Desktop / Notebook
    ============================================================ */
    @media (max-width: 768px) {

        .block-container {
            padding-left: 0.85rem !important;
            padding-right: 0.85rem !important;
            padding-top: 0.75rem !important;
            padding-bottom: 1.5rem !important;
            max-width: 100% !important;
        }

        .stApp {
            background: linear-gradient(180deg, #FBFCFE 0%, #F3F6FA 100%);
        }

        /* Header บนมือถือให้สั้นลง ไม่กินพื้นที่ */
        .app-hero {
            border-radius: 18px;
            padding: 0.9rem 0.95rem;
            margin-bottom: 0.9rem;
            box-shadow: 0 8px 22px rgba(15, 35, 65, 0.07);
        }

        .app-hero:before {
            width: 220px;
            height: 220px;
            right: -90px;
            top: -110px;
        }

        .app-hero:after {
            left: 18px;
            width: 58%;
            height: 4px;
        }

        .hero-grid {
            display: block;
        }

        .logo-stage {
            height: auto;
            margin-bottom: 0.55rem;
            justify-content: flex-start;
        }

        .brand-logo {
            width: 86px;
            height: 86px;
            margin: 0;
        }

        .hero-content {
            padding: 0;
        }

        .eyebrow {
            font-size: 0.58rem;
            letter-spacing: 0.08em;
            padding: 0.25rem 0.55rem;
            margin-bottom: 0.5rem;
        }

        .eyebrow-dot {
            width: 6px;
            height: 6px;
            box-shadow: 0 0 0 3px rgba(215,25,32,0.12);
        }

        .hero-title {
            font-size: 1.45rem;
            line-height: 1.12;
            letter-spacing: -0.025em;
            margin-bottom: 0.45rem;
            text-align: left;
        }

        .hero-title span {
            font-size: 0.62em;
            margin-top: 0.18rem;
        }

        .hero-sub {
            font-size: 0.78rem;
            line-height: 1.48;
            margin-bottom: 0.65rem;
        }

        .hero-actions {
            gap: 0.35rem;
        }

        .hero-chip {
            font-size: 0.66rem;
            padding: 0.32rem 0.5rem;
        }

        /* Search card ให้กระชับ */
        .search-card-title {
            border-radius: 16px;
            padding: 0.75rem 0.85rem;
            margin-bottom: 0.65rem;
            box-shadow: 0 8px 20px rgba(15,23,42,0.05);
        }

        .search-card-title .search-title {
            font-size: 0.98rem;
        }

        .search-card-title .search-caption {
            font-size: 0.73rem;
            line-height: 1.35;
        }

        /* บังคับ Streamlit columns ให้เรียงลงบนมือถือ */
        div[data-testid="column"] {
            width: 100% !important;
            flex: 1 1 100% !important;
            min-width: 100% !important;
        }

        div[data-testid="stHorizontalBlock"] {
            gap: 0.65rem !important;
        }

        /* Detail card บนมือถือ */
        .detail-card {
            border-radius: 18px;
            padding: 0.9rem 0.85rem;
            margin-bottom: 0.85rem;
            box-shadow: 0 8px 24px rgba(15,23,42,0.06);
        }

        .code-title {
            font-size: 1.25rem;
            line-height: 1.2;
            letter-spacing: -0.02em;
            padding-bottom: 0.35rem;
            border-bottom: 1px solid #E7EDF4;
            margin-bottom: 0.45rem;
        }

        .subheading {
            font-size: 0.78rem;
            line-height: 1.4;
            margin-bottom: 0.55rem;
        }

        .toa-highlight {
            padding: 0.62rem 0.7rem;
            border-radius: 14px;
            margin: 0.45rem 0 0.6rem 0;
        }

        .toa-highlight-label {
            font-size: 0.58rem;
            letter-spacing: 0.08em;
        }

        .toa-highlight-name {
            font-size: 0.82rem;
            line-height: 1.42;
        }

        .section-title {
            font-size: 0.86rem;
            margin-top: 0.65rem;
            padding-top: 0.62rem;
        }

        /* จาก 3 คอลัมน์ ให้เป็น 1 คอลัมน์บนมือถือ */
        .info-grid,
        .info-grid.two-col {
            grid-template-columns: 1fr !important;
            gap: 0.15rem;
            margin-top: 0.25rem;
        }

        .field-box {
            padding: 0.34rem 0;
            border-bottom: 1px solid #EEF2F6;
        }

        .field-box small {
            font-size: 0.6rem;
            letter-spacing: 0.035em;
            color: #7A8798;
        }

        .field-box div {
            font-size: 0.82rem;
            line-height: 1.35;
            font-weight: 720;
        }

        /* รูปภาพบนมือถือให้เล็กลง ไม่ดันข้อมูลยาวเกิน */
        .visual-stack {
            gap: 0.45rem;
            margin-bottom: 0.55rem;
        }

        .visual-panel {
            border-radius: 14px;
            padding: 0.45rem;
            box-shadow: 0 6px 16px rgba(15,23,42,0.04);
        }

        .visual-title {
            font-size: 0.58rem;
            letter-spacing: 0.07em;
            margin-bottom: 0.25rem;
        }

        .no-image {
            min-height: 72px;
            font-size: 0.72rem;
            border-radius: 12px;
        }

        /* ตาราง Summary บนมือถือให้เลื่อนได้ ไม่กินจอ */
        .summary-note {
            border-radius: 14px;
            padding: 0.65rem 0.75rem;
            font-size: 0.82rem;
            margin-bottom: 0.65rem;
        }

        .stDataFrame {
            max-height: 220px;
            overflow: auto;
        }

        /* ลดช่องไฟ Streamlit widget */
        div[data-testid="stVerticalBlock"] {
            gap: 0.55rem;
        }

        div[data-testid="stMetric"] {
            padding: 0.55rem 0.65rem;
            border-radius: 13px;
        }

        /* input / dropdown / radio ให้พอดีมือถือ */
        .stTextInput input {
            font-size: 0.9rem;
        }

        .stSelectbox,
        .stRadio,
        .stTextInput,
        .stCheckbox {
            font-size: 0.85rem;
        }

        h3 {
            font-size: 1.05rem !important;
            margin-top: 0.7rem !important;
        }
    }

    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# HELPERS
# ============================================================
def clean_value(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat", "null"}:
        return ""
    return text


def safe_html(value) -> str:
    return html.escape(clean_value(value))


def safe_filename(value: str) -> str:
    """ใช้ logic เดียวกับ extract_images.py เพื่อให้ชื่อไฟล์รูปตรงกัน
    เช่น 11030-2-1/31K-TH11 -> 11030-2-1_31K-TH11
    """
    text = clean_value(value)
    if not text:
        return ""
    text = text.replace("/", "_").replace("\\", "_")
    text = re.sub(r"[^0-9A-Za-zก-๙_\-\.]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_. ")
    return text[:150]


def legacy_safe_filename(value: str) -> str:
    """ชื่อไฟล์จาก app เวอร์ชันเก่า เผื่อมีรูปที่ถูกสร้างไว้ด้วย pattern เดิม"""
    text = clean_value(value)
    if not text:
        return ""
    return "".join(ch if ch.isalnum() else "_" for ch in text)


def first_existing(row: pd.Series, candidates: list[str]) -> str:
    for col in candidates:
        if col in row.index:
            val = clean_value(row.get(col, ""))
            if val:
                return val
    return ""


def find_col(df: pd.DataFrame, keywords: list[str]) -> str | None:
    for col in df.columns:
        low = str(col).lower().strip()
        if all(k.lower() in low for k in keywords):
            return col
    return None


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [clean_value(c) for c in df.columns]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]

    rename = {}
    for col in df.columns:
        c = col.strip()
        low = c.lower()

        if c in ["Spare part code", "Spare part code ", "Spare Part code", "Spare part Code", "Spare part code New"]:
            rename[col] = "Spare Part Code" if "new" not in low else "Spare Part Code New"
        elif c in ["Description", "Description (EN)"]:
            rename[col] = "Description (EN)"
        elif c in ["Description（Thai）", "Description(Thai)", "Description （Thai）", "Description (TH)"]:
            rename[col] = "Description (TH)"
        elif c in ["Description（Chinese）", "Description(Chinese)", "Description （Chinese）", "Description (CN)"]:
            rename[col] = "Description (CN)"
        elif c in ["Material description", "Material Description"]:
            rename[col] = "TOA Name EN"
        elif c in ["Material description TH", "Material description Thai", "Material Description TH", "ชื่อไทย", "รายละเอียดภาษาไทย"]:
            rename[col] = "TOA Name TH"
        elif c in ["EAN/UPC", "Barcode", "barcode", "EAN", "UPC"]:
            rename[col] = "Barcode"
        elif c in ["Material", "Material No", "Material No.", "Material Code"]:
            rename[col] = "Material No."
        elif c in ["Product Model", "Model"]:
            rename[col] = "Model"
        elif c in ["Product name", "Product Name"]:
            rename[col] = "Product Name"
        elif c in ["Waranty", "Warranty"]:
            rename[col] = "Warranty Type"
        elif c in ["Unit Price\n(CNY)", "Unit Price (CNY)"]:
            rename[col] = "Unit Price (CNY)"
        elif c in ["Spare parts quantity", "Spare Parts Qty"]:
            rename[col] = "Spare Parts Qty"
        elif "net" in low and "weight" in low:
            rename[col] = "Net Weight"
        elif "gross" in low and "weight" in low:
            rename[col] = "Gross Weight"

    if rename:
        df = df.rename(columns=rename)
        df = df.loc[:, ~pd.Index(df.columns).duplicated()]

    return df


@st.cache_data(show_spinner="กำลังโหลดข้อมูลอะไหล่...")
def load_data() -> pd.DataFrame:
    source = DATA_FILE if DATA_FILE.exists() else RAW_DATA_FILE
    if not source.exists():
        raise FileNotFoundError("ไม่พบ master_data_DATA.xlsx หรือ master_data.xlsx")

    xls = pd.ExcelFile(source, engine="openpyxl")

    if "ALL_COMBINED" in xls.sheet_names:
        df = pd.read_excel(source, sheet_name="ALL_COMBINED", dtype=str, engine="openpyxl")
        df = normalize_columns(df)
    else:
        frames = []
        for sheet in xls.sheet_names:
            raw = pd.read_excel(source, sheet_name=sheet, dtype=str, engine="openpyxl")
            if raw.empty:
                continue
            raw = raw.dropna(how="all")
            raw = normalize_columns(raw)
            if "Category" not in raw.columns:
                raw.insert(0, "Category", sheet)
            frames.append(raw)
        if not frames:
            raise ValueError("ไม่พบข้อมูลใน Excel")
        df = pd.concat(frames, ignore_index=True)
        df = normalize_columns(df)

    df = df.fillna("").astype(str)

    for col in ["Category", "Model", "Product Name"]:
        if col not in df.columns:
            df[col] = ""

    # สร้างคอลัมน์ค้นหาจากหลายข้อมูล โดยไม่ใช้ภาษาจีนเป็นหลัก
    code_cols = [c for c in [
        "Spare Part Code", "Spare Part Code New", "JOMOO Spare Part Code", "Material No.", "Barcode", "EAN/UPC",
        "Material No. (Master)", "Barcode (EAN/UPC) (Master)", "Barcode (EAN/UPC)", "Spare Part Code Alt 1"
    ] if c in df.columns]
    name_cols = [c for c in [
        "TOA Spare Part Name (TH Display)", "TOA Spare Part Name (TH)", "TOA Spare Part Name (TH Generated)", "TOA Spare Part Name (EN)",
        "TOA Name TH", "TOA Name EN", "Description (TH)", "Description (EN)", "Product Name", "Model"
    ] if c in df.columns]

    if "All Spare Part Codes" not in df.columns:
        df["All Spare Part Codes"] = df[code_cols].agg(" | ".join, axis=1) if code_cols else ""
    if "All Names / Descriptions" not in df.columns:
        df["All Names / Descriptions"] = df[name_cols].agg(" | ".join, axis=1) if name_cols else ""

    return df


def filter_any(df: pd.DataFrame, cols: list[str], keyword: str) -> pd.Series:
    if not keyword:
        return pd.Series(False, index=df.index)
    mask = pd.Series(False, index=df.index)
    for col in cols:
        if col in df.columns:
            mask = mask | df[col].astype(str).str.contains(keyword, case=False, na=False, regex=False)
    return mask


def build_model_options(df: pd.DataFrame, keyword: str = "", category: str = "ทั้งหมด") -> list[tuple[str, str]]:
    mdf = df.copy()
    if category != "ทั้งหมด" and "Category" in mdf.columns:
        mdf = mdf[mdf["Category"].astype(str) == category]

    if keyword:
        mask = filter_any(mdf, ["Model", "Product Name", "All Names / Descriptions"], keyword)
        mdf = mdf[mask]

    if "Model" not in mdf.columns:
        return []

    cols = [c for c in ["Category", "Model", "Product Name"] if c in mdf.columns]
    out = []
    for _, r in mdf[cols].drop_duplicates().iterrows():
        model = clean_value(r.get("Model", ""))
        if not model:
            continue
        cat = clean_value(r.get("Category", ""))
        pname = clean_value(r.get("Product Name", ""))
        parts = []
        if cat:
            parts.append(cat)
        parts.append(model)
        if pname:
            parts.append(pname)
        out.append((" | ".join(parts), model))
    return sorted(out, key=lambda x: x[0].lower())


def image_candidates_for_code(code: str) -> list[Path]:
    code = clean_value(code)
    if not code:
        return []
    names = list(dict.fromkeys([
        code,
        code.replace("/", "_"),
        safe_filename(code),
    ]))
    paths = []
    for name in names:
        for ext in [".png", ".jpg", ".jpeg", ".webp"]:
            paths.append(SPARE_IMAGE_DIR / f"{name}{ext}")
            paths.append(IMAGE_DIR / f"{name}{ext}")
    return paths


def image_candidates_for_product(model: str, pname: str) -> list[Path]:
    # รองรับชื่อไฟล์หลายแบบ เพราะรูปจาก Excel อาจถูกเซฟจากชื่อ Model ที่มี / แล้วแปลงเป็น _
    # เช่น Model: 11030-2-1/31K-TH11 -> file: 11030-2-1_31K-TH11.png
    raw_keys = [model, pname]
    keys = []
    for x in raw_keys:
        x = clean_value(x)
        if not x:
            continue
        keys.extend([
            x,
            x.replace("/", "_").replace("\\", "_"),
            safe_filename(x),
            legacy_safe_filename(x),
        ])
    keys = [k for k in dict.fromkeys([clean_value(x) for x in keys]) if k]
    paths = []
    for key in keys:
        for ext in [".png", ".jpg", ".jpeg", ".webp"]:
            paths.append(PRODUCT_IMAGE_DIR / f"{key}{ext}")
    return paths


def first_image(paths: list[Path]) -> Path | None:
    for p in paths:
        if p.exists():
            return p
    return None


def render_field_grid(items: list[tuple[str, str]], two_col: bool = False):
    cls = "info-grid two-col" if two_col else "info-grid"
    html_items = [f"<div class='field-box'><small>{html.escape(label)}</small><div>{safe_html(value) or '-'}</div></div>" for label, value in items]
    st.markdown(f"<div class='{cls}'>" + "".join(html_items) + "</div>", unsafe_allow_html=True)


def get_toa_official_name_parts(row: pd.Series) -> tuple[str, str]:
    th = first_existing(row, [
        "TOA Spare Part Name (TH Display)", "TOA Spare Part Name (TH)", "TOA Spare Part Name (TH Generated)",
        "TOA Name TH", "Material description TH", "Description (TH)", "Description TH"
    ])
    en = first_existing(row, [
        "TOA Spare Part Name (EN)", "TOA Name EN", "Material description",
        "Material Description", "Description (EN) Alt", "Description (EN)", "Description EN"
    ])
    return th, en


def get_toa_official_name(row: pd.Series) -> str:
    th, en = get_toa_official_name_parts(row)
    if th and en and th != en:
        return f"{th} / {en}"
    return th or en or ""


def render_visual_panel(title: str, img_path: Path | None, spare: bool = False):
    title_cls = "visual-title spare-title" if spare else "visual-title"
    st.markdown('<div class="visual-panel">', unsafe_allow_html=True)
    st.markdown(f'<div class="{title_cls}">{html.escape(title)}<span></span></div>', unsafe_allow_html=True)
    if img_path:
        st.image(str(img_path), use_container_width=True)
    else:
        st.markdown('<div class="no-image">No image available</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)


def render_card(row: pd.Series):
    jomoo_code = first_existing(row, ["Spare Part Code", "JOMOO Spare Part Code", "Spare Part Code New", "Spare Part Code Alt 1"])
    code_new = first_existing(row, ["Spare Part Code New"])
    code_alt = first_existing(row, ["Spare Part Code Alt 1", "Spare Part Code Alt 2"])
    material_no = first_existing(row, ["Material No. (Master)", "Material No.", "Material", "Material Code"])
    barcode = first_existing(row, ["Barcode (EAN/UPC) (Master)", "Barcode (EAN/UPC)", "Barcode", "EAN/UPC", "EAN"])
    material_sheet = first_existing(row, ["Material No.", "Material", "Material Code"])
    barcode_sheet = first_existing(row, ["Barcode (EAN/UPC)", "Barcode", "EAN/UPC", "EAN"])
    model = first_existing(row, ["Model", "Product Model"])
    secondary_model = first_existing(row, ["Product Model (Secondary Code)"])
    pname = first_existing(row, ["Product Name", "Product name"])
    category = first_existing(row, ["Category"])
    toa_th, toa_en = get_toa_official_name_parts(row)
    toa_th_source = first_existing(row, ["TOA Spare Part Name (TH Source)"])
    toa_name = get_toa_official_name(row)

    spare_img = None
    for code in [jomoo_code, code_new, code_alt, material_no, material_sheet]:
        spare_img = first_image(image_candidates_for_code(code))
        if spare_img:
            break
    product_img = first_image(image_candidates_for_product(model, pname))

    st.markdown('<div class="detail-card">', unsafe_allow_html=True)
    col_img, col_info = st.columns([0.92, 2.7], gap="large")

    with col_img:
        st.markdown('<div class="visual-stack">', unsafe_allow_html=True)
        render_visual_panel("Product Image", product_img, spare=False)
        render_visual_panel("Spare Part Image", spare_img, spare=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with col_info:
        st.markdown(f'<div class="code-title">{safe_html(jomoo_code) or "No Spare Part Code"}</div>', unsafe_allow_html=True)
        subtitle = " · ".join([x for x in [model, secondary_model, pname] if x])
        if subtitle:
            st.markdown(f'<div class="subheading">{safe_html(subtitle)}</div>', unsafe_allow_html=True)

        # Official section must show only verified / official names.
        # If Thai name is generated from English, keep it out of Official Name
        # and show it later in the Description section as a suggested Thai name.
        is_generated_th = toa_th_source.lower().startswith("generated")
        official_th = "" if is_generated_th else toa_th

        if official_th or toa_en:
            label = "TOA OFFICIAL SPARE PART NAME"
            name_lines = []
            if official_th:
                name_lines.append(f"<div><strong>TH:</strong> {safe_html(official_th)}</div>")
            if toa_en:
                name_lines.append(f"<div><strong>EN:</strong> {safe_html(toa_en)}</div>")
            st.markdown(
                f"""
                <div class="toa-highlight">
                    <div class="toa-highlight-label">{label}</div>
                    <div class="toa-highlight-name">{''.join(name_lines)}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown('<div class="section-title">Basic Info</div>', unsafe_allow_html=True)
        render_field_grid([
            ("Category", category),
            ("Model", model),
            ("Secondary Model", secondary_model),
            ("Product Name", pname),
            ("JOMOO Spare Part Code", jomoo_code),
            ("Spare Code New", code_new),
            ("TOA Material No.", material_no),
            ("TOA Barcode", barcode),
            ("Unit Price (CNY)", first_existing(row, ["Unit Price (CNY)"])),
            ("Spare Parts Qty", first_existing(row, ["Spare Parts Qty", "Spare parts quantity"])),
            ("Net Weight", first_existing(row, ["Net Weight (G)", "Net Weight"])),
            ("Gross Weight", first_existing(row, ["Gross Weight (G)", "Gross Weight"])),
        ])

        st.markdown('<div class="section-title">Description / รายละเอียด</div>', unsafe_allow_html=True)
        desc_rows = [
            ("TOA Official Name TH", official_th),
            ("TOA Official Name EN", toa_en),
        ]
        if is_generated_th and toa_th:
            desc_rows.extend([
                ("Suggested Thai Name", toa_th),
                ("Thai Name Status", "Generated from English / รอตรวจยืนยันก่อนใช้เป็น Official"),
            ])
        else:
            desc_rows.append(("Thai Name Status", toa_th_source))
        desc_rows.extend([
            ("JOMOO Description TH", first_existing(row, ["Description (TH)"])),
            ("JOMOO Description EN", first_existing(row, ["Description (EN)"])),
            ("All Names / Descriptions", first_existing(row, ["All Names / Descriptions"])),
            ("Remark", first_existing(row, ["Remark", "Remarks"])),
        ])
        render_field_grid(desc_rows, two_col=True)

        st.markdown('<div class="section-title">Code / Barcode Mapping</div>', unsafe_allow_html=True)
        render_field_grid([
            ("All Spare Codes", first_existing(row, ["All Spare Part Codes"])),
            ("JOMOO Code / Old Code", jomoo_code),
            ("New / THAD Code", code_new or code_alt),
            ("Master Material No.", first_existing(row, ["Material No. (Master)"])),
            ("Master Barcode", first_existing(row, ["Barcode (EAN/UPC) (Master)"])),
            ("Sheet Material No.", material_sheet),
            ("Sheet Barcode", barcode_sheet),
            ("Source", " / ".join([x for x in [first_existing(row, ["Source Sheet"]), first_existing(row, ["Source Row"])] if x])),
        ])

    st.markdown('</div>', unsafe_allow_html=True)


def get_logo_data_uri() -> str:
    if not LOGO_FILE.exists():
        return ""
    try:
        ext = LOGO_FILE.suffix.lower().replace(".", "") or "jpg"
        mime = "jpeg" if ext in {"jpg", "jpeg"} else ext
        encoded = base64.b64encode(LOGO_FILE.read_bytes()).decode("utf-8")
        return f"data:image/{mime};base64,{encoded}"
    except Exception:
        return ""

def render_brand_header():
    logo_uri = get_logo_data_uri()
    if logo_uri:
        logo_html = f'<img class="brand-logo" src="{logo_uri}" alt="TOA JOMOO After Sale Service Logo">'
    else:
        logo_html = '<div class="logo-fallback">วางโลโก้ที่<br><code>assets/after_sale_logo.jpg</code></div>'

    st.markdown(
        f"""
        <section class="app-hero">
            <div class="hero-grid">
                <div class="logo-stage">
                    {logo_html}
                </div>
                <div class="hero-content">
                    <div class="eyebrow"><span class="eyebrow-dot"></span> TOA | JOMOO AFTER SALE SERVICE</div>
                    <div class="hero-title">Spare Part Platform<span>Clean master search for service team</span></div>
                    <div class="hero-sub">
                        ระบบค้นหาอะไหล่จาก Master Data สำหรับทีม After Sale Service — รองรับ JOMOO Spare Part Code,
                        TOA Material No., Barcode, Model และชื่ออะไหล่ทางการภาษาไทย / อังกฤษ พร้อมรูป Product และ Spare Part
                    </div>
                    <div class="hero-actions">
                        <span class="hero-chip">Spare Part Code</span>
                        <span class="hero-chip">TOA Official Name</span>
                        <span class="hero-chip">Barcode / Material</span>
                        <span class="hero-chip">Product & Spare Images</span>
                    </div>
                </div>
            </div>
        </section>
        """,
        unsafe_allow_html=True,
    )

# ============================================================
# MAIN APP
# ============================================================
def main():
    render_brand_header()

    try:
        df = load_data()
    except PermissionError:
        st.error("เปิดไฟล์ Excel ไม่ได้ กรุณาปิดไฟล์ master_data_DATA.xlsx / master_data.xlsx ก่อน แล้ว Refresh ใหม่")
        st.stop()
    except Exception as e:
        st.error(f"โหลดข้อมูลไม่ได้: {e}")
        st.stop()

    search_col, result_col = st.columns([0.92, 2.25], gap="large")

    with search_col:
        st.markdown(
            '<div class="search-card-title"><div class="search-title">🔍 Search Spare Part</div>'
            '<div class="search-caption">เลือกวิธีค้นหา แล้วระบบจะแสดงรายละเอียดพร้อมรูปอะไหล่และชื่อทางการฝั่ง TOA</div></div>',
            unsafe_allow_html=True,
        )

        cat_list = ["ทั้งหมด"]
        if "Category" in df.columns:
            cats = sorted([x for x in df["Category"].astype(str).str.strip().unique() if x and x.lower() != "nan"])
            cat_list += cats
        category = st.selectbox("Category / Sheet", cat_list)

        search_mode = st.radio(
            "Search mode",
            [
                "Code / Barcode / Material",
                "ชื่ออะไหล่ / รายละเอียด",
                "Product / Model",
                "ค้นหาทุกคอลัมน์",
            ],
        )

        exact_match = False
        keyword = ""
        model_selected = ""

        if search_mode == "Product / Model":
            product_mode = st.radio("Product search", ["เลือกจาก Model dropdown", "พิมพ์คำค้น"], label_visibility="collapsed")
            if product_mode == "เลือกจาก Model dropdown":
                filter_kw = st.text_input("ตัวกรอง Model", placeholder="เช่น X70, TS3, 11252").strip()
                options = build_model_options(df, keyword=filter_kw, category=category)
                labels = ["— เลือก Model —"] + [x[0] for x in options]
                label_selected = st.selectbox("Model", labels)
                if label_selected != "— เลือก Model —":
                    model_selected = dict(options).get(label_selected, "")
            else:
                keyword = st.text_input("Product / Model", placeholder="เช่น X70, 11252, Smart Toilet").strip()
        elif search_mode == "Code / Barcode / Material":
            keyword = st.text_input("Code / Barcode / Material", placeholder="เช่น K1125208-1, T2I..., 885...").strip()
            exact_match = st.checkbox("ค้นหาแบบตรงตัว", value=False)
        elif search_mode == "ชื่ออะไหล่ / รายละเอียด":
            keyword = st.text_input("ชื่ออะไหล่ / รายละเอียด", placeholder="เช่น ฝาถังพักน้ำ, Tank cover, ปุ่มกด").strip()
        else:
            keyword = st.text_input("ค้นหาทุกคอลัมน์", placeholder="ใส่คำค้นใดก็ได้").strip()


    result_df = None
    status = "พิมพ์คำค้นหรือเลือก Model เพื่อเริ่มค้นหา"
    status_type = "info"

    working = df.copy()
    if category != "ทั้งหมด" and "Category" in working.columns:
        working = working[working["Category"].astype(str) == category]

    if model_selected:
        result_df = working[working["Model"].astype(str).str.strip() == model_selected].copy()
        status = f"พบ {len(result_df):,} รายการสำหรับ Model: {model_selected}"
        status_type = "success" if not result_df.empty else "warning"
    elif keyword:
        if search_mode == "Code / Barcode / Material":
            cols = [c for c in ["Spare Part Code", "Spare Part Code New", "Spare Part Code Alt 1", "JOMOO Spare Part Code", "Material No.", "Material No. (Master)", "Barcode", "Barcode (EAN/UPC)", "Barcode (EAN/UPC) (Master)", "EAN/UPC", "All Spare Part Codes"] if c in working.columns]
            if exact_match:
                mask = pd.Series(False, index=working.index)
                for col in cols:
                    mask = mask | (working[col].astype(str).str.lower().str.strip() == keyword.lower().strip())
            else:
                mask = filter_any(working, cols, keyword)
        elif search_mode == "ชื่ออะไหล่ / รายละเอียด":
            cols = [c for c in ["TOA Spare Part Name (TH Display)", "TOA Spare Part Name (TH)", "TOA Spare Part Name (TH Generated)", "TOA Spare Part Name (EN)", "TOA Name TH", "TOA Name EN", "Description (TH)", "Description (EN)", "All Names / Descriptions"] if c in working.columns]
            mask = filter_any(working, cols, keyword)
        elif search_mode == "Product / Model":
            cols = [c for c in ["Model", "Product Name", "All Names / Descriptions"] if c in working.columns]
            mask = filter_any(working, cols, keyword)
        else:
            mask = working.apply(lambda row: row.astype(str).str.contains(keyword, case=False, na=False, regex=False).any(), axis=1)

        result_df = working[mask].copy()
        status = f"พบ {len(result_df):,} รายการสำหรับคำค้น: {keyword}" if not result_df.empty else f"ไม่พบข้อมูลสำหรับคำค้น: {keyword}"
        status_type = "success" if not result_df.empty else "warning"

    with result_col:
        if status_type == "success":
            st.success(status)
        elif status_type == "warning":
            st.warning(status)
        else:
            st.info(status)

        with st.expander("Data coverage / ตรวจสอบข้อมูล", expanded=False):
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total rows", f"{len(df):,}")
            c2.metric("Filtered category", f"{len(working):,}")
            c3.metric("Product images", f"{len(list(PRODUCT_IMAGE_DIR.glob('*'))) if PRODUCT_IMAGE_DIR.exists() else 0:,}")
            c4.metric("Spare images", f"{len(list(SPARE_IMAGE_DIR.glob('*'))) if SPARE_IMAGE_DIR.exists() else 0:,}")
            st.caption(f"Data file: {DATA_FILE.name if DATA_FILE.exists() else RAW_DATA_FILE.name}")

        if result_df is not None and not result_df.empty:
            sort_cols = [c for c in ["Category", "Model", "Spare Part Code"] if c in result_df.columns]
            if sort_cols:
                result_df = result_df.sort_values(sort_cols)

            summary_cols = [c for c in [
                "Category", "Model", "Product Name", "Spare Part Code", "Spare Part Code New",
                "TOA Spare Part Name (TH Display)", "TOA Spare Part Name (TH)", "TOA Spare Part Name (TH Generated)", "TOA Spare Part Name (EN)", "TOA Name TH", "TOA Name EN",
                "Material No. (Master)", "Barcode (EAN/UPC) (Master)", "Material No.", "Barcode (EAN/UPC)", "Unit Price (CNY)"
            ] if c in result_df.columns]

            st.markdown('<div class="summary-note">ภาพรวมรายการที่ค้นพบ</div>', unsafe_allow_html=True)
            if summary_cols:
                st.dataframe(result_df[summary_cols].reset_index(drop=True), use_container_width=True, hide_index=True, height=260)

            st.markdown("### Detail View")
            max_cards = st.slider("จำนวน Card ที่แสดง", min_value=1, max_value=min(100, len(result_df)), value=min(20, len(result_df)))
            for _, row in result_df.head(max_cards).iterrows():
                render_card(row)


if __name__ == "__main__":
    main()

