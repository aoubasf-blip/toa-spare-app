import streamlit as st
import pandas as pd
from pathlib import Path

# =========================
# Page Config
# =========================
st.set_page_config(
    page_title="TOA JOMOO",
    page_icon="🧰",
    layout="wide"
)

# =========================
# Custom CSS
# =========================
st.markdown("""
<style>
    .main {
        padding-top: 1rem;
    }

    .hero-box {
        background: linear-gradient(135deg, #ffffff 0%, #f4f7fb 100%);
        border: 1px solid #e6eaf0;
        border-radius: 22px;
        padding: 24px;
        margin-bottom: 20px;
        box-shadow: 0 4px 18px rgba(0,0,0,0.05);
    }

    .hero-title {
        font-size: 34px;
        font-weight: 800;
        color: #0f2440;
        margin-bottom: 4px;
        line-height: 1.15;
    }

    .hero-subtitle {
        font-size: 16px;
        color: #53657d;
        margin-bottom: 12px;
    }

    .result-count {
        font-size: 14px;
        color: #52616f;
        margin: 8px 0 16px 0;
    }

    .part-card {
        background-color: #ffffff;
        border: 1px solid #dfe5ec;
        border-radius: 18px;
        padding: 16px;
        margin-bottom: 14px;
        box-shadow: 0 3px 12px rgba(0,0,0,0.05);
    }

    .part-name-th {
        font-size: 18px;
        font-weight: 800;
        color: #0f2440;
        line-height: 1.35;
        margin-bottom: 8px;
    }

    .part-name-en {
        font-size: 14px;
        color: #65758b;
        line-height: 1.35;
        margin-bottom: 14px;
    }

    .info-row {
        border-top: 1px solid #edf0f4;
        padding-top: 10px;
        margin-top: 10px;
    }

    .info-label {
        font-size: 12px;
        color: #7c8a9b;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.03em;
        margin-bottom: 3px;
    }

    .info-value {
        font-size: 15px;
        color: #14213d;
        font-weight: 600;
        word-break: break-word;
    }

    .barcode-value {
        font-size: 17px;
        color: #0b5cad;
        font-weight: 800;
        word-break: break-word;
    }

    .small-note {
        font-size: 12px;
        color: #7b8794;
    }

    @media (max-width: 768px) {
        .block-container {
            padding-left: 1rem !important;
            padding-right: 1rem !important;
            padding-top: 1rem !important;
        }

        .hero-box {
            padding: 18px;
            border-radius: 18px;
        }

        .hero-title {
            font-size: 24px;
        }

        .hero-subtitle {
            font-size: 14px;
        }

        .part-card {
            padding: 14px;
            border-radius: 16px;
            margin-bottom: 12px;
        }

        .part-name-th {
            font-size: 17px;
        }

        .part-name-en {
            font-size: 13px;
        }

        .info-value {
            font-size: 14px;
        }

        .barcode-value {
            font-size: 16px;
        }
    }
</style>
""", unsafe_allow_html=True)

# =========================
# Load Data
# =========================
DATA_FILE = Path("master_data_DATA.xlsx")

@st.cache_data
def load_data():
    if not DATA_FILE.exists():
        return pd.DataFrame()

    df = pd.read_excel(DATA_FILE)

    # Clean column names
    df.columns = [str(c).strip() for c in df.columns]

    # Fix duplicated Material description columns if found
    cols = list(df.columns)

    material_desc_cols = [c for c in cols if c.lower().startswith("material description")]

    rename_map = {}

    if len(material_desc_cols) >= 1:
        rename_map[material_desc_cols[0]] = "Official Name EN"

    if len(material_desc_cols) >= 2:
        rename_map[material_desc_cols[1]] = "Generated Name TH"

    # Common column mapping
    for c in cols:
        low = c.lower().strip()

        if low == "material":
            rename_map[c] = "Spare Part Code"

        if low in ["ean/upc", "ean", "upc", "barcode", "barcode no.", "barcode no"]:
            rename_map[c] = "Barcode"

        if "model" == low or "model no" in low or "model no." in low:
            rename_map[c] = "Model"

    df = df.rename(columns=rename_map)

    # Make sure key columns exist
    required_cols = [
        "Spare Part Code",
        "Official Name EN",
        "Generated Name TH",
        "Barcode",
        "Model"
    ]

    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    # Fill blank
    df = df.fillna("")

    # Remove fully blank rows
    df = df[df.astype(str).apply(lambda x: "".join(x), axis=1).str.strip() != ""]

    return df


df = load_data()

# =========================
# Header
# =========================
st.markdown("""
<div class="hero-box">
    <div class="hero-title">Spare Part Platform</div>
    <div class="hero-subtitle">TOA | JOMOO After Sale Service</div>
    <div class="small-note">
        ค้นหา Spare Part Code, Barcode, Official Name, ชื่อภาษาไทย หรือ Model ได้ในช่องเดียว
    </div>
</div>
""", unsafe_allow_html=True)

if df.empty:
    st.error("ไม่พบไฟล์ master_data_DATA.xlsx หรือไฟล์ยังอ่านไม่ได้ กรุณาตรวจสอบชื่อไฟล์ใน GitHub")
    st.stop()

# =========================
# Search
# =========================
search_text = st.text_input(
    "ค้นหา Spare Part",
    placeholder="พิมพ์รหัสอะไหล่ / Barcode / ชื่อสินค้า / Model"
)

# Searchable text
search_cols = df.columns.tolist()

if search_text.strip():
    keyword = search_text.strip().lower()

    mask = df[search_cols].astype(str).apply(
        lambda row: row.str.lower().str.contains(keyword, na=False).any(),
        axis=1
    )

    filtered_df = df[mask].copy()
else:
    filtered_df = df.copy()

# =========================
# Filter / Limit
# =========================
col1, col2 = st.columns([1, 1])

with col1:
    view_mode = st.radio(
        "รูปแบบการแสดงผล",
        ["Card View", "Table View"],
        horizontal=True
    )

with col2:
    max_items = st.selectbox(
        "จำนวนที่แสดง",
        [20, 50, 100, 200, "ทั้งหมด"],
        index=1
    )

if max_items != "ทั้งหมด":
    display_df = filtered_df.head(int(max_items))
else:
    display_df = filtered_df

st.markdown(
    f'<div class="result-count">พบข้อมูลทั้งหมด <b>{len(filtered_df)}</b> รายการ / แสดง <b>{len(display_df)}</b> รายการ</div>',
    unsafe_allow_html=True
)

# =========================
# Helper
# =========================
def safe_value(row, col):
    value = row.get(col, "")
    if pd.isna(value):
        return ""
    value = str(value).strip()
    if value.lower() in ["nan", "none", "nat"]:
        return ""
    return value


# =========================
# Card View
# =========================
if view_mode == "Card View":

    for _, row in display_df.iterrows():
        code = safe_value(row, "Spare Part Code")
        name_en = safe_value(row, "Official Name EN")
        name_th = safe_value(row, "Generated Name TH")
        barcode = safe_value(row, "Barcode")
        model = safe_value(row, "Model")

        # If Thai name is blank, use English name as card title
        title = name_th if name_th else name_en

        st.markdown(f"""
        <div class="part-card">
            <div class="part-name-th">{title}</div>
            <div class="part-name-en">{name_en}</div>

            <div class="info-row">
                <div class="info-label">Spare Part Code</div>
                <div class="info-value">{code}</div>
            </div>

            <div class="info-row">
                <div class="info-label">Barcode / EAN</div>
                <div class="barcode-value">{barcode}</div>
            </div>

            <div class="info-row">
                <div class="info-label">Model</div>
                <div class="info-value">{model}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

# =========================
# Table View
# =========================
else:
    main_cols = [
        "Spare Part Code",
        "Generated Name TH",
        "Official Name EN",
        "Barcode",
        "Model"
    ]

    available_cols = [c for c in main_cols if c in display_df.columns]

    st.dataframe(
        display_df[available_cols],
        use_container_width=True,
        hide_index=True
    )
