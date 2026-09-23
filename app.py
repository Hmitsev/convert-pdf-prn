import streamlit as st
import re
import pandas as pd
from PyPDF2 import PdfReader
import io
import base64
import pdfplumber
import os
import io
import re
import csv
import zipfile
import unicodedata
import pandas as pd
import pdfplumber

from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side
)
# ======================================================
# ✅ BACKGROUND FUNCTIONS
# ======================================================

def set_bg(image_file):

    try:

        with open(image_file, "rb") as f:
            encoded = base64.b64encode(
                f.read()
            ).decode()

        st.markdown(
            f"""
            <style>
            .stApp {{
                background-image: url("data:image/png;base64,{encoded}");
                background-size: cover;
                background-position: center;
                background-repeat: no-repeat;
            }}
            </style>
            """,
            unsafe_allow_html=True
        )

    except:
        pass


def set_login_bg(image_file):

    try:

        with open(image_file, "rb") as f:
            encoded = base64.b64encode(
                f.read()
            ).decode()

        st.markdown(
            f"""
            <style>
            .stApp {{
                background-image: url("data:image/png;base64,{encoded}");
                background-size: 70% auto;
                background-position: center top;
                background-repeat: no-repeat;
                background-color: #050505;
            }}
            </style>
            """,
            unsafe_allow_html=True
        )

    except:
        pass
# ======================================================
# ✅ LOGOUT BUTTON (FIXED TOP RIGHT)
# ======================================================

logout_col1, logout_col2, logout_col3 = st.columns([8,1,1])

with logout_col3:
    if st.button("🚪", help="Logout"):
        st.session_state["logged_in"] = False
        st.rerun()

st.markdown("""
<style>
div[data-testid="column"]:nth-of-type(3) {
    position: fixed;
    top: 4px;
    right: 45px;
    z-index: 9999;
}
</style>
""", unsafe_allow_html=True)


# ======================================================
# ✅ LOGIN SYSTEM (FINAL INLINE LOGO WORKING)
# ======================================================
def check_login():

    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if not st.session_state["logged_in"]:

        # ✅ background
        set_bg("background_login.png")

        # ✅ HEADER В 1 РЕД (чрез columns, но правилно оразмерени)
        col1, col2 = st.columns([4,1])

        with col1:
            st.markdown("""
            <div style="
                text-align:right;
                font-size:32px;
                font-weight:900;
                color:white;
                white-space:nowrap;
            ">
            </div>
            """, unsafe_allow_html=True)

        with col2:
            st.image("Screenshot 2026-06-18 093459.png", width=60)

        # ✅ леко spacing
        st.markdown("<br>", unsafe_allow_html=True)

        # ✅ Login title
        st.markdown(
            "<h1 style='text-align:center; color:white;'>🔐  Customs Inter Cars App</h1>",
            unsafe_allow_html=True
        )

        # ✅ Inputs
        username = st.text_input("User name:")
        password = st.text_input("password", type="password")

        # ✅ Button
        if st.button("Log In"):
            if username == "mitnica" and password == "Intercars2026":
                st.session_state["logged_in"] = True
                st.rerun()
            else:
                st.error("❌ Грешно име или парола")

        return False

    return True


if not check_login():
    st.stop()

# ✅ main background
set_bg("background.png")

# ======================================================
# ✅ FINAL UI (RESET + PERFECT ORDER)
# ======================================================

st.markdown("""
<style>
.source-title {
    font-size: 23px;
    font-weight: 800;
    color: #ff8c00;
    text-shadow: 0 0 8px rgba(255,140,0,0.5);

/* ✅ Add file малък */
.add-file {
    display:inline-block;
    background: rgba(255,255,255,0.05);
    border-radius: 8px;
    padding: 4px 10px;
    color: white;
    font-size: 16px;
    margin-bottom: 6px;
}

/* ✅ маха текста от default button */
button[data-testid="baseButton-secondary"] p {
    opacity: 0;
}
</style>
""", unsafe_allow_html=True)


# ✅ SIDEBAR
menu = st.sidebar.selectbox("Suppliers", ["CASTROL & NESTE", "MOTUL", "FLUKAR", "GASOLINE", "VALVOLINE", "ORLEN", "Chempioil (FANFARO)", "FUCHS", "FEBI", "ELROMI RONAX","NISTA", "AMTRA" , "AUTO MEGA" ,"EMINIA" ,"Brehman","PRN за Navision" ,"PDF към Excel" ])
# ✅ статичен списък (като таблица в sidebar)
st.sidebar.markdown("### 📋 Suppliers & File type")

suppliers_table = [
    ("FLUKAR", "Excel"),
    ("ELROMI RONAX", "Excel"),
    ("VALVOLINE", "Excel"),
    ("ORLEN", "Excel"),
    ("Chempioil (FANFARO)", "Excel"),
    ("AMTRA", "Excel"),
    ("FUCHS", "Excel or PDF"),
    ("CASTROL", "Excel"),
    ("MOTUL", "PDF"),
    ("NESTE", "Excel"),
    ("Gasoline", "PDF"),
    ("FEBI", "PDF"),
    ("NISTA", "Excel+NewSheet"),
    ("AUTO MEGA", "Excel"),
    ("EMINIA", "Excel"),
    ("Brehman", "Excel"),
    ("PRN за Navision"),
]

# ✅ обръщаме реда (както искаш)
suppliers_table = suppliers_table[::-1]

# ✅ Suppliers & File Type (Custom Table)

st.sidebar.markdown("""
<style>
.sup-table{
    width:100%;
    border-collapse:collapse;
    font-size:13px;
}

.sup-table th{
    background:#d71919;
    color:white;
    padding:6px;
    text-align:center;
    border:1px solid rgba(255,255,255,0.25);
}

.sup-table td{
    padding:5px;
    border:1px solid rgba(255,255,255,0.15);
    background:rgba(255,255,255,0.04);
    color:white;
}

.sup-table tr:hover td{
    background:rgba(255,255,255,0.10);
}
</style>
""", unsafe_allow_html=True)

suppliers_html = """
<table class="sup-table">
<tr>
    <th>Supplier</th>
    <th>File</th>
</tr>
<tr><td>FLUKAR</td><td>Excel</td></tr>
<tr><td>ELROMI RONAX</td><td>Excel</td></tr>
<tr><td>VALVOLINE</td><td>Excel</td></tr>
<tr><td>ORLEN</td><td>Excel</td></tr>
<tr><td>Chempioil (FANFARO)</td><td>Excel or PDF</td></tr>
<tr><td>AMTRA</td><td>Excel+NewSheet</td></tr>
<tr><td>FUCHS</td><td>Excel or PDF</td></tr>
<tr><td>CASTROL</td><td>Excel</td></tr>
<tr><td>MOTUL</td><td>PDF</td></tr>
<tr><td>NESTE</td><td>Excel</td></tr>
<tr><td>NISTA</td><td>Excel+NewSheet</td></tr>
<tr><td>AUTO MEGA</td><td>Excel</td></tr>
<tr><td>EMINIA</td><td>Excel</td></tr>
<tr><td>Brehman</td><td>Excel</td></tr>
<tr><td>Gasoline</td><td>PDF</td></tr>
<tr><td>FEBI</td><td>Excel+NewFile.Xls</td></tr>
<tr><td>PRN за Navision</td><td>Excel (invoice/item/qty/price) </td></tr>
<tr><td>PDF към Excel</td><td>PDF → XLSX</td></tr>
</table>
"""

st.sidebar.markdown(
    suppliers_html,
    unsafe_allow_html=True
)

# ✅ RESET при смяна на supplier
if "prev_supplier" not in st.session_state:
    st.session_state["prev_supplier"] = menu

if st.session_state["prev_supplier"] != menu:
    st.session_state["source_type"] = ""
    st.session_state["prev_supplier"] = menu

# ✅ заглавие
st.markdown('<div class="source-title">👇 Choose Source</div>', unsafe_allow_html=True)


# ✅ STATE
if "source_type" not in st.session_state:
    st.session_state["source_type"] = ""


# ✅ бутони
col1, col2 = st.columns(2)

with col1:
    if st.button("PDF", use_container_width=True):
        st.session_state["source_type"] = "PDF"
        st.rerun()

with col2:
    if st.button("Excel", use_container_width=True):
        st.session_state["source_type"] = "Excel"
        st.rerun()


source_type = st.session_state["source_type"]


# ✅ цветове + overlay
if source_type == "PDF":
    pdf_color = "#ff3b3b"
    excel_color = "#444"
    pdf_overlay = "<span style='color:#ff3b3b;'>You chose: PDF</span>"
    excel_overlay = "<span style='color:white;'>Excel</span>"

elif source_type == "Excel":
    pdf_color = "#444"
    excel_color = "#36c165"
    pdf_overlay = "<span style='color:white;'>PDF</span>"
    excel_overlay = "<span style='color:#36c165;'>You chose: Excel</span>"

else:
    pdf_color = "#444"
    excel_color = "#444"
    pdf_overlay = "<span style='color:white;'>PDF</span>"
    excel_overlay = "<span style='color:white;'>Excel</span>"


# ✅ стил на бутоните
st.markdown(f"""
<style>
div[data-testid="column"]:nth-of-type(1) button {{
    background-color: {pdf_color};
    height: 60px;
    border-radius: 12px;
}}

div[data-testid="column"]:nth-of-type(2) button {{
    background-color: {excel_color};
    height: 60px;
    border-radius: 12px;
}}
</style>
""", unsafe_allow_html=True)


# ✅ overlay текст
col1, col2 = st.columns(2)

with col1:
    st.markdown(f"""
    <div style="
        margin-top:-65px;
        display:flex;
        justify-content:flex-end;
        padding-right:20px;
        pointer-events:none;
    ">
        {pdf_overlay}
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div style="
        margin-top:-65px;
        display:flex;
        justify-content:flex-end;
        padding-right:20px;
        pointer-events:none;
    ">
        {excel_overlay}
    </div>
    """, unsafe_allow_html=True)


# ✅ ✅ ПРАВИЛЕН РЕД (FIX)

# ======================================================
# ✅ FILE UPLOADER
# ======================================================

if menu == "PDF към Excel":

    st.session_state["source_type"] = "PDF"
    source_type = "PDF"

    uploaded_files = st.file_uploader(
        "Качи един или няколко PDF файла",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_to_excel_uploader"
    )

else:

    uploaded_files = st.file_uploader(
        "",
        type=["pdf"] if source_type == "PDF" else ["xlsx", "xls"],
        accept_multiple_files=True,
        key="supplier_file_uploader"
    )
# ======================================================
# ✅ ULTRA GLASS SIDEBAR (PRO VERSION)
# ======================================================

st.markdown("""
<style>

/* ✅ Sidebar container */
section[data-testid="stSidebar"] {
    background: transparent !important;
}

/* ✅ GLASS EFFECT */
section[data-testid="stSidebar"] > div {
    background: rgba(0,0,0,0.01) !important;  /* почти прозрачно */

    backdrop-filter: blur(18px) saturate(140%);
    -webkit-backdrop-filter: blur(18px) saturate(140%);

    border-right: 4px solid rgba(255,255,255,0.7);  /* силен метален борд */

    /* ✅ вътрешен glow */
    box-shadow:
        inset 0 0 10px rgba(255,255,255,0.05),
        0 0 20px rgba(255,255,255,0.1);
}


/* ✅ текст */
section[data-testid="stSidebar"] * {
    color: white !important;
}


/* ✅ SELECT BOX */
div[data-baseweb="select"] {
    background: rgba(255,255,255,0.04) !important;
    backdrop-filter: blur(8px);
    border-radius: 10px;
    border: 1px solid rgba(255,255,255,0.25);

    cursor: pointer !important;
}


/* ✅ hover ефект (много фин) */
div[data-baseweb="select"]:hover {
    background: rgba(255,255,255,0.08) !important;
    border: 1px solid rgba(255,255,255,0.4);
    box-shadow: 0 0 8px rgba(255,255,255,0.2);
}


/* ✅ pointer fix */
div[data-baseweb="select"] * {
    cursor: inherit !important;
}

</style>
""", unsafe_allow_html=True)

# ======================================================
# ✅ КОДОВЕ
# ======================================================
ALLOWED_CODES = [
    "27101991","27101981","27101983","27101987",
    "27101993","27101999","34031910","34039900",
    "34031980","38119000","38112100","38249992",
    "27101225","38140090","38249996","00000000"
]


# ======================================================
# ✅ CASTROL
# ======================================================
def parse_castrol(text):

    rows = []
    lines = text.split("\n")
    current_liters = 0

    for line in lines:

        multi = re.search(r"(\d+)X(\d+)L", line)
        single = re.search(r"(\d+)L", line)

        if multi:
            current_liters = int(multi.group(1)) * int(multi.group(2))
        elif single:
            current_liters = int(single.group(1))

        if "Cod Vamal" in line:
            try:
                code = re.search(r"Cod Vamal:(\d+)", line).group(1)
                qty = int(re.search(r"ST\s*(\d+)", line).group(1))

                rows.append({
                    "Тарифен код": code,
                    "Количество": qty,
                    "wid": current_liters,
                    "kolichestvo": qty * current_liters,
                    "тегло": 0
                })
            except:
                pass

    return pd.DataFrame(rows)
    # ======================================================
# ✅ CASTROL EXCEL
# ======================================================
def parse_castrol_excel(file):

    df = pd.read_excel(file)

    rows = []

    for _, row in df.iterrows():

        try:

            code = str(
                row["Commodity code"]
            )

            code = re.sub(
                r"\D",
                "",
                code
            )[:8]

            if code not in ALLOWED_CODES:
                continue

            qty = pd.to_numeric(
                row["Delivery quantity"],
                errors="coerce"
            )

            volume = pd.to_numeric(
                row["Volume"],
                errors="coerce"
            )

            net_weight = pd.to_numeric(
                row["Net Weight"],
                errors="coerce"
            )

            wid = pd.to_numeric(
                row["Type of packaging"],
                errors="coerce"
            )

            if pd.isna(qty):
                continue

            if pd.isna(volume):
                continue

            if pd.isna(net_weight):
                continue

            if pd.isna(wid):
                continue

            rows.append({
                "Тарифен код": code,
                "Количество": qty,
                "wid": wid,
                "kolichestvo": volume,
                "тегло": net_weight
            })

        except:
            continue

    if not rows:
        st.error("❌ CASTROL Excel parser не извлече данни")
        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = df_out.groupby(
        ["Тарифен код", "wid"],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df_out
    
# ======================================================
# ✅ NISTA EXCEL - ALL PACKAGING FORMATS
# ======================================================
def parse_nista_excel(file):

    df = pd.read_excel(
        file,
        header=None
    )

    rows = []

    for i in range(len(df)):

        row = df.iloc[i]

        try:

            row_text = " ".join(
                str(x)
                for x in row
                if pd.notna(x)
            )

            normalized_text = (
                row_text
                .lower()
                .replace("×", "x")
                .replace("х", "x")
                .replace(",", ".")
            )

            # ==========================================
            # ✅ ТАРИФЕН КОД
            # Пример: 2710 19 81 -> 27101981
            # ==========================================

            code_match = re.search(
                r"27[\d\s]{6,}",
                normalized_text
            )

            if not code_match:
                continue

            digits = re.sub(
                r"\D",
                "",
                code_match.group(0)
            )

            if len(digits) < 8:
                continue

            code = digits[:8]

            if code not in ALLOWED_CODES:
                continue

            # ==========================================
            # ✅ ОБЩО КОЛИЧЕСТВО
            # Пример: 5643 liter -> 5643
            # ==========================================

            menge_match = re.search(
                r"(\d+(?:\.\d+)?)\s*liter\b",
                normalized_text
            )

            if not menge_match:
                continue

            menge = float(
                menge_match.group(1)
            )

            # ==========================================
            # ✅ WID
            #
            # 12x1l   -> 1
            # 3x5l    -> 5
            # 4x4l    -> 4
            # 1x20l   -> 20
            # 1x55l   -> 55
            # 1x209l  -> 209
            # 60l     -> 60
            # 200l    -> 200
            # ==========================================

            wid = None

            package_match = re.search(
                r"\d+(?:\.\d+)?\s*x\s*"
                r"(\d+(?:\.\d+)?)\s*l\b",
                normalized_text
            )

            if package_match:

                wid = float(
                    package_match.group(1)
                )

            else:

                single_matches = re.findall(
                    r"(\d+(?:\.\d+)?)\s*l\b",
                    normalized_text
                )

                if single_matches:
                    wid = float(
                        single_matches[-1]
                    )

            if wid is None or wid <= 0:
                continue

            # ==========================================
            # ✅ ТЕГЛО
            # Взима последната числова клетка от реда
            # ==========================================

            weight = None

            for cell in reversed(row):

                try:

                    cell_text = (
                        str(cell)
                        .strip()
                        .replace(" ", "")
                        .replace(",", ".")
                    )

                    val = float(cell_text)

                    if val > 0:
                        weight = val
                        break

                except:
                    continue

            if weight is None:
                continue

            # ==========================================
            # ✅ BROJ
            # Общо литри / разфасовка
            # ==========================================

            broj = menge / wid

            rows.append({
                "Тарифен код": code,
                "Количество": broj,
                "wid": wid,
                "kolichestvo": menge,
                "тегло": weight
            })

        except:
            continue

    if not rows:

        st.error(
            "❌ NISTA parser не извлече данни"
        )

        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = df_out.groupby(
        [
            "Тарифен код",
            "wid"
        ],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df_out

# ======================================================
# ✅ MOTUL (FINAL REAL WORKING + FILTER ✅)
# ======================================================
def parse_motul(text):

    rows = []
    lines = text.split("\n")

    current_qty = 0
    current_weight = 0
    liters_per_unit = 0
    units_in_box = 1

    for line in lines:

        # ✅ КОЛИЧЕСТВО
        match = re.search(r"\d+\s+\d+\s+(\d+)\s+[\d,\.]+\s+[\d,\.]+", line)
        if match:
            try:
                current_qty = int(match.group(1))
            except:
                pass

        # ✅ РАЗФАСОВКА
        multi = re.findall(r"(\d+)X([\d\.,]+)(?:L|kg)", line, re.IGNORECASE)
        single = re.search(r"([\d\.,]+)(?:L|kg)", line, re.IGNORECASE)

        if multi:
            units_in_box = int(multi[-1][0])
            liters_per_unit = float(multi[-1][1].replace(",", "."))
        elif single:
            units_in_box = 1
            liters_per_unit = float(single.group(1).replace(",", "."))

        # ✅ ТЕГЛО
        weight_match = re.search(
            r"\d+\s+\d+\s+(\d+)\s+([\d\s,]+)\s+([\d\s,]+)",
            line
        )

        if weight_match:
            try:
                net_weight = float(
                    weight_match.group(2).replace(" ", "").replace(",", ".")
                )

                if net_weight < 100000:
                    current_weight = net_weight
            except:
                pass

        # ✅ HS CODE (само веднъж!)
        if "HS code" in line:
            code = re.search(r"HS code\s*:\s*(\d+)", line)

            if code:
                code_value = code.group(1)[:8]

                # ✅ ✅ ✅ ФИЛТЪР ПО ALLOWED_CODES
                if code_value not in ALLOWED_CODES:
                    continue

                # ✅ ЛОГИКА ЗА КОЛИЧЕСТВО
                if current_qty * units_in_box * liters_per_unit > 100000:
                    real_qty = current_qty
                else:
                    if units_in_box > 1 and liters_per_unit <= 5:
                        real_qty = current_qty * units_in_box
                    else:
                        real_qty = current_qty

                rows.append({
                    "Тарифен код": code_value,
                    "Количество": real_qty,
                    "wid": liters_per_unit,
                    "kolichestvo": round(real_qty * liters_per_unit, 3),
                    "тегло": round(current_weight, 3)
                })

                # ✅ RESET
                current_qty = 0
                current_weight = 0
                liters_per_unit = 0
                units_in_box = 1

    return pd.DataFrame(rows)
    
    # ======================================================
# ✅ AUTO MEGA FINAL
# ======================================================
def parse_auto_mega_excel(file):

    raw = pd.read_excel(file, header=None)

    header_row = None

    for i in range(len(raw)):

        row_text = " ".join(
            str(x)
            for x in raw.iloc[i]
            if pd.notna(x)
        ).upper()

        if (
            "TARIFF" in row_text
            and "DESCRIPTION" in row_text
        ):
            header_row = i
            break

    if header_row is None:
        st.error("❌ AUTO MEGA header не е намерен")
        return pd.DataFrame()

    df = pd.read_excel(
        file,
        header=header_row
    )

    df.columns = (
        df.columns.astype(str)
        .str.strip()
    )

    rows = []

    for _, row in df.iterrows():

        try:

            code = str(
                row["Tariff Code"]
            ).strip()

            code = re.sub(
                r"\D",
                "",
                code
            )[:8]

            if code not in ALLOWED_CODES:
                continue

            description = str(
                row["Description"]
            ).upper()

            qty = pd.to_numeric(
                row["Delivery"],
                errors="coerce"
            )

            net_weight = pd.to_numeric(
                row["wt./net"],
                errors="coerce"
            )

            item_weight = pd.to_numeric(
                row["wt./item"],
                errors="coerce"
            )

            if pd.isna(qty):
                continue

            if pd.isna(net_weight):
                continue

            wid = None

            # ✅ търси литраж в описанието
            match = re.search(
                r"(\d+(?:[.,]\d+)?)\s*L",
                description
            )

            if match:

                wid = float(
                    match.group(1)
                    .replace(",", ".")
                )

            # ✅ специални случаи
            if wid is None:

                if "TOYOTA SAE 5W40" in description:
                    wid = 5

                elif "AUTOMATIC TRANSMISSION OIL" in description:
                    wid = 1

                elif "TRANSMISSION OIL" in description:
                    wid = 1

            # ✅ fallback по wt./item
            if wid is None and pd.notna(item_weight):

                w = float(item_weight)

                if 0.90 <= w <= 0.999:
                    wid = 1

                elif 1.10 <= w <= 1.999:
                    wid = 2

                elif 2.10 <= w <= 2.999:
                    wid = 3

                elif 3.10 <= w <= 3.999:
                    wid = 4

                elif 4.10 <= w <= 4.999:
                    wid = 5

                elif 5.10 <= w <= 5.999:
                    wid = 6

                elif 17.00 <= w <= 18.999:
                    wid = 18

                elif 19.00 <= w <= 19.999:
                    wid = 20

                elif 55.00 <= w <= 59.999:
                    wid = 60

                elif 70.00 <= w <= 70.999:
                    wid = 80

                elif 170.00 <= w <= 170.999:
                    wid = 180

                elif 181.00 <= w <= 181.999:
                    wid = 200

                elif 207.00 <= w <= 208.999:
                    wid = 208

            if wid is None:
                continue

            rows.append({
                "Тарифен код": code,
                "Количество": qty,
                "wid": wid,
                "kolichestvo": qty * wid,
                "тегло": net_weight
            })

        except:
            continue

    if not rows:

        st.error("❌ AUTO MEGA parser не извлече данни")

        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = df_out.groupby(
        ["Тарифен код", "wid"],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df_out

 
# ======================================================
# ✅ FUCHS PDF - FIXED ML / M / PLA / SMALL PACKAGES
# ======================================================
def parse_fuchs(text):

    rows = []

    current_code = None
    current_wid = None
    current_qty = 0
    current_net = 0

    lines = text.split("\n")

    # ==================================================
    # ✅ ИЗВЛИЧАНЕ НА РАЗФАСОВКАТА ОТ MATERIAL
    # ==================================================
    def extract_fuchs_wid(line):

        txt = (
            str(line)
            .upper()
            .replace(",", ".")
        )

        txt = re.sub(
            r"\s+",
            " ",
            txt
        ).strip()

        # ==============================================
        # ✅ 500ML / 250ML / 125ML
        # ==============================================
        match = re.search(
            r"\b(\d+(?:\.\d+)?)\s*ML\b",
            txt
        )

        if match:
            return float(match.group(1)) / 1000

        # ==============================================
        # ✅ 500M AER / 125M PLA
        # M означава ML при малките опаковки
        # ==============================================
        match = re.search(
            r"\b(\d+(?:\.\d+)?)\s*M\s+(?:AER|PLA|CRT)\b",
            txt
        )

        if match:

            value = float(
                match.group(1)
            )

            if value <= 1000:
                return value / 1000

        # ==============================================
        # ✅ Стандартни литри
        # 1L PLA / 4L CUB / 20L PLA / 205L MET
        # ==============================================
        match = re.search(
            r"\b(\d+(?:\.\d+)?)\s*L\b",
            txt
        )

        if match:
            return float(
                match.group(1)
            )

        # ==============================================
        # ✅ FUCHS без буквата L
        # 20 PLA / 40 PLA / 60 PLA / 180 PLA
        # 200 PLA / 208 PLA
        # ==============================================
        match = re.search(
            r"\b(20|40|60|180|200|208)\s*PLA\b",
            txt
        )

        if match:
            return float(
                match.group(1)
            )

        # ==============================================
        # ✅ KG
        # ==============================================
        match = re.search(
            r"\b(\d+(?:\.\d+)?)\s*KG\b",
            txt
        )

        if match:
            return float(
                match.group(1)
            )

        # ==============================================
        # ✅ G / GR
        # ==============================================
        match = re.search(
            r"\b(\d+(?:\.\d+)?)\s*(?:G|GR)\b",
            txt
        )

        if match:
            return float(
                match.group(1)
            ) / 1000

        return None

    # ==================================================
    # ✅ ЗАПИС НА ТЕКУЩИЯ АРТИКУЛ
    # ==================================================
    def flush_article():

        nonlocal current_code
        nonlocal current_wid
        nonlocal current_qty
        nonlocal current_net

        if (
            current_code
            and current_wid is not None
            and current_qty > 0
            and current_net > 0
        ):

            kolichestvo = (
                current_qty * current_wid
            )

            rows.append({
                "Тарифен код": current_code,
                "Количество": current_qty,
                "wid": current_wid,
                "kolichestvo": round(kolichestvo, 3),
                "тегло": round(current_net, 3)
            })

        current_code = None
        current_wid = None
        current_qty = 0
        current_net = 0

    # ==================================================
    # ✅ ОБРАБОТКА НА PDF РЕДОВЕТЕ
    # ==================================================
    for line in lines:

        line = " ".join(
            str(line).split()
        )

        # ==============================================
        # ✅ НАЧАЛО НА НОВ АРТИКУЛ
        # ==============================================
        if re.search(
            r"\bMaterial\b",
            line,
            re.IGNORECASE
        ):

            flush_article()

            current_wid = extract_fuchs_wid(
                line
            )

        # ==============================================
        # ✅ МИТНИЧЕСКИ КОД
        # ==============================================
        if re.search(
            r"Commodity\s+Code",
            line,
            re.IGNORECASE
        ):

            code_match = re.search(
                r"Commodity\s+Code\D*(\d[\d\s.]*)",
                line,
                re.IGNORECASE
            )

            if code_match:

                code = re.sub(
                    r"\D",
                    "",
                    code_match.group(1)
                )[:8]

                if code in ALLOWED_CODES:
                    current_code = code
                else:
                    current_code = None

        # ==============================================
        # ✅ QUANTITY / NET / GROSS
        # Поддържа 432.00 EA / 365.90 KG
        # и числа със запетая
        # ==============================================
        if re.search(
            r"Quantity\s*/\s*net\s*/\s*gross",
            line,
            re.IGNORECASE
        ):

            quantity_match = re.search(
                r"([\d.,]+)\s*EA\s*/\s*([\d.,]+)\s*KG",
                line,
                re.IGNORECASE
            )

            if quantity_match:

                try:
                    qty_text = (
                        quantity_match.group(1)
                        .replace(",", "")
                    )

                    net_text = (
                        quantity_match.group(2)
                        .replace(",", "")
                    )

                    qty = float(qty_text)
                    net = float(net_text)

                    current_qty += qty
                    current_net += net

                except Exception:
                    continue

    # ✅ Последният артикул във файла
    flush_article()

    if not rows:
        st.error(
            "❌ FUCHS PDF parser не извлече данни"
        )
        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = df_out.groupby(
        [
            "Тарифен код",
            "wid"
        ],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df_out


# ======================================================
# ✅ FUCHS EXCEL
# ======================================================
def parse_fuchs_excel(file):

    try:
        df = pd.read_excel(
            file,
            engine="openpyxl"
        )

    except Exception as e:
        st.error(
            f"❌ FUCHS Excel файлът не може да бъде прочетен: {e}"
        )
        return pd.DataFrame()

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    description_col = None
    quantity_col = None
    code_col = None
    volume_col = None
    net_weight_col = None

    # ==================================================
    # ✅ НАМИРАНЕ НА КОЛОНИТЕ
    # ==================================================
    for col in df.columns:

        normalized_col = (
            str(col)
            .strip()
            .lower()
        )

        if normalized_col == "description":
            description_col = col

        elif normalized_col == "delivery quantity":
            quantity_col = col

        elif (
            "comm./imp. code" in normalized_col
            or "commodity code" in normalized_col
            or "customs code" in normalized_col
        ):
            code_col = col

        elif normalized_col == "volume":
            volume_col = col

        elif normalized_col == "net weight":
            net_weight_col = col

    required_columns = {
        "Description": description_col,
        "Delivery quantity": quantity_col,
        "Comm./imp. code no.": code_col,
        "Volume": volume_col,
        "Net Weight": net_weight_col
    }

    missing_columns = [
        name
        for name, column in required_columns.items()
        if column is None
    ]

    if missing_columns:
        st.error(
            "❌ FUCHS Excel: липсват колони: "
            + ", ".join(missing_columns)
        )
        return pd.DataFrame()

    # ==================================================
    # ✅ ИЗВЛИЧАНЕ НА РАЗФАСОВКАТА
    # ==================================================
    def extract_fuchs_excel_wid(
        description,
        quantity,
        volume
    ):

        txt = (
            str(description)
            .upper()
            .replace(",", ".")
            .replace("×", "X")
            .replace("Х", "X")
        )

        txt = re.sub(
            r"\s+",
            " ",
            txt
        ).strip()

        # ✅ ML
        match = re.search(
            r"\b(\d+(?:\.\d+)?)\s*ML\b",
            txt
        )

        if match:
            return float(match.group(1)) / 1000

        # ✅ 500M AER / 125M PLA
        match = re.search(
            r"\b(\d+(?:\.\d+)?)\s*M\s+(?:AER|PLA|CRT)\b",
            txt
        )

        if match:

            value = float(
                match.group(1)
            )

            if value <= 1000:
                return value / 1000

        # ✅ L
        match = re.search(
            r"\b(\d+(?:\.\d+)?)\s*L\b",
            txt
        )

        if match:
            return float(
                match.group(1)
            )

        # ✅ 20 PLA / 40 PLA / 60 PLA и т.н.
        match = re.search(
            r"\b(20|40|60|180|200|208)\s*PLA\b",
            txt
        )

        if match:
            return float(
                match.group(1)
            )

        # ✅ KG
        match = re.search(
            r"\b(\d+(?:\.\d+)?)\s*KG\b",
            txt
        )

        if match:
            return float(
                match.group(1)
            )

        # ✅ G / GR
        match = re.search(
            r"\b(\d+(?:\.\d+)?)\s*(?:G|GR)\b",
            txt
        )

        if match:
            return float(
                match.group(1)
            ) / 1000

        # ✅ FALLBACK: Volume / Delivery quantity
        if (
            quantity is not None
            and volume is not None
            and quantity > 0
            and volume > 0
        ):
            return volume / quantity

        return None

    rows = []

    # ==================================================
    # ✅ ОБРАБОТКА НА EXCEL РЕДОВЕТЕ
    # ==================================================
    for _, row in df.iterrows():

        try:
            description = str(
                row.get(description_col, "")
            ).strip()

            code = str(
                row.get(code_col, "")
            )

            code = re.sub(
                r"\D",
                "",
                code
            )[:8]

            if code not in ALLOWED_CODES:
                continue

            quantity = pd.to_numeric(
                row.get(quantity_col),
                errors="coerce"
            )

            volume = pd.to_numeric(
                row.get(volume_col),
                errors="coerce"
            )

            net_weight = pd.to_numeric(
                row.get(net_weight_col),
                errors="coerce"
            )

            if pd.isna(quantity):
                continue

            if pd.isna(net_weight):
                continue

            quantity = float(quantity)
            net_weight = float(net_weight)

            if quantity <= 0:
                continue

            if net_weight <= 0:
                continue

            if pd.isna(volume):
                volume = None
            else:
                volume = float(volume)

            wid = extract_fuchs_excel_wid(
                description=description,
                quantity=quantity,
                volume=volume
            )

            if wid is None or wid <= 0:
                continue

            if volume is not None and volume > 0:
                kolichestvo = volume
            else:
                kolichestvo = quantity * wid

            rows.append({
                "Тарифен код": code,
                "Количество": quantity,
                "wid": wid,
                "kolichestvo": kolichestvo,
                "тегло": net_weight
            })

        except Exception:
            continue

    if not rows:
        st.error(
            "❌ FUCHS Excel parser не извлече данни"
        )
        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = df_out.groupby(
        [
            "Тарифен код",
            "wid"
        ],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df_out
        
# ======================================================
# ✅ GASOLINE PDF (UNIVERSAL)
# ======================================================
def parse_gasoline(text):

    rows = []

    # ==========================================
    # ✅ С КОД
    # ==========================================

    pattern = re.compile(
        r'(\d+)\s*Liter.*?'
        r'([\d\.,]+)\s*'
        r'(\d+\s*x\s*\d+|\d+x\d+|\d+\s*Liter\s*Fass|\d+\s*Liter\s*Kanne).*?'
        r'Zolltarifnummer:\s*(\d{8})',
        re.IGNORECASE | re.DOTALL
    )

    for match in pattern.finditer(text):

        try:

            total_liters = float(match.group(1))

            weight = float(
                match.group(2)
                .replace(".", "")
                .replace(",", ".")
            )

            package = match.group(3)

            code = match.group(4)

            package_match = re.search(
                r'(\d+)\s*x\s*(\d+)',
                package,
                re.IGNORECASE
            )

            if package_match:

                wid = float(
                    package_match.group(2)
                )

            else:

                container_match = re.search(
                    r'(\d+)\s*Liter\s*(?:Fass|Kanne)',
                    package,
                    re.IGNORECASE
                )

                if not container_match:
                    continue

                wid = float(
                    container_match.group(1)
                )

            rows.append({
                "Тарифен код": code,
                "Количество": total_liters / wid,
                "wid": wid,
                "kolichestvo": total_liters,
                "тегло": weight
            })

        except:
            continue

    # ==========================================
    # ✅ БЕЗ КОД -> 00000000
    # ==========================================

    no_code_pattern = re.compile(
        r'(\d+)\s*Liter\s+'
        r'.*?'
        r'([\d\.,]+)\s+'
        r'(\d+\s*x\s*\d+|\d+x\d+)'
        r'\s*Liter\s+im\s+Karton',
        re.IGNORECASE | re.DOTALL
    )

    for match in no_code_pattern.finditer(text):

        try:

            total_liters = float(
                match.group(1)
            )

            weight = float(
                match.group(2)
                .replace(".", "")
                .replace(",", ".")
            )

            package = match.group(3)

            wid = float(
                re.search(
                    r'x\s*(\d+)',
                    package,
                    re.IGNORECASE
                ).group(1)
            )

            broj = total_liters / wid

            already_exists = any(
                abs(r["kolichestvo"] - total_liters) < 0.01
                for r in rows
            )

            if not already_exists:

                rows.append({
                    "Тарифен код": "00000000",
                    "Количество": broj,
                    "wid": wid,
                    "kolichestvo": total_liters,
                    "тегло": weight
                })

        except:
            continue

    if not rows:

        st.error(
            "❌ GASOLINE parser не извлече данни"
        )

        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = df_out.groupby(
        ["Тарифен код", "wid"],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df_out
# ======================================================
# ✅ CHEMPIOIL PDF
# ======================================================
def parse_chempioil_pdf(text):

    rows = []

    lines = text.split("\n")

    for line in lines:

        line = " ".join(line.split())

        try:

            # ✅ HS CODE
            hs_match = re.search(
                r'(\d{4}\.\d{2}\.\d{2}\.\d{2})',
                line
            )

            if not hs_match:
                continue

            code = re.sub(
                r"\D",
                "",
                hs_match.group(1)
            )[:8]

            if code not in ALLOWED_CODES:
                continue

            # ✅ Quantity + Net + Gross
            m = re.search(
                r'(\d+(?:\.\d+)?)\s+SZT\s+'
                r'(\d+(?:\.\d+)?)\s+'
                r'(\d+(?:\.\d+)?)\s+'
                r'\d{4}\.\d{2}\.\d{2}\.\d{2}',
                line
            )

            if not m:
                continue

            qty = float(m.group(1))
            net_weight = float(m.group(2))

            # ✅ WID
            wid = None

            upper_line = (
                line
                .upper()
                .replace(",", ".")
            )

            # ✅ L
            m = re.search(
                r'(\d+(?:\.\d+)?)\s*L\b',
                upper_line
            )

            if m:
                wid = float(m.group(1))

            # ✅ KG
            if wid is None:

                m = re.search(
                    r'(\d+(?:\.\d+)?)\s*KG\b',
                    upper_line
                )

                if m:
                    wid = float(m.group(1))

            # ✅ G / GR
            if wid is None:

                m = re.search(
                    r'(\d+(?:\.\d+)?)\s*(?:GR|G)\b',
                    upper_line
                )

                if m:
                    wid = float(m.group(1)) / 1000

            if wid is None:
                continue

            rows.append({
                "Тарифен код": code,
                "Количество": qty,
                "wid": wid,
                "kolichestvo": qty * wid,
                "тегло": net_weight
            })

        except:

            continue

    if not rows:

        st.error("❌ CHEMPIOIL PDF parser не извлече данни")

        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = df_out.groupby(
        ["Тарифен код", "wid"],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df_out


# ======================================================
# ✅ CHEMPIOIL EXCEL (ALL FORMATS + CODE FALLBACK FIX)
# ======================================================
def parse_chempioil_excel(file):

    raw = pd.read_excel(
        file,
        header=None
    )

    header_row = None

    for i in range(len(raw)):

        row_text = " ".join(
            str(x)
            for x in raw.iloc[i]
            if pd.notna(x)
        ).upper()

        if (
            ("HS CODE" in row_text and "NAME" in row_text)
            or
            ("PRODUCT NAME" in row_text and "CN" in row_text)
        ):
            header_row = i
            break

    if header_row is None:

        st.error("❌ CHEMPIOIL header не е намерен")

        return pd.DataFrame()

    df = pd.read_excel(
        file,
        header=header_row
    )

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    # ==================================================
    # ✅ FORMAT 1
    # ==================================================
    if "HS Code" in df.columns:

        product_code_col = "Code"
        hs_code_col = "HS Code"
        name_col = "Name"
        qty_col = "Quantity"
        net_col = "Net weight"

    # ==================================================
    # ✅ FORMAT 2
    # ==================================================
    elif "CN" in df.columns:

        product_code_col = "Product Code"
        hs_code_col = "CN"
        name_col = "Product Name"
        qty_col = "Quantity"
        net_col = "Total Weight (NET):"

    else:

        st.error("❌ CHEMPIOIL: непознат Excel формат")

        return pd.DataFrame()

    rows = []

    for _, row in df.iterrows():

        try:

            description = str(
                row.get(name_col, "")
            ).upper().replace(",", ".")

            product_code = str(
                row.get(product_code_col, "")
            ).upper().strip()

            code = str(
                row.get(hs_code_col, "")
            )

            code = re.sub(
                r"\D",
                "",
                code
            )[:8]

            if code not in ALLOWED_CODES:
                continue

            qty = pd.to_numeric(
                row.get(qty_col),
                errors="coerce"
            )

            net_weight = pd.to_numeric(
                row.get(net_col),
                errors="coerce"
            )

            if pd.isna(qty):
                continue

            if pd.isna(net_weight):
                continue

            wid = None

            # ==================================================
            # ✅ 1) Първо търсим разфасовка в NAME
            # ==================================================
            m = re.search(
                r'(\d+(?:\.\d+)?)\s*L\b',
                description
            )

            if m:
                wid = float(m.group(1))

            # ==================================================
            # ✅ 2) KG
            # ==================================================
            if wid is None:

                m = re.search(
                    r'(\d+(?:\.\d+)?)\s*KG\b',
                    description
                )

                if m:
                    wid = float(m.group(1))

            # ==================================================
            # ✅ 3) G / GR
            # ==================================================
            if wid is None:

                m = re.search(
                    r'(\d+(?:\.\d+)?)\s*(?:GR|G)\b',
                    description
                )

                if m:
                    wid = float(m.group(1)) / 1000

            # ==================================================
            # ✅ 4) FALLBACK ОТ PRODUCT CODE
            #
            # CH8901-20  -> 20L
            # CH8901-60  -> 60L
            # CH8901-DR  -> 208L
            # CH9729-60  -> 60L
            # CH9729-DR  -> 208L
            # ==================================================
            if wid is None:

                m = re.search(
                    r"-(\d+)$",
                    product_code
                )

                if m:
                    wid = float(m.group(1))

            if wid is None and product_code.endswith("-DR"):

                wid = 208

            if wid is None:
                continue

            rows.append({
                "Тарифен код": code,
                "Количество": float(qty),
                "wid": wid,
                "kolichestvo": float(qty) * wid,
                "тегло": float(net_weight)
            })

        except:

            continue

    if not rows:

        st.error("❌ CHEMPIOIL Excel parser не извлече данни")

        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = df_out.groupby(
        ["Тарифен код", "wid"],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df_out
# ======================================================
# ✅ VALVOLINE EXCEL
# ======================================================
def parse_valvoline_excel(file):

    xl = pd.ExcelFile(file)

    sheet_name = None

    for s in xl.sheet_names:

        if "PL" in str(s).upper():
            sheet_name = s
            break

    if sheet_name is None:
        sheet_name = xl.sheet_names[0]

    raw = pd.read_excel(
        file,
        sheet_name=sheet_name,
        header=None
    )

    header_row = None

    for i in range(len(raw)):

        row_text = " ".join(
            str(x)
            for x in raw.iloc[i]
            if pd.notna(x)
        ).upper()

        if (
            "TARIFF NO." in row_text
            and
            "PACKAGING" in row_text
            and
            "NET KG" in row_text
        ):
            header_row = i
            break

    if header_row is None:
        st.error("❌ VALVOLINE header не е намерен")
        return pd.DataFrame()

    df = pd.read_excel(
        file,
        sheet_name=sheet_name,
        header=header_row
    )

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    rows = []

    for _, row in df.iterrows():

        try:

            tariff = str(
                row["Tariff No."]
            )

            code = re.sub(
                r"\D",
                "",
                tariff
            )[:8]

            if code not in ALLOWED_CODES:
                continue

            packaging = str(
                row["Packaging"]
            ).upper().replace(",", ".")

            description = str(
                row["Description"]
            ).upper().replace(",", ".")

            uom = str(
                row["UoM"]
            ).strip().lower()

            qty = pd.to_numeric(
                row["Qty"],
                errors="coerce"
            )

            net_weight = pd.to_numeric(
                row["Net Kg"],
                errors="coerce"
            )

            packages = pd.to_numeric(
                row["No. of packages"],
                errors="coerce"
            )

            if pd.isna(qty):
                continue

            if pd.isna(net_weight):
                continue

            if pd.isna(packages):
                continue

            wid = None

            # =====================================
            # ✅ CASE
            # =====================================
            if uom == "case":

                # ✅ Хваща:
                # 12 x 500 ML
                # 12/500 ML
                # 12 x 300 ML
                # 12 x 1 L
                # 4 x 5 L
                # 4/5 L
                case_match = re.search(
                    r'(\d+)\s*(?:[Xx]|/)\s*(\d+(?:[.,]\d+)?)\s*(ML|L|KG|G|GR)',
                    packaging
                )

                # ✅ fallback към Description
                if not case_match:

                    case_match = re.search(
                        r'(\d+)\s*(?:[Xx]|/)\s*(\d+(?:[.,]\d+)?)\s*(ML|L|KG|G|GR)',
                        description
                    )

                if case_match:

                    units_per_case = float(
                        case_match.group(1)
                    )

                    wid = float(
                        case_match.group(2)
                        .replace(",", ".")
                    )

                    unit_type = case_match.group(3)

                    if unit_type == "ML":
                        wid = wid / 1000

                    elif unit_type in ["G", "GR"]:
                        wid = wid / 1000

                    # ✅ Брой = кашони x бройки в кашон
                    broj = packages * units_per_case

                    # ✅ Литри/количество = брой x разфасовка
                    colic = broj * wid

                else:

                    continue

            # =====================================
            # ✅ LIT / KG / G / ML
            # =====================================
            else:

                # ✅ нормален Packaging
                # 208 L, 60 L, 20 L, 18 KG, 500 ML
                m = re.search(
                    r'(\d+(?:[.,]\d+)?)\s*(ML|L|KG|G|GR)',
                    packaging
                )

                if m:

                    wid = float(
                        m.group(1)
                        .replace(",", ".")
                    )

                    unit_type = m.group(2)

                    if unit_type == "ML":
                        wid = wid / 1000

                    elif unit_type in ["G", "GR"]:
                        wid = wid / 1000

                # ✅ ако Packaging е само число
                elif re.fullmatch(
                    r'\d+(?:[.,]\d+)?',
                    packaging.strip()
                ):

                    m = re.search(
                        r'(\d+(?:[.,]\d+)?)\s*(ML|L|KG|G|GR)',
                        description
                    )

                    if m:

                        wid = float(
                            m.group(1)
                            .replace(",", ".")
                        )

                        unit_type = m.group(2)

                        if unit_type == "ML":
                            wid = wid / 1000

                        elif unit_type in ["G", "GR"]:
                            wid = wid / 1000

                if wid is None:
                    continue

                # При UoM = Lit / kg логиката остава както при теб
                broj = packages
                colic = qty

            rows.append({
                "Тарифен код": code,
                "Количество": broj,
                "wid": wid,
                "kolichestvo": colic,
                "тегло": net_weight
            })

        except:

            continue

    if not rows:
        st.error("❌ VALVOLINE parser не извлече данни")
        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = df_out.groupby(
        ["Тарифен код", "wid"],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df_out
# ======================================================
# ✅ EMINIA EXCEL
# ======================================================
def parse_eminia_excel(file):

    df = pd.read_excel(file)

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    # ==========================================
    # ✅ намиране на колони
    # ==========================================

    desc_col = None
    code_col = None
    qty_col = None
    weight_col = None

    for col in df.columns:

        c = str(col).strip().lower()

        if "description" in c:
            desc_col = col

        elif "customs" in c and "code" in c:
            code_col = col

        elif c == "qty":
            qty_col = col

        elif (
            "total net weight" in c
            or "total weight" in c
        ):
            weight_col = col

    # ==========================================
    # ✅ fallback по позиция
    # D,E,H,L
    # ==========================================

    cols = list(df.columns)

    if desc_col is None and len(cols) > 3:
        desc_col = cols[3]

    if code_col is None and len(cols) > 4:
        code_col = cols[4]

    if qty_col is None and len(cols) > 6:
        qty_col = cols[6]

    if weight_col is None:

        # L колона
        if len(cols) > 11:
            weight_col = cols[11]

        # K колона
        elif len(cols) > 10:
            weight_col = cols[10]

    rows = []

    for _, row in df.iterrows():

        try:

            description = str(
                row[desc_col]
            )

            code = str(
                row[code_col]
            )

            code = re.sub(
                r"\D",
                "",
                code
            )[:8]

            if code not in ALLOWED_CODES:
                continue

            # ==================================
            # ✅ WID
            # ==================================

            wid = None

            txt = description.upper().replace(",", ".")

            m = re.search(
                r'(\d+(?:\.\d+)?)\s*L\b',
                txt
            )

            if m:
                wid = float(m.group(1))

            if wid is None:

                m = re.search(
                    r'(\d+(?:\.\d+)?)\s*KG\b',
                    txt
                )

                if m:
                    wid = float(m.group(1))

            if wid is None:

                m = re.search(
                    r'(\d+(?:\.\d+)?)\s*(?:GR|G)\b',
                    txt
                )

                if m:
                    wid = float(m.group(1)) / 1000

            # не е масло
            if wid is None:
                continue

            qty = pd.to_numeric(
                row[qty_col],
                errors="coerce"
            )

            total_weight = pd.to_numeric(
                row[weight_col],
                errors="coerce"
            )

            if pd.isna(qty):
                continue

            if pd.isna(total_weight):
                continue

            rows.append({
                "Тарифен код": code,
                "Количество": qty,
                "wid": wid,

                # както поиска
                "kolichestvo": total_weight,

                "тегло": total_weight
            })

        except:
            continue

    if not rows:
        st.error("❌ EMINIA parser не извлече данни")
        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = df_out.groupby(
        ["Тарифен код", "wid"],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df_out
  
# ======================================================
# ✅ FLUKAR (EXCEL ONLY ✅)
# ======================================================
def parse_flukar_excel(file):

    df_raw = pd.read_excel(file, header=None)

    header_row = None

    # ✅ намираме header ред
    for i in range(len(df_raw)):
        row = df_raw.iloc[i]

        if any("cn" in str(cell).lower() for cell in row if pd.notna(cell)):
            header_row = i
            break

    if header_row is None:
        st.error("❌ Не може да се намери header ред (CN code)")
        return pd.DataFrame()

    df = pd.read_excel(file, header=header_row)
    df.columns = df.columns.astype(str).str.strip()

    # ✅ извличаме само нужните колони
    result = pd.DataFrame()

    for col in df.columns:
        c = col.lower()

        if "cn" in c:
            result["Тарифен код"] = df[col]

        elif "quantity" in c or "pcs" in c or "колич" in c:
            result["Количество"] = df[col]

        elif "capacity" in c or "package" in c:
            if "wid" not in result.columns:
                result["wid"] = df[col]

        elif "liter" in c:
            result["kolichestvo"] = df[col]

        elif "nett" in c or "net" in c or "тегло" in c:
            result["тегло"] = df[col]

    # ✅ проверки
    required = ["Тарифен код", "Количество", "wid", "тегло"]

    for col in required:
        if col not in result.columns:
            st.error(f"❌ Липсва колона: {col}")
            return pd.DataFrame()

    # ✅ cleaning
    result = result.dropna(subset=["Тарифен код"])

    result["Количество"] = pd.to_numeric(result["Количество"], errors="coerce")
    result["wid"] = pd.to_numeric(result["wid"], errors="coerce")
    result["тегло"] = pd.to_numeric(result["тегло"], errors="coerce")

    if "kolichestvo" not in result.columns:
        result["kolichestvo"] = result["Количество"] * result["wid"]
    else:
        result["kolichestvo"] = pd.to_numeric(result["kolichestvo"], errors="coerce")

    # ✅ ✅ 🔥 ВАЖНО — ROUND САМО НА ТЕГЛО (като FLUKAR)

    result = result.dropna(subset=["Количество", "wid", "тегло"])

    # ✅ group
    result = result.groupby(
        ["Тарифен код", "wid"],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return result

from decimal import Decimal

# ======================================================
# ✅ FINAL REPORT (NO ROUNDING FOR FLUKAR)
# ======================================================
def build_final_report(df, supplier):

    def precise_sum(values):

        total = Decimal("0")

        for x in values:

            try:

                if pd.notna(x) and str(x).strip() != "":
                    total += Decimal(str(x))

            except:
                continue

        return total

    # ==================================================
    # ✅ FLUKAR
    # ==================================================
    if supplier == "FLUKAR":

        grouped = df.groupby(
            ["Тарифен код", "wid"],
            as_index=False
        ).agg({
            "Количество": "sum",
            "kolichestvo": "sum",
            "тегло": list
        })

        rows = []

        for code, group in grouped.groupby("Тарифен код"):

            for _, r in group.iterrows():

                weight_sum = precise_sum(
                    r["тегло"]
                )

                rows.append({
                    "Тарифен код": r["Тарифен код"],
                    "wid": r["wid"],
                    "Количество": r["Количество"],
                    "kolichestvo": r["kolichestvo"],
                    "тегло": float(weight_sum)
                })

            code_sum = precise_sum(
                x
                for sublist in group["тегло"]
                for x in sublist
            )

            rows.append({
                "Тарифен код": str(code) + " -",
                "wid": "",
                "Количество": "",
                "kolichestvo": group["kolichestvo"].sum(),
                "тегло": float(code_sum)
            })

            rows.append({
                "Тарифен код": "",
                "wid": "",
                "Количество": "",
                "kolichestvo": "",
                "тегло": ""
            })

        total_sum = precise_sum(
            x
            for sublist in grouped["тегло"]
            for x in sublist
        )

        rows.append({
            "Тарифен код": "GRAND TOTAL",
            "wid": "",
            "Количество": "",
            "kolichestvo": grouped["kolichestvo"].sum(),
            "тегло": float(total_sum)
        })

        return pd.DataFrame(rows)

    # ==================================================
    # ✅ ВСИЧКИ ОСТАНАЛИ ДОСТАВЧИЦИ
    # ==================================================
    else:

        grouped = df.groupby(
            ["Тарифен код", "wid"],
            as_index=False
        ).agg({
            "Количество": "sum",
            "kolichestvo": "sum",
            "тегло": "sum"
        })

        rows = []

        for code, group in grouped.groupby("Тарифен код"):

            for _, r in group.iterrows():
                rows.append(r.to_dict())

            rows.append({
                "Тарифен код": str(code) + " -",
                "wid": "",
                "Количество": "",
                "kolichestvo": group["kolichestvo"].sum(),
                "тегло": group["тегло"].sum()
            })

            rows.append({
                "Тарифен код": "",
                "wid": "",
                "Количество": "",
                "kolichestvo": "",
                "тегло": ""
            })

        rows.append({
            "Тарифен код": "GRAND TOTAL",
            "wid": "",
            "Количество": "",
            "kolichestvo": grouped["kolichestvo"].sum(),
            "тегло": grouped["тегло"].sum()
        })

        return pd.DataFrame(rows)

# ======================================================
# ✅ AMTRA EXCEL
# ======================================================
def parse_amtra_excel(file):

    raw = pd.read_excel(
        file,
        sheet_name=1,
        header=None
    )

    header_row = None

    for i in range(len(raw)):

        row_text = " ".join(
            str(x)
            for x in raw.iloc[i]
            if pd.notna(x)
        ).upper()

        if (
            "PRODUCT NAME" in row_text
            and
            "CODE CN" in row_text
        ):
            header_row = i
            break

    if header_row is None:
        st.error("❌ AMTRA header не е намерен")
        return pd.DataFrame()

    df = pd.read_excel(
        file,
        sheet_name=1,
        header=header_row
    )

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    rows = []

    for _, row in df.iterrows():

        try:

            description = str(df.iloc[_ , 2])
            qty = pd.to_numeric(df.iloc[_ , 3], errors="coerce")
            code = str(df.iloc[_ , 6])
            net_weight = pd.to_numeric(df.iloc[_ , 9], errors="coerce")

            code = re.sub(r"\D", "", code)[:8]

            if code not in ALLOWED_CODES:
                continue

            txt = re.sub(
                r"\([^)]*\)",
                "",
                description
            ).upper()

            wid = None

            m = re.search(r"(\d+(?:[.,]\d+)?)\s*ML", txt)
            if m:
                wid = float(m.group(1).replace(",", ".")) / 1000

            if wid is None:
                m = re.search(r"(\d+(?:[.,]\d+)?)\s*L\b", txt)
                if m:
                    wid = float(m.group(1).replace(",", "."))

            if wid is None:
                m = re.search(r"(\d+(?:[.,]\d+)?)\s*KG", txt)
                if m:
                    wid = float(m.group(1).replace(",", "."))

            if wid is None:
                m = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:GR|G)\b", txt)
                if m:
                    wid = float(m.group(1).replace(",", ".")) / 1000

            if wid is None:
                continue

            if pd.isna(qty):
                continue

            if pd.isna(net_weight):
                continue

            rows.append({
                "Тарифен код": code,
                "Количество": qty,
                "wid": wid,
                "kolichestvo": qty * wid,
                "тегло": net_weight
            })

        except:
            continue

    if not rows:
        st.error("❌ AMTRA parser не извлече данни")
        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = df_out.groupby(
        ["Тарифен код", "wid"],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df_out
    # ======================================================
# ✅ ELROMI RONAX
# ======================================================
def parse_elromi_excel(file):

    xl = pd.ExcelFile(file)

    sheet_name = None

    for s in xl.sheet_names:

        if str(s).strip().upper() == "INHALT":
            sheet_name = s
            break

    if sheet_name is None:

        if len(xl.sheet_names) > 1:
            sheet_name = xl.sheet_names[1]
        else:
            sheet_name = xl.sheet_names[0]

    df = pd.read_excel(
        file,
        sheet_name=sheet_name
    )

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    rows = []

    for _, row in df.iterrows():

        try:

            code = str(row["tariff"])

            code = re.sub(
                r"\D",
                "",
                code
            )[:8]

            if code not in ALLOWED_CODES:
                continue

            qty = pd.to_numeric(
                row["Menge"],
                errors="coerce"
            )

            unit_weight = pd.to_numeric(
                row["Gewicht, kg"],
                errors="coerce"
            )

            if pd.isna(unit_weight):
                continue

            # ✅ игнорирай всичко под 0.50 кг
            if unit_weight < 0.50:
                continue

            if unit_weight <= 0:
                continue

            wid = max(
                1,
                round(float(unit_weight))
            )

            rows.append({
                "Тарифен код": code,
                "Количество": qty,
                "wid": wid,
                "kolichestvo": qty * wid,
                "тегло": qty * unit_weight
            })

        except:
            continue

    if not rows:
        st.error("❌ ELROMI parser не извлече данни")
        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = df_out.groupby(
        ["Тарифен код", "wid"],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df_out
 # ======================================================
# ✅ BRECHMANN
# ======================================================
def parse_brechmann_excel(file):

    # търси реда с истинските заглавия
    raw = pd.read_excel(
        file,
        header=None
    )

    header_row = None

    for idx in range(len(raw)):

        row_text = " ".join(
            [str(x) for x in raw.iloc[idx].tolist()]
        )

        if (
            "HS Code" in row_text
            and "Qty" in row_text
        ):
            header_row = idx
            break

    if header_row is None:
        st.error("❌ Не намерих header на BRECHMANN")
        return pd.DataFrame()

    df = pd.read_excel(
        file,
        header=header_row
    )

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    rows = []

    for _, row in df.iterrows():

        try:

            # -------------------------------------------------
            # CODE
            # -------------------------------------------------
            code = str(
                row.get("HS Code", "")
            )

            code = re.sub(
                r"\D",
                "",
                code
            )[:8]

            if not code:
                continue

            if code not in ALLOWED_CODES:
                continue

            # -------------------------------------------------
            # BROJ
            # -------------------------------------------------
            qty = pd.to_numeric(
                row.get("Qty"),
                errors="coerce"
            )

            if pd.isna(qty):
                continue

            # -------------------------------------------------
            # NET WEIGHT
            # -------------------------------------------------
            net_weight = pd.to_numeric(
                row.get("Net Weight"),
                errors="coerce"
            )

            if pd.isna(net_weight):
                net_weight = 0

            # -------------------------------------------------
            # WID
            # -------------------------------------------------
            package = str(
                row.get(
                    "Gebindegröße Öl",
                    ""
                )
            ).upper()

            wid = None

            # 12x1L -> 1
            m = re.search(
                r'X\s*(\d+(?:[.,]\d+)?)\s*L',
                package
            )

            if m:
                wid = float(
                    m.group(1).replace(",", ".")
                )

            # 4x20KG -> 20
            if wid is None:

                m = re.search(
                    r'X\s*(\d+(?:[.,]\d+)?)\s*KG',
                    package
                )

                if m:
                    wid = float(
                        m.group(1).replace(",", ".")
                    )

            # 12x500G -> 0.5
            if wid is None:

                m = re.search(
                    r'X\s*(\d+(?:[.,]\d+)?)\s*G',
                    package
                )

                if m:
                    wid = (
                        float(
                            m.group(1).replace(",", ".")
                        )
                        / 1000
                    )

            if wid is None:
                continue

            # -------------------------------------------------
            # REPORT ROW
            # -------------------------------------------------
            rows.append({
                "Тарифен код": code,
                "Количество": qty,
                "wid": wid,
                "kolichestvo": qty * wid,
                "тегло": qty * net_weight
            })

        except Exception:
            continue

    if not rows:
        st.error(
            "❌ BRECHMANN parser не извлече данни"
        )
        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = (
        df_out.groupby(
            ["Тарифен код", "wid"],
            as_index=False
        )
        .agg({
            "Количество": "sum",
            "kolichestvo": "sum",
            "тегло": "sum"
        })
    )

    return df_out
# ======================================================
# ✅ FEBI EXCEL
# ======================================================
def parse_febi_excel(file):

    try:

        df = pd.read_excel(
            file,
            engine="openpyxl"
        )

    except:

        st.error(
            "❌ FEBI файлът трябва да бъде записан като XLSX"
        )

        return pd.DataFrame()

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    rows = []

    for _, row in df.iterrows():

        try:

            code = str(
                row["HS-Code"]
            )

            code = re.sub(
                r"\D",
                "",
                code
            )[:8]

            if code not in ALLOWED_CODES:
                continue

            qty = pd.to_numeric(
                row["Quantity"],
                errors="coerce"
            )

            if pd.isna(qty):
                continue

            net_weight = pd.to_numeric(
                row["Net weight"],
                errors="coerce"
            )

            if pd.isna(net_weight):
                continue

            description = str(
                row.get("Description", "")
            ).upper()

            customer_material = str(
                row.get("Customer material", "")
            ).upper()

            # ✅ ВАЖНО:
            # нормализира латинско X, кирилско Х и символ ×
            customer_material = (
                customer_material
                .replace("Х", "X")
                .replace("×", "X")
                .replace(",", ".")
            )

            wid = None
            real_qty = float(qty)

            # ==================================================
            # ✅ CASE 1:
            # X6 / X7 / X8 / X12 и т.н.
            #
            # Това означава:
            # 1 кашон/комплект съдържа X броя по 1L
            #
            # Пример:
            # Quantity 8 + X6 = 48 бр. по 1L
            # ==================================================
            x_pack_match = re.search(
                r"\bX\s*(\d+)\b",
                customer_material
            )

            if x_pack_match:

                multiplier = int(
                    x_pack_match.group(1)
                )

                wid = 1

                real_qty = float(qty) * multiplier

            # ==================================================
            # ✅ CASE 2:
            # 1PC=1L или 1PC = 5L
            #
            # Ако няма X6/X7/X8, тогава хващаме директно = 1L / = 5L
            # ==================================================
            if wid is None:

                pc_match = re.search(
                    r"=\s*(\d+(?:\.\d+)?)\s*L",
                    customer_material
                )

                if pc_match:

                    wid = float(
                        pc_match.group(1)
                    )

                    real_qty = float(qty)

            # ==================================================
            # ✅ CASE 3:
            # директно написано 1L / 4L / 5L
            #
            # Пример:
            # Customer material = 4l
            # Quantity = 8
            # => 8 бр. x 4L
            # ==================================================
            if wid is None:

                liter_match = re.search(
                    r"\b(\d+(?:\.\d+)?)\s*L\b",
                    customer_material
                )

                if liter_match:

                    wid = float(
                        liter_match.group(1)
                    )

                    real_qty = float(qty)

            # ==================================================
            # ✅ CASE 4:
            # KG
            # ==================================================
            if wid is None:

                kg_match = re.search(
                    r"\b(\d+(?:\.\d+)?)\s*KG\b",
                    customer_material
                )

                if kg_match:

                    wid = float(
                        kg_match.group(1)
                    )

                    real_qty = float(qty)

            # ==================================================
            # ✅ CASE 5:
            # G / GR
            # ==================================================
            if wid is None:

                g_match = re.search(
                    r"\b(\d+(?:\.\d+)?)\s*(?:G|GR)\b",
                    customer_material
                )

                if g_match:

                    wid = float(
                        g_match.group(1)
                    ) / 1000

                    real_qty = float(qty)

            if wid is None:
                continue

            rows.append({
                "Тарифен код": code,
                "Количество": real_qty,
                "wid": wid,
                "kolichestvo": real_qty * wid,
                "тегло": float(net_weight)
            })

        except:

            continue

    if not rows:

        st.error(
            "❌ FEBI parser не извлече данни"
        )

        return pd.DataFrame()

    df_out = pd.DataFrame(rows)

    df_out = df_out.groupby(
        ["Тарифен код", "wid"],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df_out

# ======================================================
# ✅ ORLEN (EXCEL) - MATERIAL PACKAGING FIX
# ======================================================

def parse_orlen_excel(file):

    df = pd.read_excel(file)

    df.columns = df.columns.astype(str).str.strip()

    rename_map = {}

    for col in df.columns:

        c = col.lower().strip()

        if "code cn" in c or "cn code" in c:
            rename_map[col] = "Тарифен код"

        elif c == "qty" or "qty" in c:
            rename_map[col] = "Количество"

        elif "net weight" in c:
            rename_map[col] = "тегло"

        elif "material" in c:
            rename_map[col] = "material"

    df = df.rename(columns=rename_map)

    required = [
        "Тарифен код",
        "Количество",
        "тегло",
        "material"
    ]

    for col in required:

        if col not in df.columns:

            st.error(
                f"❌ ORLEN: липсва колона {col}"
            )

            return pd.DataFrame()

    # ==================================================
    # ✅ CLEAN CODE CN
    # ==================================================

    df["Тарифен код"] = (
        df["Тарифен код"]
        .astype(str)
        .str.replace(r"\D", "", regex=True)
        .str[:8]
    )

    df = df[
        df["Тарифен код"].isin(ALLOWED_CODES)
    ]

    # ==================================================
    # ✅ EXTRACT PACKAGING FROM MATERIAL
    # ==================================================

    def extract_wid(material):

        txt = str(material).upper()

        txt = txt.replace(",", ".")
        txt = txt.replace("×", "X")
        txt = txt.replace("Х", "X")

        txt = re.sub(
            r"\s+",
            " ",
            txt
        ).strip()

        # ==================================================
        # ✅ ML
        # Examples:
        # 100ML, 400ML, 500ML, 300ML, 200ML
        # 0.100ML, 0.400ML, 0.300ML
        # ==================================================

        ml_matches = re.findall(
            r"(\d+(?:\.\d+)?)\s*ML\b",
            txt
        )

        if ml_matches:

            value = float(
                ml_matches[-1]
            )

            if value <= 1:
                return value

            return value / 1000

                # ==================================================
        # ✅ LITERS / PLA
        # Examples:
        # 1L PLA
        # 4L CUB
        # 20L PLA
        # 205L MET
        # 20 PLA
        # 40 PLA
        # 60 PLA
        # 180 PLA
        # 200 PLA
        # 208 PLA
        # ==================================================

        # ==================================================
        # ✅ Първо стандартните L
        # ==================================================
        liter_matches = re.findall(
            r"(\d+(?:\.\d+)?)\s*L\b",
            txt
        )
        
        if liter_matches:
        
            return float(
                liter_matches[-1]
            )
        
        # ==================================================
        # ✅ После 20 PLA / 60 PLA / 208 PLA
        # ==================================================
        pla_match = re.search(
            r"\b(\d+(?:\.\d+)?)\s*PLA\b",
            txt
        )
        
        if pla_match:
        
            return float(
                pla_match.group(1)
            )

        # ==================================================
        # ✅ KG
        # Examples:
        # 4KG, 4.5KG, 5KG, 8KG, 9KG, 17KG
        # P9KG, H17KG, P4.5KG
        # ==================================================

        kg_matches = re.findall(
            r"(\d+(?:\.\d+)?)\s*KG\b",
            txt
        )

        if kg_matches:

            return float(
                kg_matches[-1]
            )

        # ==================================================
        # ✅ GRAMS
        # Examples:
        # 400G, 500G, 300G, 200G
        # 0.600G, 0.400G, 0.800G
        # K400G
        # ==================================================

        gram_matches = re.findall(
            r"(\d+(?:\.\d+)?)\s*G\b",
            txt
        )

        if gram_matches:

            value = float(
                gram_matches[-1]
            )

            if value <= 1:
                return value

            return value / 1000

        return None

    df["wid"] = df["material"].apply(
        extract_wid
    )

    df = df[
        df["wid"].notna()
    ]

    # ==================================================
    # ✅ NUMERIC COLUMNS
    # ==================================================

    df["Количество"] = pd.to_numeric(
        df["Количество"],
        errors="coerce"
    )

    df["тегло"] = pd.to_numeric(
        df["тегло"],
        errors="coerce"
    )

    df = df.dropna(
        subset=[
            "Тарифен код",
            "Количество",
            "wid",
            "тегло"
        ]
    )

    # ==================================================
    # ✅ CALCULATION
    # Qty × packaging
    # ==================================================

    df["kolichestvo"] = (
        df["Количество"] * df["wid"]
    )

    # ==================================================
    # ✅ GROUP RESULT
    # ==================================================

    df = df.groupby(
        [
            "Тарифен код",
            "wid"
        ],
        as_index=False
    ).agg({
        "Количество": "sum",
        "kolichestvo": "sum",
        "тегло": "sum"
    })

    return df
    # ======================================================
# ✅ PDF INVOICE TO EXCEL - VENDOR CROSS REFERENCE
# Cross-Reference Type No. / Cross-Reference No. / Qty / Price
# ======================================================

CROSS_REFERENCE_ZIPS = [
    "cross_ref_1.zip",
    "cross_ref_2.zip",
    "cross_ref_3.zip"
]


def clean_value(value):
    if value is None or pd.isna(value):
        return ""
    value = str(value).strip()
    if value.lower() in {"", "nan", "none", "null"}:
        return ""
    if re.fullmatch(r"\d+\.0", value):
        value = value[:-2]
    return value


def normalize_key(value):
    value = clean_value(value).upper()
    value = unicodedata.normalize("NFKC", value)
    return re.sub(r"[^A-ZА-Я0-9]", "", value)


def normalize_header_name(value):
    value = clean_value(value).lower()
    value = unicodedata.normalize("NFKC", value)
    value = re.sub(r"[^a-zа-я0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def find_column(columns, aliases):
    normalized = {
        column: normalize_header_name(column)
        for column in columns
    }
    normalized_aliases = [
        normalize_header_name(alias)
        for alias in aliases
    ]
    for column, header in normalized.items():
        if header in normalized_aliases:
            return column
    for column, header in normalized.items():
        if any(alias in header for alias in normalized_aliases):
            return column
    return None


def read_excel_all_sheets(file_or_buffer, engine="openpyxl"):
    result = []
    try:
        workbook = pd.ExcelFile(file_or_buffer, engine=engine)
        for sheet_name in workbook.sheet_names:
            try:
                df = pd.read_excel(
                    workbook,
                    sheet_name=sheet_name,
                    dtype=str
                )
                if not df.empty:
                    result.append((sheet_name, df))
            except Exception:
                continue
    except Exception:
        return []
    return result


@st.cache_data(show_spinner=False)
def load_vendor_mapping():
    diagnostics = []
    candidates = []

    for file_name in os.listdir("."):
        lower_name = file_name.lower()
        if lower_name.endswith((".xlsx", ".xls")) and (
            "cros" in lower_name
            or "vendor" in lower_name
            or "достав" in lower_name
        ):
            candidates.append(file_name)

    vendor_rows = []

    for file_name in candidates:
        engine = "xlrd" if file_name.lower().endswith(".xls") else "openpyxl"
        for sheet_name, df in read_excel_all_sheets(file_name, engine=engine):
            vendor_no_col = find_column(
                df.columns,
                [
                    "Buy-from Vendor No.",
                    "Buy from Vendor No",
                    "Vendor No.",
                    "Vendor No",
                    "Cross-Reference Type No."
                ]
            )
            vendor_name_col = find_column(
                df.columns,
                [
                    "Buy-from Vendor Name",
                    "Buy from Vendor Name",
                    "Vendor Name",
                    "Supplier Name"
                ]
            )

            if vendor_no_col is None or vendor_name_col is None:
                continue

            for _, row in df.iterrows():
                vendor_no = clean_value(row.get(vendor_no_col, ""))
                vendor_name = clean_value(row.get(vendor_name_col, ""))
                if vendor_no:
                    vendor_rows.append({
                        "Vendor No.": vendor_no,
                        "Vendor Name": vendor_name,
                        "Source": f"{file_name} / {sheet_name}"
                    })

    if not vendor_rows:
        diagnostics.append(
            "Не е намерен Excel файл с колони "
            "'Buy-from Vendor No.' и 'Buy-from Vendor Name'."
        )
        return pd.DataFrame(), diagnostics

    vendors = pd.DataFrame(vendor_rows)
    vendors = vendors.drop_duplicates(subset=["Vendor No."])
    vendors = vendors.sort_values(["Vendor Name", "Vendor No."])
    vendors = vendors.reset_index(drop=True)
    return vendors, diagnostics


@st.cache_data(show_spinner=False)
def load_vendor_cross_references(vendor_no):
    diagnostics = []
    records = []
    wanted_vendor = normalize_key(vendor_no)

    type_aliases = [
        "Cross-Reference Type No.",
        "Cross Reference Type No",
        "Cross-Reference Type Number",
        "Cross Reference Type Number"
    ]
    cross_ref_aliases = [
        "Cross-Reference No.",
        "Cross Reference No",
        "Cross-Reference Number",
        "Cross Reference Number"
    ]
    item_aliases = [
        "Item No.",
        "Item No",
        "Item Number",
        "No.",
        "Article No.",
        "Article Number",
        "Product No."
    ]

    for zip_path in CROSS_REFERENCE_ZIPS:
        if not os.path.exists(zip_path):
            diagnostics.append(f"Липсва архив: {zip_path}")
            continue

        try:
            with zipfile.ZipFile(zip_path, "r") as archive:
                for member_name in archive.namelist():
                    if member_name.endswith("/"):
                        continue

                    extension = os.path.splitext(member_name)[1].lower()
                    if extension not in {".xlsx", ".xls", ".csv"}:
                        continue

                    try:
                        raw_bytes = archive.read(member_name)
                        buffer = io.BytesIO(raw_bytes)

                        if extension == ".csv":
                            try:
                                dataframes = [("CSV", pd.read_csv(
                                    buffer,
                                    dtype=str,
                                    sep=None,
                                    engine="python",
                                    encoding="utf-8-sig"
                                ))]
                            except UnicodeDecodeError:
                                buffer.seek(0)
                                dataframes = [("CSV", pd.read_csv(
                                    buffer,
                                    dtype=str,
                                    sep=None,
                                    engine="python",
                                    encoding="cp1251"
                                ))]
                        else:
                            engine = "xlrd" if extension == ".xls" else "openpyxl"
                            dataframes = read_excel_all_sheets(buffer, engine=engine)

                        for sheet_name, df in dataframes:
                            if df.empty:
                                continue

                            type_col = find_column(df.columns, type_aliases)
                            cross_ref_col = find_column(df.columns, cross_ref_aliases)
                            item_col = find_column(df.columns, item_aliases)

                            if type_col is None or cross_ref_col is None:
                                continue

                            for _, row in df.iterrows():
                                type_no = clean_value(row.get(type_col, ""))
                                if normalize_key(type_no) != wanted_vendor:
                                    continue

                                cross_ref_no = clean_value(row.get(cross_ref_col, ""))
                                item_no = clean_value(row.get(item_col, "")) if item_col else ""

                                if not cross_ref_no:
                                    continue

                                records.append({
                                    "Cross-Reference Type No.": type_no,
                                    "Cross-Reference No.": cross_ref_no,
                                    "Item No.": item_no,
                                    "Source": f"{zip_path} / {member_name} / {sheet_name}"
                                })

                    except Exception as error:
                        diagnostics.append(
                            f"Неуспешно четене: {zip_path} / {member_name}: {error}"
                        )

        except Exception as error:
            diagnostics.append(f"Грешка в {zip_path}: {error}")

    if not records:
        diagnostics.append(
            f"В трите ZIP архива няма редове за vendor {vendor_no}."
        )
        return pd.DataFrame(), diagnostics

    result = pd.DataFrame(records)
    result = result.drop_duplicates(
        subset=[
            "Cross-Reference Type No.",
            "Cross-Reference No.",
            "Item No."
        ]
    )
    result = result.reset_index(drop=True)
    return result, diagnostics


def build_cross_reference_indexes(reference_df):
    direct_index = {}
    alias_index = {}

    for _, row in reference_df.iterrows():
        vendor_no = clean_value(row.get("Cross-Reference Type No.", ""))
        cross_ref_no = clean_value(row.get("Cross-Reference No.", ""))
        item_no = clean_value(row.get("Item No.", ""))

        cross_key = normalize_key(cross_ref_no)
        item_key = normalize_key(item_no)

        if cross_key:
            direct_index[cross_key] = {
                "vendor_no": vendor_no,
                "cross_ref_no": cross_ref_no
            }

        if item_key:
            alias_index[item_key] = {
                "vendor_no": vendor_no,
                "cross_ref_no": cross_ref_no
            }

    return direct_index, alias_index


def parse_invoice_number(value):
    text = clean_value(value)
    if not text:
        return None

    text = re.sub(r"[^\d,.\-]", "", text)
    if not text or text in {"-", ".", ","}:
        return None

    try:
        if "," in text and "." in text:
            if text.rfind(",") > text.rfind("."):
                text = text.replace(".", "").replace(",", ".")
            else:
                text = text.replace(",", "")
        elif "," in text:
            parts = text.split(",")
            if len(parts) == 2 and len(parts[1]) <= 6:
                text = text.replace(",", ".")
            else:
                text = text.replace(",", "")
        elif text.count(".") > 1:
            text = text.replace(".", "")

        return float(text)
    except Exception:
        return None


def find_reference_in_row(row, direct_index, alias_index):
    known = []

    for key, data in direct_index.items():
        known.append((key, data, "direct"))

    for key, data in alias_index.items():
        known.append((key, data, "alias"))

    known.sort(key=lambda entry: len(entry[0]), reverse=True)

    for column_index, cell in enumerate(row):
        original = clean_value(cell)
        cell_key = normalize_key(original)
        if not cell_key:
            continue

        for known_key, data, status in known:
            if cell_key == known_key or (
                len(known_key) >= 5 and known_key in cell_key
            ):
                return {
                    "vendor_no": data["vendor_no"],
                    "cross_ref_no": data["cross_ref_no"],
                    "status": status,
                    "source_value": original,
                    "column_index": column_index
                }

    return None


def find_unknown_item(row, vendor_no):
    blocked = {
        "TOTAL", "SUBTOTAL", "GESAMT", "PRICE", "PREIS",
        "QUANTITY", "MENGE", "INVOICE", "RECHNUNG"
    }
    candidates = []

    for column_index, cell in enumerate(row):
        original = clean_value(cell)
        key = normalize_key(original)
        if len(key) < 5 or len(key) > 40:
            continue
        if any(word in original.upper() for word in blocked):
            continue
        if re.fullmatch(r"\d+[,.]\d{1,6}", original):
            continue
        if re.fullmatch(r"\d{1,2}[./-]\d{1,2}[./-]\d{2,4}", original):
            continue
        if key.isdigit() and len(key) < 6:
            continue
        candidates.append((len(key), column_index, original))

    if not candidates:
        return None

    candidates.sort(reverse=True)
    _, column_index, original = candidates[0]
    return {
        "vendor_no": vendor_no,
        "cross_ref_no": f"❗ {original}",
        "status": "not_found",
        "source_value": original,
        "column_index": column_index
    }


def detect_qty_price_total(row, item_column_index):
    numeric = []

    for index, cell in enumerate(row):
        if index == item_column_index:
            continue
        number = parse_invoice_number(cell)
        if number is not None:
            numeric.append((index, number))

    best = None
    best_difference = None

    for qty_index, qty in numeric:
        if qty <= 0 or abs(qty - round(qty)) > 0.0001:
            continue
        for price_index, price in numeric:
            if price_index == qty_index or price < 0:
                continue
            for total_index, line_total in numeric:
                if total_index in {qty_index, price_index}:
                    continue
                difference = abs(qty * price - line_total)
                tolerance = max(0.05, abs(line_total) * 0.01)
                if difference <= tolerance and (
                    best_difference is None or difference < best_difference
                ):
                    best = (qty, price, line_total)
                    best_difference = difference

    if best:
        return best

    integer_values = [
        (index, value)
        for index, value in numeric
        if value > 0 and abs(value - round(value)) < 0.0001
    ]
    if not integer_values:
        return None, None, None

    qty_index, qty = integer_values[0]
    prices_after = [
        (index, value)
        for index, value in numeric
        if index > qty_index and value >= 0
    ]
    if not prices_after:
        return None, None, None

    _, price = prices_after[0]
    return qty, price, None


def convert_pdf_to_excel(pdf_files, vendor_no):
    reference_df, diagnostics = load_vendor_cross_references(vendor_no)

    if reference_df.empty:
        return {
            "output": None,
            "preview": pd.DataFrame(),
            "diagnostics": diagnostics,
            "pages": 0,
            "tables": 0,
            "direct": 0,
            "alias": 0,
            "not_found": 0
        }

    direct_index, alias_index = build_cross_reference_indexes(reference_df)

    recognized_rows = []
    processed_pages = 0
    detected_tables = 0

    for pdf_file in pdf_files:
        try:
            pdf_file.seek(0)
            with pdfplumber.open(pdf_file) as pdf:
                for page_number, page in enumerate(pdf.pages, start=1):
                    processed_pages += 1

                    settings_list = [
                        {
                            "vertical_strategy": "lines",
                            "horizontal_strategy": "lines",
                            "intersection_tolerance": 8,
                            "snap_tolerance": 5,
                            "join_tolerance": 5
                        },
                        {
                            "vertical_strategy": "text",
                            "horizontal_strategy": "text",
                            "intersection_tolerance": 8,
                            "snap_tolerance": 5,
                            "join_tolerance": 5,
                            "min_words_vertical": 1,
                            "min_words_horizontal": 1
                        }
                    ]

                    tables = []
                    for settings in settings_list:
                        try:
                            tables = page.extract_tables(settings) or []
                        except Exception:
                            tables = []
                        if tables:
                            break

                    for table_number, table in enumerate(tables, start=1):
                        if not table:
                            continue
                        detected_tables += 1

                        for row_number, row in enumerate(table, start=1):
                            row = [clean_value(cell) for cell in row]
                            if not any(row):
                                continue

                            row_text = " ".join(row).upper()
                            if any(word in row_text for word in [
                                "GRAND TOTAL", "INVOICE TOTAL", "SUBTOTAL",
                                "ZWISCHENSUMME", "GESAMTBETRAG"
                            ]):
                                continue

                            reference = find_reference_in_row(
                                row,
                                direct_index,
                                alias_index
                            )

                            if reference is None:
                                reference = find_unknown_item(row, vendor_no)

                            if reference is None:
                                continue

                            qty, price, line_total = detect_qty_price_total(
                                row,
                                reference["column_index"]
                            )

                            if qty is None or price is None:
                                continue
                            if qty <= 0 or price < 0:
                                continue

                            display_cross_ref = reference["cross_ref_no"]
                            if reference["status"] == "alias":
                                display_cross_ref = f"⚠️ {display_cross_ref}"

                            recognized_rows.append({
                                "Cross-Reference Type No.": reference["vendor_no"],
                                "Cross-Reference No.": display_cross_ref,
                                "Qty": qty,
                                "Price 1 pc": price,
                                "_status": reference["status"],
                                "_source_value": reference["source_value"],
                                "_line_total": line_total,
                                "_file": pdf_file.name,
                                "_page": page_number,
                                "_table": table_number,
                                "_row": row_number
                            })

        except Exception as error:
            diagnostics.append(f"{pdf_file.name}: {error}")

    if not recognized_rows:
        return {
            "output": None,
            "preview": pd.DataFrame(),
            "diagnostics": diagnostics,
            "pages": processed_pages,
            "tables": detected_tables,
            "direct": 0,
            "alias": 0,
            "not_found": 0
        }

    full_df = pd.DataFrame(recognized_rows)
    full_df = full_df.drop_duplicates(
        subset=[
            "Cross-Reference Type No.",
            "Cross-Reference No.",
            "Qty",
            "Price 1 pc",
            "_file",
            "_page",
            "_row"
        ]
    ).reset_index(drop=True)

    preview_df = full_df[[
        "Cross-Reference Type No.",
        "Cross-Reference No.",
        "Qty",
        "Price 1 pc"
    ]].copy()

    grand_total = (
        pd.to_numeric(preview_df["Qty"], errors="coerce")
        * pd.to_numeric(preview_df["Price 1 pc"], errors="coerce")
    ).sum()

    total_row = pd.DataFrame([{
        "Cross-Reference Type No.": "",
        "Cross-Reference No.": "TOTAL",
        "Qty": "",
        "Price 1 pc": grand_total
    }])

    export_df = pd.concat([preview_df, total_row], ignore_index=True)
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="Invoice", index=False)
        ws = writer.sheets["Invoice"]

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:D{len(export_df) + 1}"
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 18

        header_fill = PatternFill(fill_type="solid", fgColor="D71919")
        header_font = Font(color="FFFFFF", bold=True)
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9")
        )

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        last_product_row = len(preview_df) + 1
        total_excel_row = len(export_df) + 1

        for excel_row in range(2, last_product_row + 1):
            cross_cell = ws.cell(row=excel_row, column=2)
            cross_value = str(cross_cell.value or "")

            if cross_value.startswith("⚠️"):
                cross_cell.font = Font(bold=True, color="FF8C00")
            elif cross_value.startswith("❗"):
                cross_cell.font = Font(bold=True, color="FF0000")

            ws.cell(row=excel_row, column=3).number_format = "0.###"
            ws.cell(row=excel_row, column=4).number_format = "0.000000"

            for column_number in range(1, 5):
                ws.cell(row=excel_row, column=column_number).border = thin_border

        total_fill = PatternFill(fill_type="solid", fgColor="FFF2CC")
        for column_number in range(1, 5):
            cell = ws.cell(row=total_excel_row, column=column_number)
            cell.fill = total_fill
            cell.font = Font(bold=True, color="C00000")
            cell.border = thin_border

        ws.cell(row=total_excel_row, column=4).number_format = "0.00"

    output.seek(0)

    return {
        "output": output,
        "preview": preview_df,
        "diagnostics": diagnostics,
        "pages": processed_pages,
        "tables": detected_tables,
        "direct": int((full_df["_status"] == "direct").sum()),
        "alias": int((full_df["_status"] == "alias").sum()),
        "not_found": int((full_df["_status"] == "not_found").sum())
    }

# ======================================================
# ✅ PRN FOR NAVISION
# ======================================================
if menu == "PRN за Navision" and uploaded_files:

    file = uploaded_files[0]

    try:

        df = pd.read_excel(file)

        # Премахване на случайни интервали в имената на колоните
        df.columns = [str(col).strip() for col in df.columns]

        # Приема колоната за цена като "Price 1 pc" или "Price"
        # Ако съществува "Price 1 pc", тя остава с приоритет
        if "Price 1 pc" not in df.columns and "Price" in df.columns:
            df = df.rename(columns={"Price": "Price 1 pc"})

        # Excel от PDF конвертора използва Cross-Reference No.
        if "Item" not in df.columns and "Cross-Reference No." in df.columns:
            df = df.rename(columns={"Cross-Reference No.": "Item"})

        st.subheader("📄 PRN Preview")

        required_cols = [
            "Item",
            "Qty",
            "Price 1 pc"
        ]

        missing = [
            c for c in required_cols
            if c not in df.columns
        ]

        if missing:
            st.error(
                f"Липсват колони: {', '.join(missing)}. "
                "Колоната за цена може да бъде 'Price 1 pc' или 'Price'."
            )
            st.stop()

        preview = df[
            ["Item", "Qty", "Price 1 pc"]
        ]

        st.dataframe(
            preview,
            use_container_width=True
        )

        prn_lines = []

        for _, row in df.iterrows():

            item = str(
                row["Item"]
            ).strip()

            if item == "" or item.lower() == "nan":
                continue

            qty = int(
                round(float(str(row["Qty"]).replace(",", ".")))
            )

            price = float(
                str(row["Price 1 pc"]).replace(",", ".")
            )

            price_str = (
                f"{price:.6f}"
                .replace(".", ",")
            )

            line = (
                item +
                (" " * (25 - len(item) - len(str(qty)))) +
                str(qty) +
                (" " * 6) +
                price_str
            )

            prn_lines.append(line)

        prn_content = "\r\n".join(prn_lines)

        # Взима Invoice от Excel, а при липса използва името на файла
        if "Invoice" in df.columns:
            invoice_no = str(df.iloc[0]["Invoice"]).strip()
        else:
            invoice_no = os.path.splitext(file.name)[0]

        st.download_button(
            f"⬇️ Изтегли {invoice_no}.prn",
            data=prn_content.encode("utf-8"),
            file_name=f"{invoice_no}.prn",
            mime="text/plain"
        )

        st.write(f"Редове в Excel: {len(df)}")
        st.write(f"Редове в PRN: {len(prn_lines)}")

        st.stop()

    except Exception as e:
        st.error(f"Грешка: {e}")
        st.stop()
        
# ======================================================
# ✅ PDF INVOICE TO EXCEL
# ======================================================
if menu == "PDF към Excel":

    st.markdown(
        """
        <style>
        .pdf-excel-box {
            padding: 16px;
            margin-top: 10px;
            margin-bottom: 15px;
            border-radius: 14px;
            border: 1px solid rgba(255,255,255,0.20);
            background: rgba(0,0,0,0.25);
            color: white;
        }
        .pdf-excel-title {
            font-size: 22px;
            font-weight: 800;
            color: #ff8c00;
            margin-bottom: 6px;
        }
        .pdf-excel-description {
            font-size: 14px;
            color: rgba(255,255,255,0.80);
            line-height: 1.6;
        }
        </style>
        <div class="pdf-excel-box">
            <div class="pdf-excel-title">📄 PDF фактура към Excel</div>
            <div class="pdf-excel-description">
                Избира vendor от справочния Excel, филтрира трите ZIP архива
                по Cross-Reference Type No. и извежда четири контролни колони.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    vendors_df, vendor_diagnostics = load_vendor_mapping()

    if vendors_df.empty:
        st.error(
            "❌ Не е намерен vendor справочник с колони "
            "Buy-from Vendor No. и Buy-from Vendor Name."
        )
        if vendor_diagnostics:
            with st.expander("🔎 Диагностика"):
                for message in vendor_diagnostics:
                    st.write(f"• {message}")
        st.stop()

    vendor_options = vendors_df.apply(
        lambda row: (
            f"{row['Vendor No.']} | {row['Vendor Name']}"
            if row['Vendor Name']
            else row['Vendor No.']
        ),
        axis=1
    ).tolist()

    selected_vendor_label = st.selectbox(
        "Избери доставчик / Vendor",
        vendor_options,
        key="pdf_vendor_selection"
    )

    selected_vendor_no = selected_vendor_label.split(" | ", 1)[0].strip()

    reference_preview, reference_diagnostics = load_vendor_cross_references(
        selected_vendor_no
    )

    if reference_preview.empty:
        st.warning(
            f"⚠️ Не са намерени Cross Reference записи за {selected_vendor_no}."
        )
    else:
        st.caption(
            f"Намерени Cross Reference записи за vendor: {len(reference_preview)}"
        )

    if uploaded_files:
        with st.spinner("Разпознаване чрез Vendor Cross Reference..."):
            conversion_result = convert_pdf_to_excel(
                uploaded_files,
                selected_vendor_no
            )

        if conversion_result["output"] is None:
            st.error(
                "❌ Не бяха разпознати редове с Cross Reference, "
                "количество и цена."
            )
            with st.expander("🔎 Диагностика"):
                for message in conversion_result["diagnostics"]:
                    st.write(f"• {message}")
            st.stop()

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("✅ Директни", conversion_result["direct"])
        with col2:
            st.metric("⚠️ Чрез Item No.", conversion_result["alias"])
        with col3:
            st.metric("❗ За проверка", conversion_result["not_found"])

        st.success(
            f"Страници: {conversion_result['pages']} | "
            f"Таблици: {conversion_result['tables']}"
        )

        st.subheader("📋 Контролен преглед")
        st.dataframe(
            conversion_result["preview"],
            use_container_width=True
        )

        st.caption(
            "⚠️ = Cross Reference е намерен чрез Item No. | "
            "❗ = номерът не е намерен за избрания vendor"
        )

        if len(uploaded_files) == 1:
            excel_file_name = re.sub(
                r"\.pdf$",
                ".xlsx",
                uploaded_files[0].name,
                flags=re.IGNORECASE
            )
        else:
            excel_file_name = "pdf_invoices_converted.xlsx"

        st.download_button(
            label="📥 Изтегли Excel за PRN",
            data=conversion_result["output"],
            file_name=excel_file_name,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

    else:
        st.info("⬆️ Качи PDF фактура.")

    st.stop()

# ======================================================
# ✅ PROCESS
# ======================================================
if uploaded_files:

    all_data = []

    for file in uploaded_files:

        df = None

        # ✅ CASTROL + NESTE
        if menu == "CASTROL & NESTE" and source_type == "Excel":
            df = parse_castrol_excel(file)

        # ✅ FLUKAR
        elif menu == "FLUKAR":
            df = parse_flukar_excel(file)

        # ✅ ELROMI
        elif menu == "ELROMI RONAX":
            df = parse_elromi_excel(file)

        # ✅ NISTA
        elif menu == "NISTA":
            df = parse_nista_excel(file)

        # ✅ ORLEN
        elif menu == "ORLEN":
            df = parse_orlen_excel(file)

        # ✅ AUTO MEGA
        elif menu == "AUTO MEGA":
            df = parse_auto_mega_excel(file)

        # ✅ AMTRA
        elif menu == "AMTRA":
            df = parse_amtra_excel(file)

        # ✅ EMINIA
        elif menu == "EMINIA":
            df = parse_eminia_excel(file)

        # ✅ BRECHMANN
        elif menu == "Brehman":
            df = parse_brechmann_excel(file)

        # ✅ FEBI
        elif menu == "FEBI":
            df = parse_febi_excel(file)

        # ✅ CHEMPIOIL EXCEL
        elif menu == "Chempioil (FANFARO)" and source_type == "Excel":
            df = parse_chempioil_excel(file)

                # ✅ VALVOLINE
        elif menu == "VALVOLINE":
            df = parse_valvoline_excel(file)

        # ✅ FUCHS EXCEL
        elif menu == "FUCHS" and source_type == "Excel":
            df = parse_fuchs_excel(file)

        # ✅ PDF SECTION
        elif source_type == "PDF":

            reader = PdfReader(file)

            text = ""

            for page in reader.pages:

                page_text = page.extract_text()

                if page_text:
                    text += page_text + "\n"

            # ✅ FUCHS PDF
            if menu == "FUCHS":
                df = parse_fuchs(text)

            # ✅ CHEMPIOIL PDF
            elif menu == "Chempioil (FANFARO)":
                df = parse_chempioil_pdf(text)

            # ✅ GASOLINE PDF
            elif menu == "GASOLINE":
                df = parse_gasoline(text)

            # ✅ MOTUL + ОСТАНАЛИТЕ PDF
            else:
                df = parse_motul(text)

            reader = PdfReader(file)

            text = ""

            for page in reader.pages:

                t = page.extract_text()

                if t:
                    text += t + "\n"

            # ✅ FUCHS
            if menu == "FUCHS":

                df = parse_fuchs(text)

            # ✅ CHEMPIOIL PDF
            elif menu == "Chempioil (FANFARO)":

                df = parse_chempioil_pdf(text)

            # ✅ GASOLINE
            elif menu == "GASOLINE":

                df = parse_gasoline(text)

            # ✅ MOTUL + останалите PDF
            else:

                df = parse_motul(text)

        else:

            df = pd.read_excel(file)
            df.columns = df.columns.str.strip()

        if isinstance(df, pd.DataFrame) and not df.empty:
            all_data.append(df)

    if not all_data:
        st.warning("⚠️ Няма данни")
        st.stop()

    final_df = pd.concat(
        all_data,
        ignore_index=True
    )

    if "Тарифен код" not in final_df.columns:
        st.warning("⚠️ Данните не съдържат тарифен код")
        st.stop()

    final_df["Тарифен код"] = (
        final_df["Тарифен код"]
        .astype(str)
    )

    final_df = final_df[
        final_df["тегло"] > 0
    ]

    report = build_final_report(
        final_df,
        menu
    )

    report["Тарифен код"] = (
        report["Тарифен код"]
        .astype(str)
    )

    report["Тарифен код"] = report["Тарифен код"].str.replace(
        "38119000",
        "38119000 - EMCS",
        regex=False
    )

    special_codes = [
        "38112100",
        "38249992",
        "27101225",
        "38140090",
        "38249996"
    ]

    for code in special_codes:

        report["Тарифен код"] = report["Тарифен код"].str.replace(
            f"{code} -",
            f"{code} - ( ! )",
            regex=False
        )

        report["Тарифен код"] = report["Тарифен код"].replace(
            {code: f"{code} ( ! )"}
        )

    # ✅ ⬆️5 кодове
    arrow_codes = [
        "34031910",
        "34039900",
        "34031980",
        "27101999"
    ]

    for code in arrow_codes:

        report["Тарифен код"] = report["Тарифен код"].str.replace(
            f"{code} -",
            f"{code} ⬆️5 -",
            regex=False
        )

        report["Тарифен код"] = report["Тарифен код"].replace(
            {code: f"{code} ⬆️5"}
        )

    st.subheader("📊 Финален отчет")
    st.dataframe(report)

    report = report.rename(columns={
        "Тарифен код": "Code",
        "wid": "wid",
        "Количество": "Broj",
        "kolichestvo": "colic-v L",
        "тегло": "teglo"
    })

    from openpyxl.styles import Font

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        report.to_excel(
            writer,
            index=False
        )

        ws = writer.sheets["Sheet1"]

        red_codes = [
    "34039900",
    "34031910",
    "34031980",

    # ✅ кодове с ( ! )
    "38112100",
    "38249992",
    "27101225",
    "38140090",
    "38249996"
]

        for row in ws.iter_rows(min_row=2):

            code_cell = row[0]

            value = str(code_cell.value)

            for code in red_codes:

                if value.startswith(code):

                    code_cell.font = Font(
                        bold=True,
                        color="FF0000"
                    )

                    break

    output.seek(0)

    st.download_button(
        label="📥 Изтегли Excel",
        data=output,
        file_name="final_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:

    st.markdown(
        "**⬆️ Upload a file to generate a report**"
    )
