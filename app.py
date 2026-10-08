import streamlit as st
import pandas as pd
import psycopg2
import base64
import io
import re
import zipfile
import pdfplumber
from psycopg2.extras import RealDictCursor
from openpyxl.styles import (
    Font,
    PatternFill,
    Border,
    Side,
    Alignment
)


# ======================================================
# APP CONFIG
# ======================================================

st.set_page_config(
    page_title="PDF → PRN Converter",
    page_icon="📄",
    layout="wide"
)


# ======================================================
# DATABASE
# ======================================================

DATABASE_URL = st.secrets["DATABASE_URL"]


# ======================================================
# BACKGROUND
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
                background-image:
                    url("data:image/png;base64,{encoded}");
                background-size: cover;
                background-position: center;
                background-repeat: no-repeat;
            }}
            </style>
            """,
            unsafe_allow_html=True
        )

    except Exception:
        pass


# ======================================================
# LOGIN
# ======================================================

def check_login():

    if "logged_in" not in st.session_state:

        st.session_state["logged_in"] = False

    if st.session_state["logged_in"]:

        return True

    set_bg("background_login.png")

    st.markdown(
        """
        <div style="
            text-align:center;
            margin-top:40px;
        ">
        <h1 style="
            color:white;
            font-size:42px;
            font-weight:900;
        ">
        📄 PDF → PRN Converter
        </h1>

        <p style="
            color:white;
            font-size:18px;
        ">
        Inter Cars Logistics Platform
        </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    username = st.text_input(
        "User"
    )

    password = st.text_input(
        "Password",
        type="password"
    )

    if st.button(
        "Login",
        use_container_width=True
    ):

        if (
            username == "intercars"
            and
            password == "Intercars2026"
        ):

            st.session_state["logged_in"] = True

            st.rerun()

        else:

            st.error(
                "Грешно име или парола"
            )

    return False


if not check_login():

    st.stop()


# ======================================================
# MAIN BACKGROUND
# ======================================================

set_bg("background.png")


# ======================================================
# STYLES
# ======================================================

st.markdown(
    """
    <style>

    section[data-testid="stSidebar"] > div {
        background: rgba(0,0,0,0.45);
        backdrop-filter: blur(12px);
    }

    .main-card{
        background:rgba(0,0,0,0.35);
        border-radius:15px;
        padding:18px;
        border:1px solid rgba(255,255,255,0.15);
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ======================================================
# NEON CONNECTION
# ======================================================

@st.cache_resource
def get_connection():

    return psycopg2.connect(
        DATABASE_URL,
        keepalives=1,
        keepalives_idle=30,
        keepalives_interval=10,
        keepalives_count=5
    )


# ======================================================
# LOAD VENDORS
# ======================================================

@st.cache_data(ttl=70)
def load_vendors():

    conn = psycopg2.connect(
        DATABASE_URL
    )

    query = """
    SELECT
        vendor_no,
        vendor_name
    FROM vendors
    WHERE vendor_no IS NOT NULL
    AND vendor_name IS NOT NULL
    ORDER BY vendor_name
    """

    df = pd.read_sql_query(
        query,
        conn
    )

    conn.close()

    df["vendor_no"] = (
        df["vendor_no"]
        .astype(str)
        .str.strip()
    )

    df["vendor_name"] = (
        df["vendor_name"]
        .astype(str)
        .str.strip()
    )

    df = df[
        df["vendor_no"]
        .str.upper()
        != "VENDOR_NO"
    ]

    return df.reset_index(drop=True)    

# ======================================================
# LOAD DATA
# ======================================================

vendors_df = load_vendors()


# ======================================================
# SESSION STATE - DIRECT PDF TO PRN
# ======================================================

if "converter_page" not in st.session_state:
    st.session_state.converter_page = "📄 PDF → Excel"

if "prn_ready_df" not in st.session_state:
    st.session_state.prn_ready_df = pd.DataFrame(
        columns=[
            "Item",
            "Qty",
            "Price 1 pc"
        ]
    )

if "prn_invoice_name" not in st.session_state:
    st.session_state.prn_invoice_name = "invoice"

if "prn_source" not in st.session_state:
    st.session_state.prn_source = ""
# ======================================================
# SESSION STATE - CONTAINER PRN
# ======================================================

if "container_prn_data" not in st.session_state:
    st.session_state.container_prn_data = {}

if "container_prn_invoice_name" not in st.session_state:
    st.session_state.container_prn_invoice_name = "invoice"

if "container_prn_loaded" not in st.session_state:
    st.session_state.container_prn_loaded = False
# ======================================================
# SIDEBAR
# ======================================================

st.sidebar.title(
    "PDF → PRN Converter"
)

# ======================================================
# PAGE NAVIGATION
# ======================================================

PAGE_PDF = "📄 PDF → Excel"
PAGE_PRN = "🧾 Excel → PRN"

if "converter_page" not in st.session_state:
    st.session_state.converter_page = PAGE_PDF


def open_pdf_page():
    st.session_state.converter_page = PAGE_PDF


def open_prn_page():
    st.session_state.converter_page = PAGE_PRN


page = st.sidebar.radio(
    "Menu",
    [
        PAGE_PDF,
        PAGE_PRN
    ],
    index=(
        0
        if st.session_state.converter_page == PAGE_PDF
        else 1
    ),
    key="converter_menu",
    on_change=lambda: st.session_state.update(
        converter_page=st.session_state.converter_menu
    )
)

# ======================================================
# LOGOUT
# ======================================================

if st.sidebar.button(
    "🚪 Logout"
):

    st.session_state["logged_in"] = False

    st.rerun()


# ======================================================
# VENDOR DROPDOWN
# ======================================================

if page == "📄 PDF → Excel":

    vendors_df["display"] = (
        vendors_df["vendor_no"].astype(str)
        +
        " | "
        +
        vendors_df["vendor_name"].astype(str)
    )

    selected_vendor = st.selectbox(
        "🚚 Избери доставчик",
        vendors_df["display"],
        index=0,
        key="vendor_selector"
    )

    selected_vendor_no = (
        selected_vendor
        .split(" | ")[0]
        .strip()
    )

    st.markdown(
        f"""
        <div style="
            background:rgba(0,0,0,0.35);
            padding:14px;
            border-radius:12px;
            border:1px solid rgba(255,255,255,0.15);
            color:#00ff88;
            font-size:18px;
            font-weight:700;
        ">
            ✅ Избран Vendor:
            {selected_vendor_no}
        </div>
        """,
        unsafe_allow_html=True
    )

else:

    selected_vendor_no = None
# ======================================================
# LOAD CROSS REFERENCES FROM BOTH NEON PROJECTS
# ======================================================

@st.cache_data(ttl=3600)
def load_cross_references(vendor_no):

    databases = [
        st.secrets["DATABASE_URL"],
        st.secrets["DATABASE_URL_2"]
    ]

    all_frames = []

    query = """
    SELECT
        vendor_no,
        cross_reference_no,
        item_no,
        normalized_cross_reference,
        normalized_item_no
    FROM cross_references
    WHERE vendor_no = %s
    """

    for database_url in databases:

        conn = None

        try:

            conn = psycopg2.connect(
                database_url
            )

            df = pd.read_sql_query(
                query,
                conn,
                params=[vendor_no]
            )

            if not df.empty:

                all_frames.append(df)

        except Exception as error:

            st.warning(
                f"Neon connection problem: {error}"
            )

        finally:

            if conn is not None:

                conn.close()

    if not all_frames:

        return pd.DataFrame(
            columns=[
                "vendor_no",
                "cross_reference_no",
                "item_no",
                "normalized_cross_reference",
                "normalized_item_no"
            ]
        )

    result = pd.concat(
        all_frames,
        ignore_index=True
    )

    result = (
        result
        .drop_duplicates(
            subset=[
                "vendor_no",
                "cross_reference_no",
                "item_no"
            ]
        )
        .reset_index(drop=True)
    )

    return result

# ======================================================
# ITEM NORMALIZATION
# ======================================================

def normalize_item_number(value):

    if value is None:
        return ""

    text = str(value).strip().upper()

    if text.lower() in {
        "",
        "nan",
        "none",
        "null"
    }:
        return ""

    return re.sub(
        r"[^A-Z0-9]",
        "",
        text
    )


# ======================================================
# EUROPEAN NUMBER PARSER
# ======================================================

def parse_european_number(value):

    if value is None:
        return None

    text = str(value).strip()

    if not text:
        return None

    text = re.sub(
        r"[^\d,.\-]",
        "",
        text
    )

    if not text:
        return None

    try:

        # Пример: 2.121,60 -> 2121.60
        if "," in text and "." in text:

            if text.rfind(",") > text.rfind("."):

                text = (
                    text
                    .replace(".", "")
                    .replace(",", ".")
                )

            else:

                text = text.replace(",", "")

        # Пример: 20,25 -> 20.25
        elif "," in text:

            text = text.replace(",", ".")

        return float(text)

    except Exception:

        return None


# ======================================================
# EXTRACT INVOICE NUMBER
# ======================================================

def extract_invoice_number(
    complete_text,
    fallback_file_name
):

    patterns = [
        r"Invoice\s+Number\s+(\d+)",
        r"Invoice\s+No\.?\s*(\d+)",
        r"Invoice\s*#\s*(\d+)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            complete_text,
            re.IGNORECASE
        )

        if match:

            return match.group(1)

    return re.sub(
        r"\.pdf$",
        "",
        fallback_file_name,
        flags=re.IGNORECASE
    )


# ======================================================
# EXTRACT REAL INVOICE PRODUCT ROWS
# ======================================================

def extract_federal_rows(pdf_file):

    extracted_rows = []
    complete_text = ""
    processed_pages = 0

    # Примери:
    # PE-BJ-3322 18 EA 20,25 364,50 20,25 364,50
    # AU-ES-3721 54 EA 4,68 252,72 4,68 252,72

    row_pattern = re.compile(
        r"^\s*"
        r"([A-Z0-9][A-Z0-9\-./]+)"
        r"\s+"
        r"(\d+(?:[.,]\d+)?)"
        r"\s+EA"
        r"\s+"
        r"([\d.,]+)"
        r"\s+"
        r"([\d.,]+)"
        r"(?:"
        r"\s+"
        r"([\d.,]+)"
        r"\s+"
        r"([\d.,]+)"
        r")?",
        re.IGNORECASE
    )

    pdf_file.seek(0)

    with pdfplumber.open(pdf_file) as pdf:

        for page_number, page in enumerate(
            pdf.pages,
            start=1
        ):

            processed_pages += 1

            page_text = page.extract_text(
                x_tolerance=2,
                y_tolerance=3
            )

            if not page_text:
                continue

            complete_text += page_text + "\n"

            for line in page_text.splitlines():

                clean_line = " ".join(
                    str(line).split()
                )

                match = row_pattern.match(
                    clean_line
                )

                if not match:
                    continue

                invoice_item = (
                    match.group(1).strip()
                )

                quantity = parse_european_number(
                    match.group(2)
                )

                first_unit_price = (
                    parse_european_number(
                        match.group(3)
                    )
                )

                first_line_total = (
                    parse_european_number(
                        match.group(4)
                    )
                )

                second_unit_price = (
                    parse_european_number(
                        match.group(5)
                    )
                    if match.group(5)
                    else None
                )

                second_line_total = (
                    parse_european_number(
                        match.group(6)
                    )
                    if match.group(6)
                    else None
                )

                # Ако има повторена Net Unit Price,
                # използваме втората цена.
                unit_price = (
                    second_unit_price
                    if second_unit_price is not None
                    else first_unit_price
                )

                line_total = (
                    second_line_total
                    if second_line_total is not None
                    else first_line_total
                )

                if quantity is None:
                    continue

                if unit_price is None:
                    continue

                if quantity <= 0:
                    continue

                if unit_price < 0:
                    continue

                calculation_ok = True

                if line_total is not None:

                    expected_total = (
                        quantity * unit_price
                    )

                    tolerance = max(
                        0.05,
                        abs(line_total) * 0.01
                    )

                    calculation_ok = (
                        abs(
                            expected_total
                            - line_total
                        )
                        <= tolerance
                    )

                extracted_rows.append({
                    "invoice_item":
                        invoice_item,

                    "normalized_invoice_item":
                        normalize_item_number(
                            invoice_item
                        ),

                    "qty":
                        quantity,

                    "price":
                        unit_price,

                    "line_total":
                        line_total,

                    "calculation_ok":
                        calculation_ok,

                    "page":
                        page_number,

                    "source_line":
                        clean_line
                })

    invoice_number = extract_invoice_number(
        complete_text,
        pdf_file.name
    )

    return {
        "invoice_number":
            invoice_number,

        "pages":
            processed_pages,

        "rows":
            extracted_rows
    }


# ======================================================
# CASTROL PARSER
# ======================================================

def extract_castrol_rows(pdf_file):

    rows = []
    full_text = ""

    pdf_file.seek(0)

    with pdfplumber.open(pdf_file) as pdf:

        pages_count = len(pdf.pages)

        for page_number, page in enumerate(
            pdf.pages,
            start=1
        ):

            page_text = page.extract_text()

            if not page_text:
                continue

            full_text += page_text + "\n"

            lines = [
                line.strip()
                for line in page_text.splitlines()
                if line.strip()
            ]

            for index in range(1, len(lines)):

                current_line = lines[index]

                # Cross Ref код
                if not re.match(
                    r'^[A-Z0-9]{5,10}$',
                    current_line
                ):
                    continue

                invoice_item = current_line

                previous_line = lines[
                    index - 1
                ]

                if "ST" not in previous_line:
                    continue

                st_part = (
                    previous_line
                    .split("ST")[-1]
                )

                numbers = re.findall(
                    r'\d[\d\.,]*',
                    st_part
                )

                if len(numbers) < 4:
                    continue

                try:

                    qty = parse_european_number(
                        numbers[0]
                    )

                    price = parse_european_number(
                        numbers[1]
                    )

                    line_total = parse_european_number(
                        numbers[2]
                    )

                except Exception:
                    continue

                if qty is None:
                    continue

                if price is None:
                    continue

                rows.append({

                    "invoice_item":
                        invoice_item,

                    "normalized_invoice_item":
                        normalize_item_number(
                            invoice_item
                        ),

                    "qty":
                        qty,

                    "price":
                        price,

                    "line_total":
                        line_total,

                    "calculation_ok":
                        True,

                    "page":
                        page_number,

                    "source_line":
                        previous_line

                })

    invoice_number = (
        extract_invoice_number(
            full_text,
            pdf_file.name
        )
    )

    return {

        "invoice_number":
            invoice_number,

        "pages":
            pages_count,

        "rows":
            rows

    }
# ======================================================
# MOTUL PACK SIZE
# ======================================================

def extract_motul_pack_quantity(description):
    """
    Извлича броя артикули в един кашон от описанието.

    Примери:
    8100 ECO-LITE 0W20 4X5L       -> 4
    8100 ECO-CLEAN 0W20 12X1L     -> 12
    C5 CHAIN PASTE 12X0.150L       -> 12
    HIGH-TORQUE DCTF 12X1L D38     -> 12
    TEKMA ULTIMA 5W-30 LS 20L      -> 1
    """

    if description is None:
        return 1

    text = str(description).upper().strip()

    # Позволява:
    # 12X1L
    # 12 X 1L
    # 12X0.150L
    # 12X0,150L
    # 4X5L
    # 24X400ML
    pack_match = re.search(
        r"\b(\d{1,4})\s*[XХ×]\s*"
        r"\d+(?:[.,]\d+)?\s*"
        r"(?:ML|CL|DL|L|LT|LTR|KG|G)?\b",
        text,
        re.IGNORECASE
    )

    if not pack_match:
        return 1

    try:
        pack_quantity = int(pack_match.group(1))
    except Exception:
        return 1

    if pack_quantity <= 0:
        return 1

    return pack_quantity
# ======================================================
# MOTUL PARSER
# ======================================================

def extract_motul_rows(pdf_file):
    """
    MOTUL parser.

    Поддържа:
    - позиции, разделени между две PDF страници;
    - PCE количества;
    - CAR количества с разфасовка от Description;
    - MDEU EXP Add Disc отстъпки;
    - няколко позиции с един и същ Item-No.;
    - цена на брой след приспадане на отстъпката.
    """

    rows = []
    full_text = ""
    processed_pages = 0

    # Всички линии от всички страници се събират
    # в един общ списък. Това е важно, когато позиция
    # започва на една страница и завършва на следващата.
    document_lines = []

    pdf_file.seek(0)

    with pdfplumber.open(pdf_file) as pdf:
        processed_pages = len(pdf.pages)

        for page_number, page in enumerate(
            pdf.pages,
            start=1
        ):
            page_text = page.extract_text(
                x_tolerance=2,
                y_tolerance=3
            )

            if not page_text:
                continue

            full_text += page_text + "\n"

            for original_line in page_text.splitlines():
                clean_line = " ".join(
                    str(original_line).split()
                )

                if not clean_line:
                    continue

                document_lines.append({
                    "page": page_number,
                    "text": clean_line
                })

    # ==============================================
    # MOTUL PRODUCT START
    # ==============================================
    #
    # Примери:
    # 10 114190 TEKMA MEGA-X 10W-40 LS 20L
    # 100 114221 TEKMA MEGA-X 10W-40 20L
    # 134 104010 TRANSLUBE 20L
    #
    # search(), а не match(), защото в някои PDF
    # файлове преди Pos. може да има останал текст.
    item_start_pattern = re.compile(
        r"(?:^|\s)"
        r"(\d{1,4})"             # Pos.
        r"\s+"
        r"(\d{6})"               # Motul Item-No.
        r"\s+"
        r"(.+?)"                 # Description
        r"\s*$",
        re.IGNORECASE
    )

    # Основен ред с количество и Amount:
    #
    # 40 PCE 3.005,53
    # 7 CAR 372,69
    # 160 PCE 9.930,62
    quantity_amount_pattern = re.compile(
        r"\b"
        r"(\d+(?:[.,]\d+)?)"
        r"\s+"
        r"(PCE|CAR)"
        r"\s+"
        r"([\d.]+(?:,\d+)?)"
        r"\b",
        re.IGNORECASE
    )

    # Отстъпка:
    #
    # MDEU EXP Add Disc % 8,50- 37,65
    # MDEU EXP Add Disc % 12,00- 156,49
    discount_pattern = re.compile(
        r"MDEU\s+EXP\s+ADD\s+DISC\s*%"
        r".*?"
        r"(\d+(?:[.,]\d+)?)\s*[-−]?"
        r"\s+"
        r"([\d.]+(?:,\d+)?)",
        re.IGNORECASE
    )

    # По-свободен резервен шаблон.
    discount_fallback_pattern = re.compile(
        r"ADD\s+DISC"
        r".*?"
        r"([\d.]+(?:,\d+)?)"
        r"\s*$",
        re.IGNORECASE
    )

    item_starts = []

    # ==============================================
    # ОТКРИВАНЕ НА НАЧАЛОТО НА ВСЯКА ПОЗИЦИЯ
    # ==============================================

    for line_index, line_data in enumerate(
        document_lines
    ):
        line = line_data["text"]

        item_match = item_start_pattern.search(line)

        if not item_match:
            continue

        position_no = item_match.group(1).strip()
        item_no = item_match.group(2).strip()
        description = item_match.group(3).strip()

        description_upper = description.upper()

        # Защита срещу служебни редове, ако случайно
        # съвпаднат с числовия шаблон.
        forbidden_texts = (
            "INVOICE NO",
            "CUSTOMER",
            "SALES ORDER",
            "PURCH. ORDER",
            "PURCH ORDER",
            "DELIVERY NOTE",
            "BATCH-NO",
            "COUNTRY OF ORIGIN",
            "TOTAL GROSS",
            "TOTAL NET",
            "TOTAL PACKAGES",
            "TERMS OF PAYMENT"
        )

        if any(
            forbidden_text in description_upper
            for forbidden_text in forbidden_texts
        ):
            continue

        # Описанието трябва да съдържа поне една буква.
        if not re.search(
            r"[A-ZА-Я]",
            description_upper
        ):
            continue

        item_starts.append({
            "line_index": line_index,
            "page": line_data["page"],
            "position_no": position_no,
            "item_no": item_no,
            "description": description
        })

    # ==============================================
    # ОБРАБОТКА НА ВСЕКИ ПРОДУКТОВ БЛОК
    # ==============================================

    for item_index, item_data in enumerate(
        item_starts
    ):
        block_start = item_data["line_index"]

        if item_index + 1 < len(item_starts):
            block_end = item_starts[
                item_index + 1
            ]["line_index"]
        else:
            block_end = len(document_lines)

        block_lines = document_lines[
            block_start:block_end
        ]

        position_no = item_data["position_no"]
        invoice_item = item_data["item_no"]
        description = item_data["description"]
        start_page = item_data["page"]

        invoice_quantity = None
        invoice_unit = None
        gross_amount = None
        discount_amount = 0.0
        quantity_source_line = ""

        # ==========================================
        # QUANTITY И GROSS AMOUNT
        # ==========================================

        for block_line_data in block_lines:
            block_line = block_line_data["text"]

            quantity_match = (
                quantity_amount_pattern.search(
                    block_line
                )
            )

            if not quantity_match:
                continue

            parsed_quantity = parse_european_number(
                quantity_match.group(1)
            )

            parsed_unit = (
                quantity_match.group(2)
                .upper()
                .strip()
            )

            parsed_amount = parse_european_number(
                quantity_match.group(3)
            )

            if parsed_quantity is None:
                continue

            if parsed_amount is None:
                continue

            if parsed_quantity <= 0:
                continue

            if parsed_amount < 0:
                continue

            invoice_quantity = parsed_quantity
            invoice_unit = parsed_unit
            gross_amount = parsed_amount
            quantity_source_line = block_line

            # Вземаме само първия валиден ред
            # Quantity + PCE/CAR + Amount в блока.
            break

        if invoice_quantity is None:
            continue

        if gross_amount is None:
            continue

        # ==========================================
        # DISCOUNT
        # ==========================================

        discount_source_line = ""

        for block_line_data in block_lines:
            block_line = block_line_data["text"]

            if (
                "ADD DISC" not in block_line.upper()
                and
                "MDEU EXP" not in block_line.upper()
            ):
                continue

            discount_match = discount_pattern.search(
                block_line
            )

            if discount_match:
                parsed_discount = (
                    parse_european_number(
                        discount_match.group(2)
                    )
                )

                if parsed_discount is not None:
                    if parsed_discount >= 0:
                        discount_amount += (
                            parsed_discount
                        )
                        discount_source_line = (
                            block_line
                        )

                continue

            fallback_match = (
                discount_fallback_pattern.search(
                    block_line
                )
            )

            if fallback_match:
                parsed_discount = (
                    parse_european_number(
                        fallback_match.group(1)
                    )
                )

                if parsed_discount is not None:
                    if parsed_discount >= 0:
                        discount_amount += (
                            parsed_discount
                        )
                        discount_source_line = (
                            block_line
                        )

        # Нетна стойност след отстъпката.
        net_amount = (
            gross_amount
            - discount_amount
        )

        if net_amount < 0:
            continue

        # ==========================================
        # QUANTITY CONVERSION
        # ==========================================

        pack_quantity = extract_motul_pack_quantity(
            description
        )

        if invoice_unit == "CAR":
            final_quantity = (
                invoice_quantity
                * pack_quantity
            )
        else:
            final_quantity = invoice_quantity

        if final_quantity <= 0:
            continue

        # ==========================================
        # PRICE PER PIECE
        # ==========================================

        unit_price = (
            net_amount
            / final_quantity
        )

        # Пазим достатъчно точност.
        # В Excel стойността се визуализира според
        # зададения number_format.
        unit_price = round(
            unit_price,
            10
        )

        calculated_net_amount = (
            final_quantity
            * unit_price
        )

        tolerance = max(
            0.02,
            abs(net_amount) * 0.0001
        )

        calculation_ok = (
            abs(
                calculated_net_amount
                - net_amount
            )
            <= tolerance
        )

        rows.append({
            "invoice_item":
                invoice_item,

            "normalized_invoice_item":
                normalize_item_number(
                    invoice_item
                ),

            "description":
                description,

            "position_no":
                position_no,

            "invoice_qty":
                invoice_quantity,

            "invoice_unit":
                invoice_unit,

            "pack_qty":
                pack_quantity,

            "qty":
                final_quantity,

            "gross_amount":
                gross_amount,

            "discount_amount":
                discount_amount,

            "net_amount":
                net_amount,

            "price":
                unit_price,

            # Match функцията очаква line_total.
            # Тук подаваме нетната стойност след
            # приспадане на отстъпката.
            "line_total":
                net_amount,

            "calculation_ok":
                calculation_ok,

            "page":
                start_page,

            "source_line":
                quantity_source_line,

            "discount_source_line":
                discount_source_line
        })

    invoice_number = extract_invoice_number(
        full_text,
        pdf_file.name
    )

    return {
        "invoice_number": invoice_number,
        "pages": processed_pages,
        "rows": rows
    }
 # ======================================================
# CONTAINER INVOICE DETECTION
# ======================================================

def is_container_invoice_pdf(pdf_file):
    """
    Разпознава фактура от типа:

    Packing List:
    EAN + Description + Container No. + Seal No. + Qty + Weight

    Commercial Invoice:
    EAN + Description + Qty + Unit Price + Amount
    """

    pdf_file.seek(0)

    detection_text = ""

    try:
        with pdfplumber.open(pdf_file) as pdf:

            for page in pdf.pages[:2]:

                page_text = page.extract_text(
                    x_tolerance=2,
                    y_tolerance=3
                )

                if page_text:
                    detection_text += (
                        page_text.upper()
                        + "\n"
                    )

    except Exception:
        pdf_file.seek(0)
        return False

    pdf_file.seek(0)

    required_markers = [
        "EAN CODES",
        "CONTAINER NO",
        "SEAL NO",
        "UNIT PRICE",
        "AMOUNT"
    ]

    found_markers = sum(
        1
        for marker in required_markers
        if marker in detection_text
    )

    # Приемаме PDF-а за контейнерна фактура,
    # ако са открити поне 4 от основните маркери.
    return found_markers >= 4


# ======================================================
# CONTAINER INVOICE NUMBER
# ======================================================

def extract_container_invoice_number(
    complete_text,
    fallback_file_name
):
    """
    Поддържа буквено-цифрови номера като:

    GTH2673598AO
    """

    patterns = [
        r"INVOICE\s+NO\.?\s*[:#]?\s*"
        r"([A-Z0-9][A-Z0-9\-_/]+)",

        r"INVOICE\s+NUMBER\s*[:#]?\s*"
        r"([A-Z0-9][A-Z0-9\-_/]+)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            complete_text,
            re.IGNORECASE
        )

        if match:

            invoice_number = (
                match.group(1)
                .strip()
            )

            invoice_number = re.sub(
                r"[^A-Z0-9\-_]",
                "",
                invoice_number,
                flags=re.IGNORECASE
            )

            if invoice_number:
                return invoice_number

    return re.sub(
        r"\.pdf$",
        "",
        fallback_file_name,
        flags=re.IGNORECASE
    )


# ======================================================
# NORMALIZE EAN
# ======================================================

def normalize_ean(value):

    if value is None:
        return ""

    ean = re.sub(
        r"\D",
        "",
        str(value)
    )

    return ean.strip()


# ======================================================
# CONTAINER PDF PARSER
# ======================================================

def extract_container_invoice_rows(pdf_file):
    """
    Извлича позиции от комбиниран PDF:

    1. Packing List:
       EAN
       Description
       Container No.
       Seal No.
       Container Qty
       Weight

    2. Commercial Invoice:
       EAN
       Description
       Total Qty
       Unit Price
       Amount

    Свързването между двете части се извършва по EAN.

    Крайният резултат съдържа отделен ред за:
    EAN + Container + Container Qty + Unit Price
    """

    rows = []
    complete_text = ""
    processed_pages = 0
    document_lines = []

    pdf_file.seek(0)

    with pdfplumber.open(pdf_file) as pdf:

        processed_pages = len(pdf.pages)

        for page_number, page in enumerate(
            pdf.pages,
            start=1
        ):

            page_text = page.extract_text(
                x_tolerance=2,
                y_tolerance=3
            )

            if not page_text:
                continue

            complete_text += (
                page_text
                + "\n"
            )

            for original_line in page_text.splitlines():

                clean_line = " ".join(
                    str(original_line).split()
                )

                if not clean_line:
                    continue

                document_lines.append({
                    "page": page_number,
                    "text": clean_line
                })

    # Един нормализиран текст позволява parser-ът
    # да работи и когато PDF извличането разделя
    # една позиция между няколко визуални реда.
    normalized_document_text = " ".join(
        line_data["text"]
        for line_data in document_lines
    )

    normalized_document_text = re.sub(
        r"\s+",
        " ",
        normalized_document_text
    ).strip()

    # ==================================================
    # РАЗДЕЛЯНЕ НА PACKING LIST И COMMERCIAL INVOICE
    # ==================================================

    commercial_header_patterns = [
        (
            r"EAN\s+CODES?\s+"
            r"DESCRIPTION\s+"
            r"QTY\s+"
            r"UNIT\s+PRICE\s*"
            r"AMOUNT"
        ),
        (
            r"EAN\s+CODES?\s+"
            r"DESCRIPTION\s+"
            r"QTY\s+"
            r"UNIT\s+PRICE"
        )
    ]

    commercial_start = None

    for header_pattern in commercial_header_patterns:

        header_match = re.search(
            header_pattern,
            normalized_document_text,
            re.IGNORECASE
        )

        if header_match:

            commercial_start = (
                header_match.end()
            )

            break

    if commercial_start is None:

        return {
            "invoice_number":
                extract_container_invoice_number(
                    complete_text,
                    pdf_file.name
                ),

            "pages":
                processed_pages,

            "rows":
                [],

            "containers":
                [],

            "packing_rows_count":
                0,

            "invoice_prices_count":
                0,

            "warnings": [
                (
                    "Не е намерено началото на "
                    "Commercial Invoice таблицата."
                )
            ]
        }

    packing_text = (
        normalized_document_text[
            :commercial_start
        ]
    )

    commercial_text = (
        normalized_document_text[
            commercial_start:
        ]
    )

    # ==================================================
    # PACKING LIST
    # ==================================================

    packing_rows = []

    # Примерна позиция:
    #
    # 8859903117023
    # TH205/55R16PR[SW608]91H TRAZANO TL UL
    # HASU4554985
    # ML-TH1226702
    # 549
    # 4609.00
    #
    # Container:
    # четири букви + седем цифри
    #
    # Seal:
    # буквено-цифров код с тире

    packing_pattern = re.compile(
        r"(?P<ean>\d{13})"
        r"\s+"
        r"(?P<description>.*?)"
        r"\s+"
        r"(?P<container>[A-Z]{4}\d{7})"
        r"\s+"
        r"(?P<seal>[A-Z0-9]+(?:-[A-Z0-9]+)+)"
        r"\s+"
        r"(?P<qty>\d+(?:[.,]\d+)?)"
        r"\s+"
        r"(?P<weight>\d+(?:[.,]\d+)?)"
        r"(?="
        r"\s+\d{13}"
        r"|"
        r"\s+\d+(?:[.,]\d+)?"
        r"\s+\d+(?:[.,]\d+)?"
        r"\s+\d{13}"
        r"|"
        r"\s+TO:"
        r"|"
        r"\s+INVOICE"
        r"|"
        r"$"
        r")",
        re.IGNORECASE
    )

    for match in packing_pattern.finditer(
        packing_text
    ):

        ean = normalize_ean(
            match.group("ean")
        )

        description = " ".join(
            match.group("description").split()
        ).strip()

        container_no = (
            match.group("container")
            .upper()
            .strip()
        )


        container_qty = parse_european_number(
            match.group("qty")
        )

        weight = parse_european_number(
            match.group("weight")
        )

        if len(ean) != 13:
            continue

        if container_qty is None:
            continue

        if container_qty <= 0:
            continue

        packing_rows.append({
            "ean":
                ean,

            "description":
                description,

            "container_no":
                container_no,

            "seal_no":
                seal_no,

            "container_qty":
                container_qty,

            "weight":
                weight
        })

    # ==================================================
    # COMMERCIAL INVOICE PRICES
    # ==================================================

    invoice_price_rows = []

    # Ограничаваме таблицата преди банковата
    # информация и останалия текст след позициите.
    commercial_end_patterns = [
        r"\s+FUND\s+FOR\s+FEW\s+CLAIMS",
        r"\s+BANK\s+INFO",
        r"\s+TOTAL\s+VALUE",
        r"\s+SELLER:"
    ]

    commercial_products_text = (
        commercial_text
    )

    commercial_end_positions = []

    for end_pattern in commercial_end_patterns:

        end_match = re.search(
            end_pattern,
            commercial_products_text,
            re.IGNORECASE
        )

        if end_match:
            commercial_end_positions.append(
                end_match.start()
            )

    if commercial_end_positions:

        commercial_products_text = (
            commercial_products_text[
                :min(commercial_end_positions)
            ]
        )

    # Позицията се отделя до следващия EAN.
    invoice_block_pattern = re.compile(
        r"(?P<ean>\d{13})"
        r"\s*"
        r"(?P<body>.*?)"
        r"(?="
        r"\s+\d{13}"
        r"|"
        r"$"
        r")",
        re.IGNORECASE
    )

    for block_match in invoice_block_pattern.finditer(
        commercial_products_text
    ):

        ean = normalize_ean(
            block_match.group("ean")
        )

        body = " ".join(
            block_match.group("body").split()
        ).strip()

        if len(ean) != 13:
            continue

        if not body:
            continue

        # В края на всяка invoice позиция очакваме:
        #
        # Qty + Unit Price + Amount
        #
        # Използваме търсене от края на блока.
        numeric_tail_pattern = re.compile(
            r"^(?P<description>.*?)"
            r"\s*"
            r"(?P<qty>\d+)"
            r"\s+"
            r"(?P<unit_price>"
            r"\d+(?:[.,]\d+)?"
            r")"
            r"\s+"
            r"(?P<amount>"
            r"\d+(?:[.,]\d+)?"
            r")"
            r"\s*$",
            re.IGNORECASE
        )

        numeric_match = numeric_tail_pattern.match(
            body
        )

        if not numeric_match:
            continue

        description = " ".join(
            numeric_match
            .group("description")
            .split()
        ).strip()

        invoice_qty = parse_european_number(
            numeric_match.group("qty")
        )

        unit_price = parse_european_number(
            numeric_match.group("unit_price")
        )

        invoice_amount = parse_european_number(
            numeric_match.group("amount")
        )

        if invoice_qty is None:
            continue

        if unit_price is None:
            continue

        if invoice_amount is None:
            continue

        if invoice_qty <= 0:
            continue

        if unit_price < 0:
            continue

        expected_amount = (
            invoice_qty
            * unit_price
        )

        amount_tolerance = max(
            0.05,
            abs(invoice_amount) * 0.0001
        )

        calculation_ok = (
            abs(
                expected_amount
                - invoice_amount
            )
            <= amount_tolerance
        )

        invoice_price_rows.append({
            "ean":
                ean,

            "description":
                description,

            "invoice_qty":
                invoice_qty,

            "unit_price":
                unit_price,

            "invoice_amount":
                invoice_amount,

            "calculation_ok":
                calculation_ok
        })

    # ==================================================
    # PRICE INDEX ПО EAN
    # ==================================================

    price_index = {}

    for price_row in invoice_price_rows:

        ean = price_row["ean"]

        # Първият валиден запис за EAN се пази.
        if ean not in price_index:

            price_index[ean] = price_row

    # ==================================================
# JOIN:
# PACKING LIST + COMMERCIAL INVOICE
# ==================================================

    warnings = []

    for packing_row in packing_rows:

        ean = packing_row["ean"]

        price_data = price_index.get(
            ean
        )

        if price_data is None:

            warnings.append(
                f"Няма цена в Commercial Invoice "
                f"за EAN {ean}, контейнер "
                f"{packing_row['container_no']}."
            )

            continue

        container_qty = (
            packing_row["container_qty"]
        )

        unit_price = (
            price_data["unit_price"]
        )

        container_line_total = (
            container_qty
            * unit_price
        )

        rows.append({
            "invoice_item":
                ean,

            "normalized_invoice_item":
                normalize_item_number(
                    ean
                ),

            "description":
                packing_row["description"],

            "container_no":
                packing_row["container_no"],

            "seal_no":
                packing_row["seal_no"],

            "qty":
                container_qty,

            "price":
                unit_price,

            "line_total":
                container_line_total,

            "weight":
                packing_row.get(
                    "weight",
                    None
                ),

            "invoice_total_qty":
                price_data["invoice_qty"],

            "invoice_total_amount":
                price_data["invoice_amount"],

            "calculation_ok":
                price_data["calculation_ok"],

            "page":
                1,

            "source_line":
                (
                    f"{ean} | "
                    f"{packing_row['container_no']} | "
                    f"{container_qty} | "
                    f"{unit_price}"
                )
        })

    containers = sorted(
        {
            row["container_no"]
            for row in rows
            if row.get(
                "container_no"
            )
        }
    )

    invoice_number = (
        extract_container_invoice_number(
            complete_text,
            pdf_file.name
        )
    )

    return {
        "invoice_number":
            invoice_number,

        "pages":
            processed_pages,

        "rows":
            rows,

        "containers":
            containers,

        "packing_rows_count":
            len(packing_rows),

        "invoice_prices_count":
            len(invoice_price_rows),

        "warnings":
            warnings
    }
                        
# ======================================================
# BUILD FAST CROSS-REFERENCE INDEXES
# ======================================================

def build_cross_reference_indexes(
    cross_refs
):

    direct_index = {}
    item_index = {}

    for _, row in cross_refs.iterrows():

        vendor_no = str(
            row.get(
                "vendor_no",
                ""
            )
        ).strip()

        cross_reference_no = str(
            row.get(
                "cross_reference_no",
                ""
            )
        ).strip()

        item_no = str(
            row.get(
                "item_no",
                ""
            )
        ).strip()

        normalized_cross_reference = (
            normalize_item_number(
                row.get(
                    "normalized_cross_reference",
                    cross_reference_no
                )
            )
        )

        normalized_item_no = (
            normalize_item_number(
                row.get(
                    "normalized_item_no",
                    item_no
                )
            )
        )

        if normalized_cross_reference:

            if (
                normalized_cross_reference
                not in direct_index
            ):

                direct_index[
                    normalized_cross_reference
                ] = {
                    "vendor_no":
                        vendor_no,
                
                    "cross_reference_no":
                        cross_reference_no,
                
                    "item_no":
                        item_no,
                
                    "internal_item_no":
                        item_no
                }

        if normalized_item_no:

            if (
                normalized_item_no
                not in item_index
            ):

               item_index[
                    normalized_item_no
                ] = {
                    "vendor_no":
                        vendor_no,
                
                    "cross_reference_no":
                        cross_reference_no,
                
                    "item_no":
                        item_no,
                
                    "internal_item_no":
                        item_no
                }

    return direct_index, item_index


# ======================================================
# MATCH INVOICE ROWS TO NEON
# ======================================================

def match_invoice_rows(
    invoice_rows,
    selected_vendor_no,
    cross_refs
):

    (
        direct_index,
        item_index
    ) = build_cross_reference_indexes(
        cross_refs
    )

    matched_rows = []

    for invoice_row in invoice_rows:

        invoice_item = invoice_row[
            "invoice_item"
        ]

        normalized_item = invoice_row[
            "normalized_invoice_item"
        ]
        
        quantity = invoice_row["qty"]
        price = invoice_row["price"]

        match_status = "not_found"

        output_cross_reference = (
            f"❗ {invoice_item}"
        )

        internal_item_no = ""

        # ======================================
        # DIRECT CROSS REFERENCE
        # ======================================

        if normalized_item in direct_index:

            reference = direct_index[
                normalized_item
            ]

            output_cross_reference = (
                reference[
                    "internal_item_no"
                ]
            )
            
            internal_item_no = (
                reference[
                    "cross_reference_no"
                ]
            )

            match_status = "direct"

        # ======================================
        # FALLBACK THROUGH ITEM NO.
        # ======================================

        elif normalized_item in item_index:

            reference = item_index[
                normalized_item
            ]

            output_cross_reference = (
                "⚠️ "
                + reference[
                    "cross_reference_no"
                ]
            )

            internal_item_no = (
                reference["internal_item_no"]
            )

            match_status = "item_no"

        matched_rows.append({

            "Cross-Reference Type No.":
                selected_vendor_no,
        
            "Item No.":
                output_cross_reference,
        
            "Cross-Reference No.":
                invoice_item,
        
            "Qty":
                quantity,
        
            "Price 1 pc":
                price,
            "_invoice_item":
                invoice_item,

            "_internal_item_no":
                internal_item_no,

            "_status":
                match_status,

            "_page":
                invoice_row["page"],
            "_container_no":
                invoice_row.get(
                    "container_no",
                    ""
                ),
            
            "_seal_no":
                invoice_row.get(
                    "seal_no",
                    ""
                ),
            
            "_weight":
                invoice_row.get(
                    "weight",
                    None
                ),

            "_line_total":
                invoice_row["line_total"],

            "_calculation_ok":
                invoice_row[
                    "calculation_ok"
                ]
        })

    return pd.DataFrame(
        matched_rows
    )


# ======================================================
# CREATE EXCEL
# ======================================================

def create_invoice_excel(result_df):

    export_df = result_df[
        [
            "Cross-Reference Type No.",
            "Item No.",
            "Cross-Reference No.",
            "Qty",
            "Price 1 pc"
        ]
    
    ].copy()

    grand_total = (
        pd.to_numeric(
            export_df["Qty"],
            errors="coerce"
        )
        *
        pd.to_numeric(
            export_df["Price 1 pc"],
            errors="coerce"
        )
    ).sum()

    total_row = pd.DataFrame([
        {
            "Cross-Reference Type No.": "",
            "Cross-Reference No.": "TOTAL",
            "Qty": "",
            "Price 1 pc": grand_total
        }
    ])

    final_df = pd.concat(
        [
            export_df,
            total_row
        ],
        ignore_index=True
    )

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        final_df.to_excel(
            writer,
            sheet_name="Invoice",
            index=False
        )

        worksheet = writer.sheets[
            "Invoice"
        ]

        worksheet.freeze_panes = "A2"

        worksheet.auto_filter.ref = (
            f"A1:E{len(final_df) + 1}"
        )

        worksheet.column_dimensions[
            "A"
        ].width = 22

        worksheet.column_dimensions[
            "B"
        ].width = 18

        worksheet.column_dimensions[
            "C"
        ].width = 28

        worksheet.column_dimensions[
            "D"
        ].width = 12

        header_fill = PatternFill(
            fill_type="solid",
            fgColor="D71919"
        )

        header_font = Font(
            color="FFFFFF",
            bold=True
        )

        thin_border = Border(
            left=Side(
                style="thin",
                color="D9D9D9"
            ),
            right=Side(
                style="thin",
                color="D9D9D9"
            ),
            top=Side(
                style="thin",
                color="D9D9D9"
            ),
            bottom=Side(
                style="thin",
                color="D9D9D9"
            )
        )

        for cell in worksheet[1]:

            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        last_product_row = (
            len(export_df) + 1
        )

        for excel_row in range(
            2,
            last_product_row + 1
        ):

            reference_cell = worksheet.cell(
                row=excel_row,
                column=2
            )

            reference_value = str(
                reference_cell.value or ""
            )

            if reference_value.startswith(
                "⚠️"
            ):

                reference_cell.font = Font(
                    bold=True,
                    color="FF8C00"
                )

            elif reference_value.startswith(
                "❗"
            ):

                reference_cell.font = Font(
                    bold=True,
                    color="FF0000"
                )

            worksheet.cell(
                row=excel_row,
                column=4
            ).number_format = "0.###"
            
            worksheet.cell(
                row=excel_row,
                column=5
            ).number_format = "0.000000"

            for column_number in range(
                1,
                6
            ):

                worksheet.cell(
                    row=excel_row,
                    column=column_number
                ).border = thin_border

        total_excel_row = (
            len(final_df) + 1
        )

        total_fill = PatternFill(
            fill_type="solid",
            fgColor="FFF2CC"
        )

        for column_number in range(
            1,
            6
        ):

            total_cell = worksheet.cell(
                row=total_excel_row,
                column=column_number
            )

            total_cell.fill = total_fill

            total_cell.font = Font(
                bold=True,
                color="C00000"
            )

            total_cell.border = thin_border

        worksheet.cell(
            row=total_excel_row,
            column=4
        ).number_format = "0.000000"

    output.seek(0)

    return output
    # ======================================================
# PREPARE DATA FOR DIRECT PRN
# ======================================================

def prepare_direct_prn_dataframe(result_df):

    if result_df is None or result_df.empty:
        return pd.DataFrame(
            columns=[
                "Item",
                "Qty",
                "Price 1 pc"
            ]
        )

    prn_df = pd.DataFrame()

    # За PRN използваме вътрешния Inter Cars номер,
    # който в готовия PDF резултат е в Item No.
    prn_df["Item"] = (
        result_df["Item No."]
        .astype(str)
        .str.replace("⚠️", "", regex=False)
        .str.replace("❗", "", regex=False)
        .str.strip()
    )

    prn_df["Qty"] = pd.to_numeric(
        result_df["Qty"],
        errors="coerce"
    ).fillna(0)

    prn_df["Price 1 pc"] = pd.to_numeric(
        result_df["Price 1 pc"],
        errors="coerce"
    ).fillna(0)

    # Премахва празните и техническите редове.
    prn_df = prn_df[
        prn_df["Item"].astype(str).str.strip() != ""
    ].copy()

    prn_df = prn_df[
        prn_df["Item"].astype(str).str.upper() != "TOTAL"
    ].copy()

    prn_df = prn_df[
        prn_df["Qty"] != 0
    ].copy()

    return prn_df.reset_index(drop=True)
# ======================================================
# PREPARE PRN DATA BY CONTAINER
# ======================================================

def prepare_container_prn_data(result_df):
    """
    Разделя готовия резултат по контейнер.

    Очаква result_df да съдържа:
    - Item No.
    - Qty
    - Price 1 pc
    - _container_no

    Връща речник:

    {
        "HASU4554985": DataFrame,
        "HASU4849034": DataFrame
    }
    """

    container_data = {}

    if result_df is None:
        return container_data

    if result_df.empty:
        return container_data

    if "_container_no" not in result_df.columns:
        return container_data

    working_df = result_df.copy()

    working_df["_container_no"] = (
        working_df["_container_no"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # Оставяме само редовете с реален контейнер.
    working_df = working_df[
        working_df["_container_no"] != ""
    ].copy()

    if working_df.empty:
        return container_data

    for container_no, container_df in working_df.groupby(
        "_container_no",
        sort=True
    ):

        prn_df = pd.DataFrame()

        # Item No. съдържа вътрешния Inter Cars номер
        # след извършения Neon match.
        prn_df["Item"] = (
            container_df["Item No."]
            .fillna("")
            .astype(str)
            .str.replace(
                "⚠️",
                "",
                regex=False
            )
            .str.replace(
                "❗",
                "",
                regex=False
            )
            .str.strip()
        )

        prn_df["Qty"] = pd.to_numeric(
            container_df["Qty"],
            errors="coerce"
        )

        prn_df["Price 1 pc"] = pd.to_numeric(
            container_df["Price 1 pc"],
            errors="coerce"
        )

        # Запазваме EAN само за визуална проверка.
        if "Cross-Reference No." in container_df.columns:

            prn_df["EAN"] = (
                container_df[
                    "Cross-Reference No."
                ]
                .fillna("")
                .astype(str)
                .str.strip()
                .values
            )

        # Запазваме статуса от Neon match-а.
        if "_status" in container_df.columns:

            prn_df["_status"] = (
                container_df["_status"]
                .fillna("")
                .astype(str)
                .values
            )

        # Премахва празни и технически редове.
        prn_df = prn_df[
            prn_df["Item"] != ""
        ].copy()

        prn_df = prn_df[
            prn_df["Item"].str.lower() != "nan"
        ].copy()

        prn_df = prn_df[
            prn_df["Item"].str.upper() != "TOTAL"
        ].copy()

        # Само редове с валидно количество и цена.
        prn_df = prn_df[
            prn_df["Qty"].notna()
            &
            prn_df["Price 1 pc"].notna()
            &
            (
                prn_df["Qty"] > 0
            )
            &
            (
                prn_df["Price 1 pc"] >= 0
            )
        ].copy()

        if prn_df.empty:
            continue

        container_data[
            str(container_no).strip()
        ] = prn_df.reset_index(
            drop=True
        )

    return container_data


# ======================================================
# CREATE ONE PRN CONTENT
# ======================================================

def create_prn_content(prn_df):
    """
    Създава текстовото съдържание на един PRN файл.
    """

    prn_lines = []

    if prn_df is None:
        return ""

    if prn_df.empty:
        return ""

    for _, row in prn_df.iterrows():

        item = str(
            row.get(
                "Item",
                ""
            )
        ).strip()

        if (
            item == ""
            or item.lower() == "nan"
            or item.upper() == "TOTAL"
        ):
            continue

        try:
            qty = int(
                round(
                    float(
                        row.get(
                            "Qty",
                            0
                        )
                    )
                )
            )

            price = float(
                row.get(
                    "Price 1 pc",
                    0
                )
            )

        except Exception:
            continue

        if qty <= 0:
            continue

        if price < 0:
            continue

        price_str = (
            f"{price:.6f}"
            .replace(".", ",")
        )

        spaces_before_qty = max(
            1,
            25
            - len(item)
            - len(str(qty))
        )

        prn_line = (
            item
            + (
                " "
                * spaces_before_qty
            )
            + str(qty)
            + (" " * 6)
            + price_str
        )

        prn_lines.append(
            prn_line
        )

    return "\r\n".join(
        prn_lines
    )


# ======================================================
# CREATE ZIP WITH ALL CONTAINER PRN FILES
# ======================================================

def create_container_prn_zip(
    container_prn_contents
):
    """
    Получава речник:

    {
        "HASU4554985": "PRN content",
        "HASU4849034": "PRN content"
    }

    и създава ZIP с отделен PRN файл
    за всеки контейнер.
    """

    zip_output = io.BytesIO()

    with zipfile.ZipFile(
        zip_output,
        mode="w",
        compression=zipfile.ZIP_DEFLATED
    ) as zip_file:

        for (
            container_no,
            prn_content
        ) in container_prn_contents.items():

            if not prn_content:
                continue

            safe_container_no = re.sub(
                r"[^A-Z0-9\-_]",
                "_",
                str(container_no).upper()
            )

            prn_file_name = (
                f"{safe_container_no}.prn"
            )

            zip_file.writestr(
                prn_file_name,
                prn_content.encode(
                    "utf-8"
                )
            )

    zip_output.seek(0)

    return zip_output
# ======================================================
# PDF → EXCEL
# ======================================================

if page == "📄 PDF → Excel":

    st.markdown(
        """
        <div class="main-card">
            <h2>📄 PDF → Excel</h2>
            <p>
                Качи PDF фактура. Приложението ще
                извлече реалните продуктови позиции,
                количество и единична цена, след което
                ще провери номерата в Neon.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    uploaded_pdfs = st.file_uploader(
        "📄 Качи една или няколко PDF фактури",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_upload_main"
    )

    if uploaded_pdfs:

        with st.spinner(
            "Зареждане на Cross References от Neon..."
        ):

            cross_refs = load_cross_references(
                selected_vendor_no
            )

        if cross_refs.empty:

            if cross_refs.empty:

                st.warning(
                    f"Няма Cross References за "
                    f"{selected_vendor_no}. "
                    f"Ще се покажат оригиналните номера."
                )
            
                cross_refs = pd.DataFrame(
                    columns=[
                        "vendor_no",
                        "cross_reference_no",
                        "item_no",
                        "normalized_cross_reference",
                        "normalized_item_no"
                    ]
                )

        st.success(
            f"Cross References в Neon: "
            f"{len(cross_refs)}"
        )

        all_results = []
        processed_pages = 0
        detected_invoice_rows = 0
        invoice_numbers = []

        with st.spinner(
            "Разпознаване на фактурните позиции..."
        ):

            for pdf_file in uploaded_pdfs:

                # Нова контейнерна фактура:
                # Packing List + Commercial Invoice
                if is_container_invoice_pdf(
                    pdf_file
                ):
            
                    extracted = (
                        extract_container_invoice_rows(
                            pdf_file
                        )
                    )
            
                # Castrol
                elif selected_vendor_no == "VEN0002914":
            
                    extracted = extract_castrol_rows(
                        pdf_file
                    )
            
                # Motul
                elif selected_vendor_no == "VEN0001554":
            
                    extracted = extract_motul_rows(
                        pdf_file
                    )
            
                # Federal и останалите доставчици
                else:
            
                    extracted = extract_federal_rows(
                        pdf_file
                    )

                invoice_numbers.append(
                    extracted[
                        "invoice_number"
                    ]
                )

                processed_pages += extracted[
                    "pages"
                ]

                detected_invoice_rows += len(
                    extracted["rows"]
                )
                if extracted["rows"]:
                    st.write("FIRST PDF ROW")
                    st.write(extracted["rows"][0])

                matched_df = match_invoice_rows(
                    invoice_rows=extracted[
                        "rows"
                    ],
                    selected_vendor_no=(
                        selected_vendor_no
                    ),
                    cross_refs=cross_refs
                )

                if not matched_df.empty:

                    matched_df["_file"] = (
                        pdf_file.name
                    )

                    all_results.append(
                        matched_df
                    )

        if not all_results:

            st.error(
                "Не бяха открити фактурни "
                "позиции в PDF."
            )

            st.stop()
        final_result_df = pd.concat(
            all_results,
            ignore_index=True
        )

        final_result_df = (
            final_result_df
            .drop_duplicates(
                subset=[
                    "_file",
                    "_invoice_item",
                    "_container_no",
                    "Qty",
                    "Price 1 pc",
                    "_page"
                ]
            )
            .reset_index(drop=True)
        )

        direct_count = int(
            (
                final_result_df["_status"]
                == "direct"
            ).sum()
        )

        item_match_count = int(
            (
                final_result_df["_status"]
                == "item_no"
            ).sum()
        )

        not_found_count = int(
            (
                final_result_df["_status"]
                == "not_found"
            ).sum()
        )

        invalid_total_count = int(
            (
                final_result_df[
                    "_calculation_ok"
                ]
                == False
            ).sum()
        )

        col1, col2, col3, col4 = (
            st.columns(4)
        )

        with col1:

            st.metric(
                "📄 Страници",
                processed_pages
            )

        with col2:

            st.metric(
                "✅ Директни",
                direct_count
            )

        with col3:

            st.metric(
                "⚠️ Чрез Item No.",
                item_match_count
            )

        with col4:

            st.metric(
                "❗ Ненамерени",
                not_found_count
            )

        st.info(
            f"Открити фактурни позиции: "
            f"{detected_invoice_rows} | "
            f"Редове в резултата: "
            f"{len(final_result_df)} | "
            f"Редове с непотвърден Total: "
            f"{invalid_total_count}"
        )
        preview_df = final_result_df[
            [
                "Cross-Reference Type No.",
                "Item No.",
                "Cross-Reference No.",
                "Qty",
                "Price 1 pc"
            ]
        ].copy()

        st.subheader(
            "📋 Разпознати фактурни позиции"
        )

        st.dataframe(
            preview_df,
            use_container_width=True,
            hide_index=True
        )
        if selected_vendor_no == "ТУК_ПОСТАВИ_VENDOR_NO_НА_MOTUL":

            with st.expander(
                "🔎 Motul изчисления"
            ):
                motul_debug_df = final_result_df[
                    [
                        "_invoice_item",
                        "Qty",
                        "Price 1 pc",
                        "_line_total",
                        "_calculation_ok",
                        "_page"
                    ]
                ].copy()
        
                motul_debug_df["Calculated Total"] = (
                    pd.to_numeric(
                        motul_debug_df["Qty"],
                        errors="coerce"
                    )
                    *
                    pd.to_numeric(
                        motul_debug_df["Price 1 pc"],
                        errors="coerce"
                    )
                )
        
                st.dataframe(
                    motul_debug_df,
                    use_container_width=True,
                    hide_index=True
                )

        st.caption(
            "⚠️ = намерен чрез Item No. | "
            "❗ = оригиналният номер от фактурата "
            "не е намерен в Cross Reference базата"
        )

        if len(invoice_numbers) == 1:

            invoice_number = (
                invoice_numbers[0]
            )

            excel_file_name = (
                f"{invoice_number}.xlsx"
            )

        else:

            excel_file_name = (
                "multiple_invoices.xlsx"
            )

        excel_output = create_invoice_excel(
            final_result_df
        )
        
        # Проверяваме дали резултатът съдържа контейнери.
        has_container_data = (
            "_container_no"
            in final_result_df.columns
            and
            final_result_df[
                "_container_no"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne("")
            .any()
        )
        
        if has_container_data:
        
            download_col, prn_col, container_col = (
                st.columns(3)
            )
        
        else:
        
            download_col, prn_col = st.columns(2)
            container_col = None
        
        
        # ==================================================
        # DOWNLOAD EXCEL
        # ==================================================
        
        with download_col:
        
            st.download_button(
                label="📥 Изтегли Excel",
                data=excel_output.getvalue(),
                file_name=excel_file_name,
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True,
                key="download_invoice_excel"
            )
        
        
        # ==================================================
        # STANDARD SINGLE PRN
        # ==================================================
        
        with prn_col:
        
            if st.button(
                "🧾 Зареди общо за PRN",
                use_container_width=True,
                key="load_current_invoice_to_prn"
            ):
        
                direct_prn_df = (
                    prepare_direct_prn_dataframe(
                        final_result_df
                    )
                )
        
                if direct_prn_df.empty:
        
                    st.error(
                        "Няма валидни редове "
                        "за зареждане в PRN."
                    )
        
                else:
        
                    st.session_state[
                        "direct_prn_df"
                    ] = direct_prn_df.copy()
        
                    st.session_state[
                        "direct_prn_name"
                    ] = (
                        str(invoice_numbers[0])
                        if len(invoice_numbers) == 1
                        else "multiple_invoices"
                    )
        
                    # Изчистваме контейнерния режим.
                    st.session_state[
                        "container_prn_data"
                    ] = {}
        
                    st.session_state[
                        "container_prn_loaded"
                    ] = False
        
                    st.session_state[
                        "converter_page"
                    ] = PAGE_PRN
        
                    st.session_state[
                        "prn_loaded"
                    ] = True
        
                    st.rerun()
        
        
                # ==================================================
        # CONTAINER PRN
        # ==================================================

        if (
            has_container_data
            and
            container_col is not None
        ):

            with container_col:

                if st.button(
                    "📦 Раздели по контейнери",
                    use_container_width=True,
                    key="load_containers_to_prn"
                ):

                    container_prn_data = (
                        prepare_container_prn_data(
                            final_result_df
                        )
                    )

                    if not container_prn_data:

                        st.error(
                            "Не бяха подготвени "
                            "валидни контейнерни PRN данни."
                        )

                    else:

                        st.session_state[
                            "container_prn_data"
                        ] = container_prn_data

                        st.session_state[
                            "container_prn_invoice_name"
                        ] = (
                            str(invoice_numbers[0])
                            if len(invoice_numbers) == 1
                            else "multiple_invoices"
                        )

                        st.session_state[
                            "container_prn_loaded"
                        ] = True

                        # Изчистваме стандартния PRN режим.
                        st.session_state[
                            "direct_prn_df"
                        ] = pd.DataFrame(
                            columns=[
                                "Item",
                                "Qty",
                                "Price 1 pc"
                            ]
                        )

                        st.session_state[
                            "prn_loaded"
                        ] = False

                        # Задаваме страницата, която трябва
                        # да се отвори след rerun.
                        st.session_state[
                            "converter_page"
                        ] = PAGE_PRN

                        # converter_menu вече е създаден radio widget.
                        # Не му присвояваме стойност директно.
                        # Изтриваме старото му състояние, за да бъде
                        # създаден наново според converter_page.
                        st.session_state.pop(
                            "converter_menu",
                            None
                        )

                        st.rerun()
           
# ======================================================
# EXCEL → PRN
# ======================================================

if page == PAGE_PRN:

    st.markdown(
        """
        <div class="main-card">
            <h2>🧾 Excel → PRN</h2>
            <p>
                Използвай директно заредените позиции
                от PDF конвертора или качи готов Excel файл.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )
    # ==================================================
    # CONTAINER PRN MODE
    # ==================================================
    
    if (
        st.session_state.get(
            "container_prn_loaded",
            False
        )
        and
        st.session_state.get(
            "container_prn_data",
            {}
        )
    ):
    
        container_data = (
            st.session_state[
                "container_prn_data"
            ]
        )
    
        invoice_name = str(
            st.session_state.get(
                "container_prn_invoice_name",
                "invoice"
            )
        ).strip()
    
        st.success(
            f"✅ Контейнерният прием е зареден. "
            f"Контейнери: {len(container_data)}"
        )
    
        total_rows = sum(
            len(container_df)
            for container_df
            in container_data.values()
        )
    
        total_qty = sum(
            pd.to_numeric(
                container_df["Qty"],
                errors="coerce"
            )
            .fillna(0)
            .sum()
            for container_df
            in container_data.values()
        )
    
        col1, col2, col3 = st.columns(3)
    
        with col1:
            st.metric(
                "📦 Контейнери",
                len(container_data)
            )
    
        with col2:
            st.metric(
                "📋 PRN редове",
                total_rows
            )
    
        with col3:
            st.metric(
                "🔢 Общо количество",
                int(round(total_qty))
            )
    
        st.markdown("---")
    
        container_prn_contents = {}
    
        for container_no, original_df in sorted(
            container_data.items()
        ):
    
            container_df = original_df.copy()
    
            display_columns = [
                column
                for column in [
                    "Item",
                    "EAN",
                    "Qty",
                    "Price 1 pc"
                ]
                if column in container_df.columns
            ]
    
            display_df = container_df[
                display_columns
            ].copy()
    
            container_qty = pd.to_numeric(
                display_df["Qty"],
                errors="coerce"
            ).fillna(0).sum()
    
            unmatched_count = 0
    
            if "_status" in container_df.columns:
    
                unmatched_count = int(
                    (
                        container_df["_status"]
                        == "not_found"
                    ).sum()
                )
    
            with st.expander(
                (
                    f"📦 {container_no} | "
                    f"Позиции: {len(display_df)} | "
                    f"Количество: "
                    f"{int(round(container_qty))}"
                ),
                expanded=True
            ):
    
                if unmatched_count > 0:
    
                    st.warning(
                        f"Ненамерени номера в Neon: "
                        f"{unmatched_count}"
                    )
    
                edited_df = st.data_editor(
                    display_df,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="dynamic",
                    key=(
                        f"container_editor_"
                        f"{container_no}"
                    ),
                    column_config={
                        "Item":
                            st.column_config.TextColumn(
                                "Item",
                                required=True
                            ),
    
                        "EAN":
                            st.column_config.TextColumn(
                                "EAN",
                                disabled=True
                            ),
    
                        "Qty":
                            st.column_config.NumberColumn(
                                "Qty",
                                min_value=0,
                                step=1,
                                format="%.0f"
                            ),
    
                        "Price 1 pc":
                            st.column_config.NumberColumn(
                                "Price 1 pc",
                                min_value=0.0,
                                format="%.6f"
                            )
                    }
                )
    
                edited_prn_df = edited_df[
                    [
                        "Item",
                        "Qty",
                        "Price 1 pc"
                    ]
                ].copy()
    
                edited_prn_df["Item"] = (
                    edited_prn_df["Item"]
                    .fillna("")
                    .astype(str)
                    .str.replace(
                        "⚠️",
                        "",
                        regex=False
                    )
                    .str.replace(
                        "❗",
                        "",
                        regex=False
                    )
                    .str.strip()
                )
    
                edited_prn_df["Qty"] = pd.to_numeric(
                    edited_prn_df["Qty"],
                    errors="coerce"
                )
    
                edited_prn_df[
                    "Price 1 pc"
                ] = pd.to_numeric(
                    edited_prn_df[
                        "Price 1 pc"
                    ],
                    errors="coerce"
                )
    
                edited_prn_df = edited_prn_df[
                    (
                        edited_prn_df["Item"] != ""
                    )
                    &
                    (
                        edited_prn_df["Qty"].notna()
                    )
                    &
                    (
                        edited_prn_df[
                            "Price 1 pc"
                        ].notna()
                    )
                    &
                    (
                        edited_prn_df["Qty"] > 0
                    )
                    &
                    (
                        edited_prn_df[
                            "Price 1 pc"
                        ] >= 0
                    )
                ].copy()
    
                prn_content = create_prn_content(
                    edited_prn_df
                )
    
                container_prn_contents[
                    container_no
                ] = prn_content
    
                if prn_content:
    
                    st.text_area(
                        "PRN съдържание",
                        value=prn_content,
                        height=180,
                        key=(
                            f"container_preview_"
                            f"{container_no}"
                        )
                    )
    
                    st.download_button(
                        label=(
                            f"📥 Изтегли "
                            f"{container_no}.prn"
                        ),
                        data=prn_content.encode(
                            "utf-8"
                        ),
                        file_name=(
                            f"{container_no}.prn"
                        ),
                        mime="text/plain",
                        use_container_width=True,
                        key=(
                            f"download_container_"
                            f"{container_no}"
                        )
                    )
    
                else:
    
                    st.warning(
                        "Няма валидни PRN редове "
                        "за този контейнер."
                    )
    
        valid_contents = {
            container_no: content
            for container_no, content
            in container_prn_contents.items()
            if content
        }
    
        if valid_contents:
    
            zip_output = create_container_prn_zip(
                valid_contents
            )
    
            safe_invoice_name = re.sub(
                r"[^A-Z0-9\-_]",
                "_",
                invoice_name,
                flags=re.IGNORECASE
            )
    
            st.markdown("---")
    
            st.download_button(
                label=(
                    "📦 Изтегли всички "
                    "контейнерни PRN файлове"
                ),
                data=zip_output.getvalue(),
                file_name=(
                    f"{safe_invoice_name}"
                    f"_containers_prn.zip"
                ),
                mime="application/zip",
                use_container_width=True,
                key="download_all_container_prn"
            )
    
            st.success(
                f"✅ ZIP файлът съдържа "
                f"{len(valid_contents)} PRN файла."
            )
    
        if st.button(
            "🧹 Изчисти контейнерния прием",
            use_container_width=True,
            key="clear_container_prn"
        ):
    
            st.session_state[
                "container_prn_data"
            ] = {}
    
            st.session_state[
                "container_prn_invoice_name"
            ] = "invoice"
    
            st.session_state[
                "container_prn_loaded"
            ] = False
    
            st.rerun()
    
        # Не изпълнява стандартния PRN екран,
        # когато работим в контейнерен режим.
        st.stop()

    # ==================================================
    # SESSION STATE
    # ==================================================

    if "direct_prn_df" not in st.session_state:
        st.session_state.direct_prn_df = pd.DataFrame(
            columns=[
                "Item",
                "Qty",
                "Price 1 pc"
            ]
        )

    if "direct_prn_name" not in st.session_state:
        st.session_state.direct_prn_name = "invoice"

    if "prn_loaded" not in st.session_state:
        st.session_state.prn_loaded = False

    # ==================================================
    # EXCEL UPLOAD
    # ==================================================

    uploaded_excel = st.file_uploader(
        "📊 Качи Excel файл",
        type=["xlsx"],
        key="prn_excel_upload"
    )

    # Тези променливи винаги се създават,
    # преди да бъдат използвани.
    prn_source_df = pd.DataFrame(
        columns=[
            "Item",
            "Qty",
            "Price 1 pc"
        ]
    )

    prn_file_name = (
        st.session_state.direct_prn_name
    )

    source_description = ""

    # ==================================================
    # СЪОБЩЕНИЕ ЗА УСПЕШНО ПРЕХВЪРЛЯНЕ
    # ==================================================

    if st.session_state.prn_loaded:

        loaded_rows = len(
            st.session_state.direct_prn_df
        )

        st.success(
            f"✅ Приемът е получен от PDF модула "
            f"и е готов за PRN. "
            f"Заредени редове: {loaded_rows}"
        )

    # ==================================================
    # 1. ДИРЕКТНО ЗАРЕДЕНИ ДАННИ ОТ PDF
    # ==================================================

    if not st.session_state.direct_prn_df.empty:

        prn_source_df = (
            st.session_state.direct_prn_df.copy()
        )

        prn_file_name = (
            st.session_state.direct_prn_name
        )

        source_description = (
            "✅ Данните са заредени директно "
            "от PDF → Excel."
        )

    # ==================================================
    # 2. РЪЧНО КАЧЕН EXCEL
    # ==================================================

    if uploaded_excel is not None:

        try:

            uploaded_df = pd.read_excel(
                uploaded_excel,
                engine="openpyxl"
            )

            uploaded_df.columns = [
                str(col).strip()
                for col in uploaded_df.columns
            ]

            # Поддържа както Item, така и Item No.
            if (
                "Item" not in uploaded_df.columns
                and
                "Item No." in uploaded_df.columns
            ):
                uploaded_df = uploaded_df.rename(
                    columns={
                        "Item No.": "Item"
                    }
                )

            required_cols = [
                "Item",
                "Qty",
                "Price 1 pc"
            ]

            missing_cols = [
                col
                for col in required_cols
                if col not in uploaded_df.columns
            ]

            if missing_cols:

                st.error(
                    "Липсват необходимите колони: "
                    + ", ".join(missing_cols)
                )

            else:

                # Каченият Excel има приоритет
                # пред директно заредените PDF данни.
                prn_source_df = uploaded_df[
                    required_cols
                ].copy()

                prn_file_name = (
                    uploaded_excel.name
                    .replace(".xlsx", "")
                    .replace(".xls", "")
                )

                source_description = (
                    f"✅ Използва се каченият Excel: "
                    f"{uploaded_excel.name}"
                )

        except Exception as error:

            st.error(
                f"Грешка при четене на Excel: "
                f"{type(error).__name__}: {error}"
            )

    # ==================================================
    # PRN PREVIEW И ГЕНЕРИРАНЕ
    # ==================================================

    if not prn_source_df.empty:

        prn_source_df = prn_source_df.copy()

        # Почистване на Item.
        prn_source_df["Item"] = (
            prn_source_df["Item"]
            .astype(str)
            .str.replace(
                "⚠️",
                "",
                regex=False
            )
            .str.replace(
                "❗",
                "",
                regex=False
            )
            .str.strip()
        )

        # Безопасно преобразуване на Qty.
        prn_source_df["Qty"] = pd.to_numeric(
            prn_source_df["Qty"]
            .astype(str)
            .str.replace(
                ",",
                ".",
                regex=False
            ),
            errors="coerce"
        )

        # Безопасно преобразуване на Price.
        prn_source_df["Price 1 pc"] = pd.to_numeric(
            prn_source_df["Price 1 pc"]
            .astype(str)
            .str.replace(
                ",",
                ".",
                regex=False
            ),
            errors="coerce"
        )

        # Премахва празни и TOTAL редове.
        prn_source_df = prn_source_df[
            (
                prn_source_df["Item"] != ""
            )
            &
            (
                prn_source_df["Item"]
                .str.lower()
                != "nan"
            )
            &
            (
                prn_source_df["Item"]
                .str.upper()
                != "TOTAL"
            )
        ].copy()

        # Отделя невалидните редове.
        invalid_rows = prn_source_df[
            prn_source_df["Qty"].isna()
            |
            prn_source_df["Price 1 pc"].isna()
        ].copy()

        # Оставя само валидните PRN редове.
        valid_prn_df = prn_source_df[
            prn_source_df["Qty"].notna()
            &
            prn_source_df["Price 1 pc"].notna()
            &
            (
                prn_source_df["Qty"] > 0
            )
            &
            (
                prn_source_df["Price 1 pc"] >= 0
            )
        ].copy()

        if source_description:

            st.info(
                source_description
            )

        st.subheader(
            "📋 PRN Preview"
        )

        # Тук има само един data_editor
        # и ключът е уникален.
        edited_prn_df = st.data_editor(
            valid_prn_df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="prn_editor_final",
            column_config={
                "Item":
                    st.column_config.TextColumn(
                        "Item",
                        required=True
                    ),

                "Qty":
                    st.column_config.NumberColumn(
                        "Qty",
                        min_value=0,
                        step=1,
                        format="%.0f"
                    ),

                "Price 1 pc":
                    st.column_config.NumberColumn(
                        "Price 1 pc",
                        min_value=0.0,
                        format="%.6f"
                    )
            }
        )

        if not invalid_rows.empty:

            st.warning(
                f"Пропуснати невалидни редове: "
                f"{len(invalid_rows)}"
            )

            with st.expander(
                "Покажи невалидните редове"
            ):

                st.dataframe(
                    invalid_rows,
                    use_container_width=True,
                    hide_index=True
                )

        # ==================================================
        # СЪЗДАВАНЕ НА PRN СЪДЪРЖАНИЕТО
        # ==================================================

        prn_lines = []

        for _, row in edited_prn_df.iterrows():

            item = str(
                row.get(
                    "Item",
                    ""
                )
            ).strip()

            if (
                item == ""
                or item.lower() == "nan"
                or item.upper() == "TOTAL"
            ):
                continue

            try:

                qty = int(
                    round(
                        float(
                            row.get(
                                "Qty",
                                0
                            )
                        )
                    )
                )

                price = float(
                    row.get(
                        "Price 1 pc",
                        0
                    )
                )

            except Exception:
                continue

            if qty <= 0:
                continue

            if price < 0:
                continue

            price_str = (
                f"{price:.6f}"
                .replace(".", ",")
            )

            spaces_before_qty = max(
                1,
                25
                - len(item)
                - len(str(qty))
            )

            prn_line = (
                item
                + (
                    " "
                    * spaces_before_qty
                )
                + str(qty)
                + (" " * 6)
                + price_str
            )

            prn_lines.append(
                prn_line
            )

        prn_content = "\r\n".join(
            prn_lines
        )

        st.metric(
            "Готови PRN редове",
            len(prn_lines)
        )

        if prn_lines:

            st.text_area(
                "PRN съдържание",
                value=prn_content,
                height=250,
                key="prn_text_preview_final"
            )

            st.download_button(
                label="📥 Изтегли PRN",
                data=prn_content.encode(
                    "utf-8"
                ),
                file_name=(
                    f"{prn_file_name}.prn"
                ),
                mime="text/plain",
                use_container_width=True,
                key="download_prn_final"
            )

            st.success(
                f"✅ PRN файлът е готов. "
                f"Генерирани редове: "
                f"{len(prn_lines)}"
            )

        else:

            st.warning(
                "Няма валидни редове "
                "за генериране на PRN."
            )

        # ==================================================
        # ИЗЧИСТВАНЕ НА ПРИЕМА
        # ==================================================

        if st.button(
            "🧹 Изчисти заредения прием",
            use_container_width=True,
            key="clear_prn_final"
        ):

            st.session_state.direct_prn_df = (
                pd.DataFrame(
                    columns=[
                        "Item",
                        "Qty",
                        "Price 1 pc"
                    ]
                )
            )

            st.session_state.direct_prn_name = (
                "invoice"
            )

            st.session_state.prn_loaded = False

            st.rerun()

    else:

        st.info(
            "Няма зареден прием. "
            "Отвори PDF → Excel и натисни "
            "'Зареди директно за PRN' "
            "или качи Excel файл."
        )
