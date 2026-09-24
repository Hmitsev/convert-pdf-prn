import streamlit as st
import pandas as pd
import psycopg2
import base64
import io
import re
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

@st.cache_data(ttl=3600)
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
# SIDEBAR
# ======================================================

st.sidebar.title(
    "PDF → PRN Converter"
)

page = st.sidebar.radio(
    "Menu",
    [
        "📄 PDF → Excel",
        "🧾 Excel → PRN"
    ]
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

            for index in range(len(lines)):

                current_line = lines[index]

                # Пример:
                # 15F710
                # 15FFAE
                # 15BF39

                if not re.match(
                    r'^[A-Z0-9]{5,10}$',
                    current_line
                ):
                    continue

                invoice_item = current_line
                st.write(
                    f"CASTROL CODE FOUND: {invoice_item}"
                )

                qty = None
                price = None

                search_end = min(
                    index + 12,
                    len(lines)
                )

                for j in range(
                    index + 1,
                    search_end
                ):

                    numeric_line = lines[j]

                    if "ST" not in numeric_line:
                        continue
                        print(
                            "CASTROL CHECK:",
                            invoice_item,
                            "->",
                            numeric_line
                        )

                    numbers = re.findall(
                        r'\d+(?:\.\d+)?',
                        numeric_line
                    )

                    if len(numbers) >= 2:

                        try:

                            qty = float(
                                numbers[0]
                            )

                            price = float(
                                numbers[1]
                            )
                            print(
                                "FOUND:",
                                invoice_item,
                                qty,
                                price
                            )

                            break

                        except Exception:

                            pass
                            st.write(
                                "CHECK:",
                                invoice_item,
                                qty,
                                price
                            )

                if (
                    qty is not None
                    and
                    price is not None
                ):

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

                        "page":
                            page_number,

                        "source_line":
                            current_line,

                        "calculation_ok":
                            True
                    })

    invoice_number = extract_invoice_number(
        full_text,
        pdf_file.name
    )

    return {
        "invoice_number":
            invoice_number,

        "pages":
            len(pdf.pages),

        "rows":
            rows
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

            st.error(
                f"Няма Cross References за "
                f"{selected_vendor_no}."
            )

            st.stop()

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

                if selected_vendor_no == "VEN0002914":

                    extracted = extract_castrol_rows(
                        pdf_file
                    )
                
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

        st.download_button(
            label="📥 Изтегли Excel за PRN",
            data=excel_output,
            file_name=excel_file_name,
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
            key="download_invoice_excel"
        )
# ======================================================
# EXCEL → PRN
# ======================================================

if page == "🧾 Excel → PRN":

    st.markdown(
        """
        <div class="main-card">
            <h2>🧾 Excel → PRN</h2>
            <p>
            Качи Excel, генериран от PDF → Excel.
            За PRN се използва колоната Item No.
            (вътрешният Inter Cars номер).
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    uploaded_excel = st.file_uploader(
        "📊 Качи Excel файл",
        type=["xlsx"],
        key="prn_excel_upload"
    )

    if uploaded_excel:

        try:

            df = pd.read_excel(
                uploaded_excel
            )

            df.columns = [
                str(col).strip()
                for col in df.columns
            ]

            required_cols = [
                "Item",
                "Qty",
                "Price 1 pc"
            ]

            missing = [
                col
                for col in required_cols
                if col not in df.columns
            ]

            if missing:

                st.error(
                    f"Липсват колони: "
                    f"{', '.join(missing)}"
                )

                st.stop()

            preview_df = df[
                [
                    "Item",
                    "Qty",
                    "Price 1 pc"
                ]
            ].copy()

            st.subheader(
                "📋 PRN Preview"
            )

            st.dataframe(
                preview_df,
                use_container_width=True,
                hide_index=True
            )

            prn_lines = []

            for _, row in df.iterrows():

                item = str(
                    row["Item"]
                ).strip()

                if (
                    item == ""
                    or
                    item.lower() == "nan"
                    or
                    item.upper() == "TOTAL"
                ):
                    continue

                qty = int(
                    round(
                        float(
                            str(
                                row["Qty"]
                            ).replace(",", ".")
                        )
                    )
                )

                price = float(
                    str(
                        row["Price 1 pc"]
                    ).replace(",", ".")
                )

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

                line = (
                    item
                    + (" " * spaces_before_qty)
                    + str(qty)
                    + (" " * 6)
                    + price_str
                )

                prn_lines.append(
                    line
                )

            prn_content = (
                "\r\n".join(
                    prn_lines
                )
            )

            invoice_name = (
                uploaded_excel.name
                .replace(".xlsx", "")
                .replace(".xls", "")
            )

            st.download_button(
                label="📥 Изтегли PRN",
                data=prn_content.encode(
                    "utf-8"
                ),
                file_name=f"{invoice_name}.prn",
                mime="text/plain",
                use_container_width=True
            )

            st.success(
                f"✅ Генерирани редове: "
                f"{len(prn_lines)}"
            )

        except Exception as error:

            st.error(
                f"Грешка: {error}"
            )
