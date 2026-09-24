import streamlit as st
import pandas as pd
import psycopg2
import base64
import io
import re
import pdfplumber
from psycopg2.extras import RealDictCursor


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
# LOAD VENDORS DATA
# ======================================================

vendors_df = load_vendors()

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
    index=0
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
# ======================================================
# LOAD CROSS REFERENCES
# ======================================================

@st.cache_data(ttl=3600)
def load_cross_references(vendor_no):

    conn = psycopg2.connect(
        DATABASE_URL
    )

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

    df = pd.read_sql_query(
        query,
        conn,
        params=[vendor_no]
    )

    conn.close()

    return df


# ======================================================
# PDF → EXCEL
# ======================================================

if page == "📄 PDF → Excel":

    st.markdown(
        """
        <div class="main-card">
            <h2>📄 PDF → Excel</h2>
            <p>
            Качи PDF фактура и приложението ще
            намери Cross Reference номерата
            за избрания Vendor.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    uploaded_pdfs = st.file_uploader(
        "📄 Качи PDF фактура",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_upload_main"
    )

    if uploaded_pdfs:

        with st.spinner(
            "Зареждане на Cross References..."
        ):

            cross_refs = load_cross_references(
                selected_vendor_no
            )

        st.success(
            f"Cross References: {len(cross_refs)}"
        )

        st.dataframe(
            cross_refs.head(50),
            use_container_width=True
        )

        st.session_state[
            "cross_refs"
        ] = cross_refs
        # ======================================================
# PDF PARSER
# ======================================================

def normalize_text(text):

    if text is None:
        return ""

    return re.sub(
        r"[^A-Z0-9]",
        "",
        str(text).upper()
    )


results = []

for pdf_file in uploaded_pdfs:

    reader = pdfplumber.open(pdf_file)

    extracted_text = ""

    for page in reader.pages:

        page_text = page.extract_text()

        if page_text:

            extracted_text += page_text + "\n"

    reader.close()

    lines = extracted_text.split("\n")

    for line in lines:

        normalized_line = normalize_text(
            line
        )

        if len(normalized_line) < 4:

            continue

        # =====================================
        # SEARCH CROSS REFERENCE
        # =====================================

        match_found = False

        for _, ref_row in cross_refs.iterrows():

            cross_ref = str(
                ref_row["normalized_cross_reference"]
            )

            item_no = str(
                ref_row["normalized_item_no"]
            )

            # direct cross reference

            if (
                cross_ref != ""
                and
                cross_ref in normalized_line
            ):

                results.append({
                    "Cross-Reference Type No.":
                        ref_row["vendor_no"],

                    "Cross-Reference No.":
                        ref_row["cross_reference_no"],

                    "Qty": "",

                    "Price 1 pc": ""
                })

                match_found = True

                break

            # item no fallback

            if (
                item_no != ""
                and
                item_no in normalized_line
            ):

                results.append({
                    "Cross-Reference Type No.":
                        ref_row["vendor_no"],

                    "Cross-Reference No.":
                        "⚠️ "
                        +
                        str(
                            ref_row[
                                "cross_reference_no"
                            ]
                        ),

                    "Qty": "",

                    "Price 1 pc": ""
                })

                match_found = True

                break

        # =====================================
        # NOT FOUND
        # =====================================

        if (
            not match_found
            and
            len(normalized_line) > 6
        ):

            tokens = re.findall(
                r"[A-Z0-9\-\/]+",
                line.upper()
            )

            if tokens:

                token = tokens[0]

                results.append({
                    "Cross-Reference Type No.":
                        selected_vendor_no,

                    "Cross-Reference No.":
                        "❗ " + token,

                    "Qty": "",

                    "Price 1 pc": ""
                })


# ======================================================
# RESULT TABLE
# ======================================================

if results:

    result_df = pd.DataFrame(
        results
    )

    result_df = result_df.drop_duplicates()

    st.subheader(
        "📋 Разпознати артикули"
    )

    st.dataframe(
        result_df,
        use_container_width=True
    )

    # =====================================
    # EXCEL EXPORT
    # =====================================

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        result_df.to_excel(
            writer,
            index=False
        )

    output.seek(0)

    st.download_button(
        label="📥 Изтегли Excel",
        data=output,
        file_name="cross_reference_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

else:

    st.warning(
        "Няма намерени артикули."
    )
