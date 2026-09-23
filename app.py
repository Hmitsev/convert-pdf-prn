import base64
import io
import os
import re
import unicodedata
import zipfile

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


# ======================================================
# APP CONFIGURATION
# ======================================================
st.set_page_config(
    page_title="PDF to Excel & PRN Converter",
    page_icon="📄",
    layout="wide"
)

CROSS_REFERENCE_ZIPS = [
    "cross_ref_1.zip",
    "cross_ref_2.zip",
    "cross_ref_3.zip",
]

VENDOR_FILE_HINTS = ("cros", "vendor", "достав")


# ======================================================
# BACKGROUND AND UI
# ======================================================
def set_background(image_file):
    if not os.path.exists(image_file):
        return

    try:
        with open(image_file, "rb") as image:
            encoded = base64.b64encode(image.read()).decode()

        st.markdown(
            f"""
            <style>
            .stApp {{
                background-image:
                    linear-gradient(rgba(0,0,0,0.35), rgba(0,0,0,0.45)),
                    url("data:image/png;base64,{encoded}");
                background-size: cover;
                background-position: center;
                background-attachment: fixed;
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )
    except Exception:
        pass


def apply_ui_style():
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 2rem;
            padding-bottom: 3rem;
        }

        .app-card {
            padding: 18px 20px;
            margin-bottom: 18px;
            border-radius: 16px;
            border: 1px solid rgba(255,255,255,0.20);
            background: rgba(0,0,0,0.38);
            backdrop-filter: blur(12px);
            color: white;
        }

        .app-title {
            font-size: 28px;
            font-weight: 900;
            color: #ff9700;
            margin-bottom: 5px;
        }

        .app-subtitle {
            font-size: 15px;
            line-height: 1.6;
            color: rgba(255,255,255,0.86);
        }

        section[data-testid="stSidebar"] > div {
            background: rgba(0,0,0,0.55);
            backdrop-filter: blur(14px);
            border-right: 1px solid rgba(255,255,255,0.18);
        }

        div[data-testid="stMetric"] {
            background: rgba(0,0,0,0.38);
            border: 1px solid rgba(255,255,255,0.14);
            padding: 12px;
            border-radius: 12px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


set_background("background.png")
apply_ui_style()


# ======================================================
# COMMON HELPERS
# ======================================================
def clean_value(value):
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    text = str(value).strip()

    if text.lower() in {"", "nan", "none", "null"}:
        return ""

    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]

    return text


def normalize_key(value):
    text = clean_value(value).upper()
    text = unicodedata.normalize("NFKC", text)
    return re.sub(r"[^A-ZА-Я0-9]", "", text)


def normalize_header(value):
    text = clean_value(value).lower()
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"[^a-zа-я0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def find_column(columns, aliases):
    normalized_columns = {
        column: normalize_header(column)
        for column in columns
    }

    normalized_aliases = [
        normalize_header(alias)
        for alias in aliases
    ]

    for column, header in normalized_columns.items():
        if header in normalized_aliases:
            return column

    for column, header in normalized_columns.items():
        if any(alias and alias in header for alias in normalized_aliases):
            return column

    return None


def read_all_excel_sheets(file_or_buffer, engine="openpyxl"):
    sheets = []

    try:
        workbook = pd.ExcelFile(file_or_buffer, engine=engine)

        for sheet_name in workbook.sheet_names:
            try:
                dataframe = pd.read_excel(
                    workbook,
                    sheet_name=sheet_name,
                    dtype=str,
                )

                if not dataframe.empty:
                    sheets.append((sheet_name, dataframe))

            except Exception:
                continue

    except Exception:
        return []

    return sheets


# ======================================================
# VENDOR DIRECTORY
# ======================================================
@st.cache_data(show_spinner=False)
def load_vendor_mapping():
    diagnostics = []
    vendor_rows = []
    candidate_files = []

    for file_name in os.listdir("."):
        lower_name = file_name.lower()

        if lower_name.endswith((".xlsx", ".xls")) and any(
            hint in lower_name
            for hint in VENDOR_FILE_HINTS
        ):
            candidate_files.append(file_name)

    if not candidate_files:
        diagnostics.append(
            "Не е намерен vendor Excel файл. Името трябва да съдържа "
            "cros, vendor или достав."
        )

    for file_name in candidate_files:
        engine = "xlrd" if file_name.lower().endswith(".xls") else "openpyxl"

        for sheet_name, dataframe in read_all_excel_sheets(
            file_name,
            engine=engine,
        ):
            vendor_number_column = find_column(
                dataframe.columns,
                [
                    "Buy-from Vendor No.",
                    "Buy from Vendor No",
                    "Vendor No.",
                    "Vendor No",
                ],
            )

            vendor_name_column = find_column(
                dataframe.columns,
                [
                    "Buy-from Vendor Name",
                    "Buy from Vendor Name",
                    "Vendor Name",
                    "Supplier Name",
                ],
            )

            if vendor_number_column is None or vendor_name_column is None:
                continue

            for _, row in dataframe.iterrows():
                vendor_number = clean_value(row.get(vendor_number_column, ""))
                vendor_name = clean_value(row.get(vendor_name_column, ""))

                if not vendor_number:
                    continue

                vendor_rows.append(
                    {
                        "Vendor No.": vendor_number,
                        "Vendor Name": vendor_name,
                        "Source": f"{file_name} / {sheet_name}",
                    }
                )

    if not vendor_rows:
        diagnostics.append(
            "Vendor файлът не съдържа разпознаваеми колони "
            "Buy-from Vendor No. и Buy-from Vendor Name."
        )
        return pd.DataFrame(), diagnostics

    vendors = pd.DataFrame(vendor_rows)
    vendors = vendors.drop_duplicates(subset=["Vendor No."])
    vendors = vendors.sort_values(["Vendor Name", "Vendor No."])
    vendors = vendors.reset_index(drop=True)

    return vendors, diagnostics


# ======================================================
# CROSS-REFERENCE DATABASE FROM THREE ZIP FILES
# ======================================================
@st.cache_data(show_spinner=False)
def load_vendor_cross_references(vendor_number):
    diagnostics = []
    records = []
    wanted_vendor = normalize_key(vendor_number)

    type_aliases = [
        "Cross-Reference Type No.",
        "Cross Reference Type No",
        "Cross-Reference Type Number",
        "Cross Reference Type Number",
    ]

    cross_reference_aliases = [
        "Cross-Reference No.",
        "Cross Reference No",
        "Cross-Reference Number",
        "Cross Reference Number",
    ]

    item_aliases = [
        "Item No.",
        "Item No",
        "Item Number",
        "Article No.",
        "Article Number",
        "Product No.",
        "Product Number",
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
                                dataframes = [
                                    (
                                        "CSV",
                                        pd.read_csv(
                                            buffer,
                                            dtype=str,
                                            sep=None,
                                            engine="python",
                                            encoding="utf-8-sig",
                                        ),
                                    )
                                ]
                            except UnicodeDecodeError:
                                buffer.seek(0)
                                dataframes = [
                                    (
                                        "CSV",
                                        pd.read_csv(
                                            buffer,
                                            dtype=str,
                                            sep=None,
                                            engine="python",
                                            encoding="cp1251",
                                        ),
                                    )
                                ]
                        else:
                            engine = "xlrd" if extension == ".xls" else "openpyxl"
                            dataframes = read_all_excel_sheets(
                                buffer,
                                engine=engine,
                            )

                        for sheet_name, dataframe in dataframes:
                            if dataframe.empty:
                                continue

                            type_column = find_column(
                                dataframe.columns,
                                type_aliases,
                            )

                            cross_reference_column = find_column(
                                dataframe.columns,
                                cross_reference_aliases,
                            )

                            item_column = find_column(
                                dataframe.columns,
                                item_aliases,
                            )

                            if type_column is None or cross_reference_column is None:
                                continue

                            for _, row in dataframe.iterrows():
                                type_number = clean_value(row.get(type_column, ""))

                                if normalize_key(type_number) != wanted_vendor:
                                    continue

                                cross_reference_number = clean_value(
                                    row.get(cross_reference_column, "")
                                )

                                item_number = (
                                    clean_value(row.get(item_column, ""))
                                    if item_column is not None
                                    else ""
                                )

                                if not cross_reference_number:
                                    continue

                                records.append(
                                    {
                                        "Cross-Reference Type No.": type_number,
                                        "Cross-Reference No.": cross_reference_number,
                                        "Item No.": item_number,
                                        "Source": (
                                            f"{zip_path} / {member_name} / {sheet_name}"
                                        ),
                                    }
                                )

                    except Exception as error:
                        diagnostics.append(
                            f"Неуспешно четене на {zip_path} / "
                            f"{member_name}: {error}"
                        )

        except Exception as error:
            diagnostics.append(f"Грешка в {zip_path}: {error}")

    if not records:
        diagnostics.append(
            f"В трите ZIP архива няма Cross Reference редове "
            f"за vendor {vendor_number}."
        )
        return pd.DataFrame(), diagnostics

    references = pd.DataFrame(records)
    references = references.drop_duplicates(
        subset=[
            "Cross-Reference Type No.",
            "Cross-Reference No.",
            "Item No.",
        ]
    )
    references = references.reset_index(drop=True)

    return references, diagnostics


def build_reference_indexes(reference_dataframe):
    direct_index = {}
    alias_index = {}

    for _, row in reference_dataframe.iterrows():
        vendor_number = clean_value(row.get("Cross-Reference Type No.", ""))
        cross_reference = clean_value(row.get("Cross-Reference No.", ""))
        item_number = clean_value(row.get("Item No.", ""))

        cross_key = normalize_key(cross_reference)
        item_key = normalize_key(item_number)

        if cross_key:
            direct_index[cross_key] = {
                "vendor_number": vendor_number,
                "cross_reference": cross_reference,
            }

        if item_key:
            alias_index[item_key] = {
                "vendor_number": vendor_number,
                "cross_reference": cross_reference,
            }

    return direct_index, alias_index


# ======================================================
# PDF ROW RECOGNITION
# ======================================================
def parse_number(value):
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


def prepare_known_references(direct_index, alias_index):
    known_references = []

    for key, data in direct_index.items():
        known_references.append((key, data, "direct"))

    for key, data in alias_index.items():
        known_references.append((key, data, "alias"))

    known_references.sort(
        key=lambda entry: len(entry[0]),
        reverse=True,
    )

    return known_references


def find_reference_in_row(row, known_references):
    for column_index, cell in enumerate(row):
        original_value = clean_value(cell)
        normalized_cell = normalize_key(original_value)

        if not normalized_cell:
            continue

        for known_key, data, status in known_references:
            if normalized_cell == known_key or (
                len(known_key) >= 5
                and known_key in normalized_cell
            ):
                return {
                    "vendor_number": data["vendor_number"],
                    "cross_reference": data["cross_reference"],
                    "status": status,
                    "source_value": original_value,
                    "column_index": column_index,
                }

    return None


def find_unknown_item(row, vendor_number):
    blocked_words = {
        "TOTAL",
        "SUBTOTAL",
        "GESAMT",
        "PRICE",
        "PREIS",
        "QUANTITY",
        "MENGE",
        "INVOICE",
        "RECHNUNG",
    }

    candidates = []

    for column_index, cell in enumerate(row):
        original_value = clean_value(cell)
        key = normalize_key(original_value)

        if len(key) < 5 or len(key) > 40:
            continue

        if any(word in original_value.upper() for word in blocked_words):
            continue

        if re.fullmatch(r"\d+[,.]\d{1,6}", original_value):
            continue

        if re.fullmatch(
            r"\d{1,2}[./-]\d{1,2}[./-]\d{2,4}",
            original_value,
        ):
            continue

        if key.isdigit() and len(key) < 6:
            continue

        candidates.append(
            (len(key), column_index, original_value)
        )

    if not candidates:
        return None

    candidates.sort(reverse=True)
    _, column_index, original_value = candidates[0]

    return {
        "vendor_number": vendor_number,
        "cross_reference": f"❗ {original_value}",
        "status": "not_found",
        "source_value": original_value,
        "column_index": column_index,
    }


def detect_quantity_price_total(row, item_column_index):
    numeric_values = []

    for index, cell in enumerate(row):
        if index == item_column_index:
            continue

        number = parse_number(cell)

        if number is not None:
            numeric_values.append((index, number))

    best_result = None
    best_difference = None

    for quantity_index, quantity in numeric_values:
        if quantity <= 0:
            continue

        if abs(quantity - round(quantity)) > 0.0001:
            continue

        for price_index, price in numeric_values:
            if price_index == quantity_index or price < 0:
                continue

            for total_index, line_total in numeric_values:
                if total_index in {quantity_index, price_index}:
                    continue

                difference = abs(quantity * price - line_total)
                tolerance = max(0.05, abs(line_total) * 0.01)

                if difference <= tolerance and (
                    best_difference is None
                    or difference < best_difference
                ):
                    best_result = (
                        quantity,
                        price,
                        line_total,
                    )
                    best_difference = difference

    if best_result is not None:
        return best_result

    integer_values = [
        (index, value)
        for index, value in numeric_values
        if value > 0
        and abs(value - round(value)) < 0.0001
    ]

    if not integer_values:
        return None, None, None

    quantity_index, quantity = integer_values[0]

    prices_after_quantity = [
        (index, value)
        for index, value in numeric_values
        if index > quantity_index
        and value >= 0
    ]

    if not prices_after_quantity:
        return None, None, None

    _, price = prices_after_quantity[0]

    return quantity, price, None


def extract_tables_from_page(page):
    settings_list = [
        {
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
            "intersection_tolerance": 8,
            "snap_tolerance": 5,
            "join_tolerance": 5,
        },
        {
            "vertical_strategy": "text",
            "horizontal_strategy": "text",
            "intersection_tolerance": 8,
            "snap_tolerance": 5,
            "join_tolerance": 5,
            "min_words_vertical": 1,
            "min_words_horizontal": 1,
        },
    ]

    for settings in settings_list:
        try:
            tables = page.extract_tables(settings) or []
        except Exception:
            tables = []

        if tables:
            return tables

    return []


# ======================================================
# PDF TO EXCEL CONVERSION
# ======================================================
def convert_pdf_to_excel(pdf_files, vendor_number):
    reference_dataframe, diagnostics = load_vendor_cross_references(
        vendor_number
    )

    if reference_dataframe.empty:
        return {
            "output": None,
            "preview": pd.DataFrame(),
            "diagnostics": diagnostics,
            "pages": 0,
            "tables": 0,
            "direct": 0,
            "alias": 0,
            "not_found": 0,
        }

    direct_index, alias_index = build_reference_indexes(
        reference_dataframe
    )

    known_references = prepare_known_references(
        direct_index,
        alias_index,
    )

    recognized_rows = []
    processed_pages = 0
    detected_tables = 0

    for pdf_file in pdf_files:
        try:
            pdf_file.seek(0)

            with pdfplumber.open(pdf_file) as pdf:
                for page_number, page in enumerate(
                    pdf.pages,
                    start=1,
                ):
                    processed_pages += 1
                    tables = extract_tables_from_page(page)

                    for table_number, table in enumerate(
                        tables,
                        start=1,
                    ):
                        if not table:
                            continue

                        detected_tables += 1

                        for row_number, row in enumerate(
                            table,
                            start=1,
                        ):
                            clean_row = [
                                clean_value(cell)
                                for cell in row
                            ]

                            if not any(clean_row):
                                continue

                            row_text = " ".join(clean_row).upper()

                            if any(
                                blocked in row_text
                                for blocked in [
                                    "GRAND TOTAL",
                                    "INVOICE TOTAL",
                                    "SUBTOTAL",
                                    "ZWISCHENSUMME",
                                    "GESAMTBETRAG",
                                ]
                            ):
                                continue

                            reference = find_reference_in_row(
                                clean_row,
                                known_references,
                            )

                            if reference is None:
                                reference = find_unknown_item(
                                    clean_row,
                                    vendor_number,
                                )

                            if reference is None:
                                continue

                            quantity, price, line_total = (
                                detect_quantity_price_total(
                                    clean_row,
                                    reference["column_index"],
                                )
                            )

                            if quantity is None or price is None:
                                continue

                            if quantity <= 0 or price < 0:
                                continue

                            display_cross_reference = reference[
                                "cross_reference"
                            ]

                            if reference["status"] == "alias":
                                display_cross_reference = (
                                    f"⚠️ {display_cross_reference}"
                                )

                            recognized_rows.append(
                                {
                                    "Cross-Reference Type No.": reference[
                                        "vendor_number"
                                    ],
                                    "Cross-Reference No.": (
                                        display_cross_reference
                                    ),
                                    "Qty": quantity,
                                    "Price 1 pc": price,
                                    "_status": reference["status"],
                                    "_source_value": reference[
                                        "source_value"
                                    ],
                                    "_line_total": line_total,
                                    "_file": pdf_file.name,
                                    "_page": page_number,
                                    "_table": table_number,
                                    "_row": row_number,
                                }
                            )

        except Exception as error:
            diagnostics.append(
                f"{pdf_file.name}: {error}"
            )

    if not recognized_rows:
        return {
            "output": None,
            "preview": pd.DataFrame(),
            "diagnostics": diagnostics,
            "pages": processed_pages,
            "tables": detected_tables,
            "direct": 0,
            "alias": 0,
            "not_found": 0,
        }

    full_dataframe = pd.DataFrame(recognized_rows)
    full_dataframe = full_dataframe.drop_duplicates(
        subset=[
            "Cross-Reference Type No.",
            "Cross-Reference No.",
            "Qty",
            "Price 1 pc",
            "_file",
            "_page",
            "_row",
        ]
    ).reset_index(drop=True)

    preview_dataframe = full_dataframe[
        [
            "Cross-Reference Type No.",
            "Cross-Reference No.",
            "Qty",
            "Price 1 pc",
        ]
    ].copy()

    grand_total = (
        pd.to_numeric(
            preview_dataframe["Qty"],
            errors="coerce",
        )
        * pd.to_numeric(
            preview_dataframe["Price 1 pc"],
            errors="coerce",
        )
    ).sum()

    total_row = pd.DataFrame(
        [
            {
                "Cross-Reference Type No.": "",
                "Cross-Reference No.": "TOTAL",
                "Qty": "",
                "Price 1 pc": grand_total,
            }
        ]
    )

    export_dataframe = pd.concat(
        [preview_dataframe, total_row],
        ignore_index=True,
    )

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        export_dataframe.to_excel(
            writer,
            sheet_name="Invoice",
            index=False,
        )

        worksheet = writer.sheets["Invoice"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = (
            f"A1:D{len(export_dataframe) + 1}"
        )

        worksheet.column_dimensions["A"].width = 28
        worksheet.column_dimensions["B"].width = 30
        worksheet.column_dimensions["C"].width = 14
        worksheet.column_dimensions["D"].width = 18

        header_fill = PatternFill(
            fill_type="solid",
            fgColor="D71919",
        )

        header_font = Font(
            color="FFFFFF",
            bold=True,
        )

        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        last_product_row = len(preview_dataframe) + 1
        total_excel_row = len(export_dataframe) + 1

        for excel_row in range(2, last_product_row + 1):
            cross_reference_cell = worksheet.cell(
                row=excel_row,
                column=2,
            )

            cross_reference_value = str(
                cross_reference_cell.value or ""
            )

            if cross_reference_value.startswith("⚠️"):
                cross_reference_cell.font = Font(
                    bold=True,
                    color="FF8C00",
                )

            elif cross_reference_value.startswith("❗"):
                cross_reference_cell.font = Font(
                    bold=True,
                    color="FF0000",
                )

            worksheet.cell(
                row=excel_row,
                column=3,
            ).number_format = "0.###"

            worksheet.cell(
                row=excel_row,
                column=4,
            ).number_format = "0.000000"

            for column_number in range(1, 5):
                worksheet.cell(
                    row=excel_row,
                    column=column_number,
                ).border = thin_border

        total_fill = PatternFill(
            fill_type="solid",
            fgColor="FFF2CC",
        )

        for column_number in range(1, 5):
            cell = worksheet.cell(
                row=total_excel_row,
                column=column_number,
            )
            cell.fill = total_fill
            cell.font = Font(
                bold=True,
                color="C00000",
            )
            cell.border = thin_border

        worksheet.cell(
            row=total_excel_row,
            column=4,
        ).number_format = "0.00"

    output.seek(0)

    return {
        "output": output,
        "preview": preview_dataframe,
        "diagnostics": diagnostics,
        "pages": processed_pages,
        "tables": detected_tables,
        "direct": int(
            (full_dataframe["_status"] == "direct").sum()
        ),
        "alias": int(
            (full_dataframe["_status"] == "alias").sum()
        ),
        "not_found": int(
            (full_dataframe["_status"] == "not_found").sum()
        ),
    }


# ======================================================
# PRN CONVERSION
# ======================================================
def convert_excel_to_prn(uploaded_excel, manual_invoice_number=""):
    dataframe = pd.read_excel(uploaded_excel, dtype=str)
    dataframe.columns = [
        str(column).strip()
        for column in dataframe.columns
    ]

    if "Price 1 pc" not in dataframe.columns and "Price" in dataframe.columns:
        dataframe = dataframe.rename(
            columns={"Price": "Price 1 pc"}
        )

    if "Item" not in dataframe.columns and "Cross-Reference No." in dataframe.columns:
        dataframe = dataframe.rename(
            columns={"Cross-Reference No.": "Item"}
        )

    required_columns = [
        "Item",
        "Qty",
        "Price 1 pc",
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in dataframe.columns
    ]

    if missing_columns:
        raise ValueError(
            "Липсват колони: "
            + ", ".join(missing_columns)
        )

    prn_lines = []
    preview_rows = []

    for _, row in dataframe.iterrows():
        item = clean_value(row.get("Item", ""))

        item = (
            item
            .replace("⚠️", "")
            .replace("❗", "")
            .strip()
        )

        if not item or item.upper() == "TOTAL":
            continue

        quantity = parse_number(row.get("Qty", ""))
        price = parse_number(row.get("Price 1 pc", ""))

        if quantity is None or price is None:
            continue

        if quantity <= 0 or price < 0:
            continue

        quantity_integer = int(round(quantity))
        price_text = f"{price:.6f}".replace(".", ",")

        spaces_before_quantity = max(
            1,
            25 - len(item) - len(str(quantity_integer)),
        )

        prn_line = (
            item
            + (" " * spaces_before_quantity)
            + str(quantity_integer)
            + (" " * 6)
            + price_text
        )

        prn_lines.append(prn_line)
        preview_rows.append(
            {
                "Item": item,
                "Qty": quantity_integer,
                "Price 1 pc": price,
            }
        )

    if not prn_lines:
        raise ValueError(
            "Не са намерени валидни редове за PRN."
        )

    invoice_number = clean_value(manual_invoice_number)

    if not invoice_number and "Invoice" in dataframe.columns:
        invoice_values = dataframe["Invoice"].dropna()
        if not invoice_values.empty:
            invoice_number = clean_value(invoice_values.iloc[0])

    if not invoice_number:
        invoice_number = os.path.splitext(uploaded_excel.name)[0]

    prn_content = "\r\n".join(prn_lines)
    preview_dataframe = pd.DataFrame(preview_rows)

    return {
        "content": prn_content.encode("utf-8"),
        "file_name": f"{invoice_number}.prn",
        "preview": preview_dataframe,
        "row_count": len(prn_lines),
    }


# ======================================================
# MAIN NAVIGATION
# ======================================================
st.markdown(
    """
    <div class="app-card">
        <div class="app-title">📄 PDF / Excel / PRN Converter</div>
        <div class="app-subtitle">
            Самостоятелно приложение за PDF фактура към Excel чрез Vendor
            Cross Reference и за Excel към PRN.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

mode = st.sidebar.radio(
    "Избери функция",
    [
        "📄 PDF към Excel",
        "🧾 Excel към PRN",
    ],
)


# ======================================================
# SCREEN 1: PDF TO EXCEL
# ======================================================
if mode == "📄 PDF към Excel":
    st.subheader("📄 PDF фактура към Excel")

    vendors_dataframe, vendor_diagnostics = load_vendor_mapping()

    if vendors_dataframe.empty:
        st.error(
            "Не е намерен vendor справочник с колони "
            "Buy-from Vendor No. и Buy-from Vendor Name."
        )

        with st.expander("🔎 Диагностика"):
            for message in vendor_diagnostics:
                st.write(f"• {message}")

        st.stop()

    vendor_options = vendors_dataframe.apply(
        lambda row: (
            f"{row['Vendor No.']} | {row['Vendor Name']}"
            if row["Vendor Name"]
            else row["Vendor No."]
        ),
        axis=1,
    ).tolist()

    selected_vendor_label = st.selectbox(
        "Избери доставчик / Vendor",
        vendor_options,
    )

    selected_vendor_number = selected_vendor_label.split(
        " | ",
        1,
    )[0].strip()

    reference_preview, reference_diagnostics = (
        load_vendor_cross_references(
            selected_vendor_number
        )
    )

    if reference_preview.empty:
        st.warning(
            f"Не са намерени Cross Reference записи за "
            f"{selected_vendor_number}."
        )

        with st.expander("🔎 Cross Reference диагностика"):
            for message in reference_diagnostics:
                st.write(f"• {message}")

    else:
        st.caption(
            f"Намерени Cross Reference записи: "
            f"{len(reference_preview)}"
        )

    uploaded_pdfs = st.file_uploader(
        "Качи един или няколко PDF файла",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_invoice_uploader",
    )

    if uploaded_pdfs:
        with st.spinner(
            "Разпознаване чрез Vendor Cross Reference..."
        ):
            conversion_result = convert_pdf_to_excel(
                uploaded_pdfs,
                selected_vendor_number,
            )

        if conversion_result["output"] is None:
            st.error(
                "Не бяха разпознати редове с Cross Reference, "
                "количество и цена."
            )

            with st.expander("🔎 Диагностика"):
                for message in conversion_result["diagnostics"]:
                    st.write(f"• {message}")

            st.stop()

        metric_col1, metric_col2, metric_col3 = st.columns(3)

        with metric_col1:
            st.metric(
                "✅ Директни",
                conversion_result["direct"],
            )

        with metric_col2:
            st.metric(
                "⚠️ Чрез Item No.",
                conversion_result["alias"],
            )

        with metric_col3:
            st.metric(
                "❗ За проверка",
                conversion_result["not_found"],
            )

        st.success(
            f"Страници: {conversion_result['pages']} | "
            f"Таблици: {conversion_result['tables']}"
        )

        st.subheader("📋 Контролен преглед")
        st.dataframe(
            conversion_result["preview"],
            use_container_width=True,
        )

        st.caption(
            "⚠️ = Cross Reference е намерен чрез Item No. | "
            "❗ = номерът не е намерен за избрания vendor"
        )

        if len(uploaded_pdfs) == 1:
            excel_file_name = re.sub(
                r"\.pdf$",
                ".xlsx",
                uploaded_pdfs[0].name,
                flags=re.IGNORECASE,
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
            use_container_width=True,
        )


# ======================================================
# SCREEN 2: EXCEL TO PRN
# ======================================================
else:
    st.subheader("🧾 Excel към PRN")

    uploaded_excel = st.file_uploader(
        "Качи Excel файл",
        type=["xlsx", "xls"],
        accept_multiple_files=False,
        key="prn_excel_uploader",
    )

    manual_invoice_number = st.text_input(
        "Номер на фактура, ако липсва колона Invoice",
        placeholder="Например 230280",
    )

    if uploaded_excel is not None:
        try:
            prn_result = convert_excel_to_prn(
                uploaded_excel,
                manual_invoice_number,
            )

            st.success(
                f"Готови PRN редове: {prn_result['row_count']}"
            )

            st.dataframe(
                prn_result["preview"],
                use_container_width=True,
            )

            st.download_button(
                label=f"⬇️ Изтегли {prn_result['file_name']}",
                data=prn_result["content"],
                file_name=prn_result["file_name"],
                mime="text/plain",
                use_container_width=True,
            )

        except Exception as error:
            st.error(f"Грешка: {error}")
