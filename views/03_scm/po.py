"""
CLX ERP - Purchase Order Module
File: views/03_scm/po.py

Purchase Order:
    Material
        0001/CLX/PO/RM/VIII/2026

    Non Material / General
        0001/CLX/PO/VIII/2026

Features:
- Create PO Material
- Create PO Non Material / General
- Auto PO numbering
- Google Sheets persistence
- Search PO
- Edit PO
- CEO Approval support
- PDF preview
- PDF download
- English amount in words
- Shared database connection
- Targeted Google Sheets update
- No st_gsheets_connection dependency
- No num2words dependency
"""

# ==============================================================================
# IMPORT
# ==============================================================================

import io
import os
import re
import json
import base64
from datetime import datetime

import pandas as pd
import streamlit as st

# Shared database layer
from core.database import get_google_sheet_connection


# ==============================================================================
# REPORTLAB
# ==============================================================================

try:

    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.enums import (
        TA_CENTER,
        TA_LEFT,
        TA_RIGHT,
    )
    from reportlab.lib.styles import (
        getSampleStyleSheet,
        ParagraphStyle,
    )
    from reportlab.lib.units import mm

    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
        Image as RLImage,
    )

    REPORTLAB_OK = True

except ImportError:

    REPORTLAB_OK = False


# ==============================================================================
# CONFIGURATION
# ==============================================================================

MATERIAL_SHEET = "DB Material IN"
GENERAL_SHEET = "DB General IN"

COUNTER_SHEET = "PO Counter"


MONTH_ROMAN = {
    1: "I",
    2: "II",
    3: "III",
    4: "IV",
    5: "V",
    6: "VI",
    7: "VII",
    8: "VIII",
    9: "IX",
    10: "X",
    11: "XI",
    12: "XII",
}


COMMON_HEADERS = [
    "PO No",
    "PO Type",
    "PO Date",
    "Supplier",
    "Supplier Phone",
    "Supplier Address",
    "Project",
    "Payment Terms",
    "Subtotal",
    "Tax Rate",
    "Tax Amount",
    "Grand Total",
    "Amount In Words",
    "Approval Status",
    "Approval By",
    "Approval Date",
    "Approval Notes",
    "Created By",
    "Created At",
    "Updated At",
    "Items JSON",
]


MATERIAL_ITEM_HEADERS = [
    "Item No",
    "Charger Type",
    "Site Name",
    "Description",
    "Qty",
    "UoM",
    "Unit Price",
    "Total Amount",
]


GENERAL_ITEM_HEADERS = [
    "Item No",
    "Description",
    "Site / Location",
    "Qty",
    "UoM",
    "Unit Price",
    "Total Amount",
]


# ==============================================================================
# BASIC UTILITY
# ==============================================================================

def _safe_float(value, default=0.0):
    """
    Konversi nilai menjadi float dengan aman.
    """

    try:

        if value is None:
            return default

        text = str(value).strip()

        if text == "":
            return default

        text = (
            text
            .replace("Rp", "")
            .replace("rp", "")
            .replace(",", "")
            .strip()
        )

        return float(text)

    except Exception:

        return default


def rupiah(value):
    """
    Format angka menjadi Rupiah.
    """

    return f"Rp{_safe_float(value):,.0f}"


def format_date(value):
    """
    Format tanggal ke DD/MM/YYYY.
    """

    if value is None:
        return ""

    if str(value).strip() == "":
        return ""

    try:

        return pd.to_datetime(value).strftime(
            "%d/%m/%Y"
        )

    except Exception:

        return str(value)


def get_current_period():
    """
    Mengembalikan:
        year
        month
        roman month
    """

    now = datetime.now()

    return (
        now.year,
        now.month,
        MONTH_ROMAN[now.month],
    )


# ==============================================================================
# ENGLISH NUMBER TO WORDS
# Tidak membutuhkan package num2words
# ==============================================================================

ONES = [
    "Zero",
    "One",
    "Two",
    "Three",
    "Four",
    "Five",
    "Six",
    "Seven",
    "Eight",
    "Nine",
    "Ten",
    "Eleven",
    "Twelve",
    "Thirteen",
    "Fourteen",
    "Fifteen",
    "Sixteen",
    "Seventeen",
    "Eighteen",
    "Nineteen",
]


TENS = [
    "",
    "",
    "Twenty",
    "Thirty",
    "Forty",
    "Fifty",
    "Sixty",
    "Seventy",
    "Eighty",
    "Ninety",
]


def _number_to_words_under_1000(number):
    """
    Convert 0 - 999 ke English words.
    """

    number = int(number)

    if number < 20:

        return ONES[number]

    if number < 100:

        tens = number // 10
        remainder = number % 10

        if remainder == 0:
            return TENS[tens]

        return (
            f"{TENS[tens]} "
            f"{ONES[remainder]}"
        )

    hundreds = number // 100
    remainder = number % 100

    result = (
        f"{ONES[hundreds]} Hundred"
    )

    if remainder:

        result += (
            f" {_number_to_words_under_1000(remainder)}"
        )

    return result


def english_number_words(number):
    """
    Convert integer sampai triliun ke English words.
    """

    number = int(number)

    if number == 0:
        return "Zero"

    if number < 0:

        return (
            "Minus "
            + english_number_words(abs(number))
        )

    scales = [
        (1_000_000_000_000, "Trillion"),
        (1_000_000_000, "Billion"),
        (1_000_000, "Million"),
        (1_000, "Thousand"),
    ]

    parts = []
    remainder = number

    for scale_value, scale_name in scales:

        if remainder >= scale_value:

            quotient = remainder // scale_value

            parts.append(
                _number_to_words_under_1000(
                    quotient
                )
            )

            parts.append(scale_name)

            remainder %= scale_value

    if remainder:

        parts.append(
            _number_to_words_under_1000(
                remainder
            )
        )

    return " ".join(parts)


def english_amount_words(amount):
    """
    Contoh:

    1998000000
    ->
    One Billion Nine Hundred Ninety Eight Million
    One Hundred Eighty Nine Thousand Rupiah
    """

    amount = int(
        round(
            _safe_float(amount)
        )
    )

    return (
        english_number_words(amount)
        + " Rupiah"
    )


# ==============================================================================
# GOOGLE SHEETS CONNECTION
# ==============================================================================

def get_sheet_client():
    """
    Menggunakan shared connection dari core.database.py.

    Tidak lagi menggunakan:
        st_gsheets_connection
    """

    return get_google_sheet_connection()


def get_worksheet(sheet_name):
    """
    Mengambil worksheet berdasarkan nama.
    """

    sh = get_sheet_client()

    return sh.worksheet(sheet_name)


# ==============================================================================
# GOOGLE SHEETS READ
# ==============================================================================

def read_sheet(sheet_name):
    """
    Membaca worksheet menjadi DataFrame.

    Menggunakan get_all_values() supaya:
    - aman terhadap merged cell
    - tidak tergantung header unik
    - mudah menentukan nomor row asli
    """

    try:

        worksheet = get_worksheet(
            sheet_name
        )

        values = worksheet.get_all_values()

        if not values:

            return pd.DataFrame()

        if len(values) == 1:

            headers = values[0]

            return pd.DataFrame(
                columns=headers
            )

        headers = values[0]

        data = values[1:]

        # Pastikan panjang setiap row sama
        normalized = []

        for row in data:

            row = list(row)

            if len(row) < len(headers):

                row += [
                    ""
                ] * (
                    len(headers)
                    - len(row)
                )

            elif len(row) > len(headers):

                row = row[
                    :len(headers)
                ]

            normalized.append(row)

        return pd.DataFrame(
            normalized,
            columns=headers
        ).fillna("")

    except Exception as e:

        st.error(
            f"Gagal membaca sheet "
            f"`{sheet_name}`: {e}"
        )

        return pd.DataFrame()


# ==============================================================================
# SHEET HEADER
# ==============================================================================

def ensure_sheet_headers(
    sheet_name,
    required_headers
):
    """
    Memastikan semua header tersedia.

    Tidak melakukan clear worksheet.
    """

    worksheet = get_worksheet(
        sheet_name
    )

    values = worksheet.get_all_values()

    if not values:

        worksheet.update(
            "A1",
            [required_headers]
        )

        return worksheet, required_headers

    existing_headers = list(
        values[0]
    )

    changed = False

    for header in required_headers:

        if header not in existing_headers:

            existing_headers.append(
                header
            )

            changed = True

    if changed:

        worksheet.update(
            "A1",
            [existing_headers]
        )

    return (
        worksheet,
        existing_headers
    )


# ==============================================================================
# SHEET WRITE - TARGETED
# ==============================================================================

def append_record(
    sheet_name,
    record
):
    """
    Menambahkan 1 record baru.

    Tidak rewrite seluruh worksheet.
    """

    worksheet, headers = (
        ensure_sheet_headers(
            sheet_name,
            COMMON_HEADERS
        )
    )

    row = [
        record.get(header, "")
        for header in headers
    ]

    worksheet.append_row(
        row,
        value_input_option="USER_ENTERED"
    )

    return True


def update_record(
    sheet_name,
    row_number,
    record
):
    """
    Update hanya 1 baris.

    Tidak menggunakan worksheet.clear().
    """

    worksheet, headers = (
        ensure_sheet_headers(
            sheet_name,
            COMMON_HEADERS
        )
    )

    row = [
        record.get(header, "")
        for header in headers
    ]

    end_col = _column_letter(
        len(headers)
    )

    cell_range = (
        f"A{row_number}:"
        f"{end_col}{row_number}"
    )

    worksheet.update(
        cell_range,
        [row],
        value_input_option="USER_ENTERED"
    )

    return True


def _column_letter(number):
    """
    1 -> A
    26 -> Z
    27 -> AA
    """

    result = ""

    while number:

        number, remainder = divmod(
            number - 1,
            26
        )

        result = (
            chr(
                65 + remainder
            )
            + result
        )

    return result


# ==============================================================================
# TARGET SHEET
# ==============================================================================

def target_sheet(po_type):

    if po_type == "Material":

        return MATERIAL_SHEET

    return GENERAL_SHEET


# ==============================================================================
# PO NUMBER
# ==============================================================================

def _extract_po_sequence(po_no):

    match = re.match(
        r"^\s*(\d+)/CLX/PO(?:/RM)?/",
        str(po_no).upper()
    )

    if not match:

        return 0

    try:

        return int(
            match.group(1)
        )

    except Exception:

        return 0


def _is_current_period_po(
    po_no,
    po_type
):
    """
    Mengecek apakah PO berada pada
    bulan + tahun berjalan.
    """

    year, month, roman = (
        get_current_period()
    )

    text = str(
        po_no
    ).strip().upper()

    if po_type == "Material":

        expected = (
            f"/CLX/PO/RM/"
            f"{roman}/{year}"
        )

    else:

        expected = (
            f"/CLX/PO/"
            f"{roman}/{year}"
        )

    return expected in text


def get_next_po_number(po_type):
    """
    Generate nomor PO.

    Sequence reset berdasarkan periode:
        Material:
        0001/CLX/PO/RM/VIII/2026

        Non Material:
        0001/CLX/PO/VIII/2026

    PO bulan sebelumnya tidak ikut
    menaikkan sequence bulan berjalan.
    """

    sheet = target_sheet(
        po_type
    )

    df = read_sheet(
        sheet
    )

    sequences = []

    if (
        not df.empty
        and "PO No" in df.columns
    ):

        for value in df["PO No"]:

            if _is_current_period_po(
                value,
                po_type
            ):

                sequences.append(
                    _extract_po_sequence(
                        value
                    )
                )

    highest = max(
        sequences or [0]
    )

    year, month, roman = (
        get_current_period()
    )

    next_number = (
        highest + 1
    )

    if po_type == "Material":

        return (
            f"{next_number:04d}"
            f"/CLX/PO/RM/"
            f"{roman}/{year}"
        )

    return (
        f"{next_number:04d}"
        f"/CLX/PO/"
        f"{roman}/{year}"
    )


def validate_unique_po(
    po_no
):
    """
    Validasi PO No pada kedua database.
    """

    for sheet_name in [
        MATERIAL_SHEET,
        GENERAL_SHEET,
    ]:

        df = read_sheet(
            sheet_name
        )

        if (
            not df.empty
            and "PO No" in df.columns
        ):

            exists = (
                df["PO No"]
                .astype(str)
                .str.strip()
                .eq(str(po_no).strip())
                .any()
            )

            if exists:

                return False

    return True


# ==============================================================================
# ITEM CALCULATION
# ==============================================================================

def calculate_items(items):

    normalized = []

    for idx, item in enumerate(
        items,
        start=1
    ):

        qty = _safe_float(
            item.get(
                "Qty",
                0
            )
        )

        price = _safe_float(
            item.get(
                "Unit Price",
                0
            )
        )

        total = (
            qty * price
        )

        row = dict(item)

        row["Item No"] = idx

        row["Qty"] = qty

        row["Unit Price"] = price

        row["Total Amount"] = total

        normalized.append(
            row
        )

    return normalized


def calculate_totals(
    items,
    tax_rate
):

    subtotal = sum(
        _safe_float(
            item.get(
                "Total Amount",
                0
            )
        )
        for item in items
    )

    tax_amount = (
        subtotal
        * (
            _safe_float(
                tax_rate
            )
            / 100
        )
    )

    grand_total = (
        subtotal
        + tax_amount
    )

    return (
        subtotal,
        tax_amount,
        grand_total
    )


# ==============================================================================
# ITEMS JSON
# ==============================================================================

def _items_to_json(items):

    return json.dumps(
        items,
        ensure_ascii=False
    )


def _items_from_json(value):

    if isinstance(
        value,
        list
    ):

        return value

    try:

        return json.loads(
            str(value)
        )

    except Exception:

        return []


# ==============================================================================
# BUILD RECORD
# ==============================================================================

def build_record(
    po_data,
    items
):

    items = calculate_items(
        items
    )

    subtotal, tax_amount, grand_total = (
        calculate_totals(
            items,
            po_data.get(
                "Tax Rate",
                11
            )
        )
    )

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    record = {

        "PO No":
            po_data["PO No"],

        "PO Type":
            po_data["PO Type"],

        "PO Date":
            po_data["PO Date"],

        "Supplier":
            po_data.get(
                "Supplier",
                ""
            ),

        "Supplier Phone":
            po_data.get(
                "Supplier Phone",
                ""
            ),

        "Supplier Address":
            po_data.get(
                "Supplier Address",
                ""
            ),

        "Project":
            po_data.get(
                "Project",
                ""
            ),

        "Payment Terms":
            po_data.get(
                "Payment Terms",
                ""
            ),

        "Subtotal":
            subtotal,

        "Tax Rate":
            _safe_float(
                po_data.get(
                    "Tax Rate",
                    11
                )
            ),

        "Tax Amount":
            tax_amount,

        "Grand Total":
            grand_total,

        "Amount In Words":
            english_amount_words(
                grand_total
            ),

        "Approval Status":
            po_data.get(
                "Approval Status",
                "Pending"
            ),

        "Approval By":
            po_data.get(
                "Approval By",
                ""
            ),

        "Approval Date":
            po_data.get(
                "Approval Date",
                ""
            ),

        "Approval Notes":
            po_data.get(
                "Approval Notes",
                ""
            ),

        "Created By":
            po_data.get(
                "Created By",
                ""
            ),

        "Created At":
            po_data.get(
                "Created At",
                ""
            )
            or now,

        "Updated At":
            now,

        "Items JSON":
            _items_to_json(
                items
            ),
    }

    return (
        record,
        items
    )


# ==============================================================================
# FIND PO ROW
# ==============================================================================

def find_po_row(
    sheet_name,
    po_no
):
    """
    Mengembalikan nomor baris Google Sheet
    jika PO ditemukan.

    Row Google Sheet:
        Header = row 1
        Data pertama = row 2
    """

    try:

        worksheet = get_worksheet(
            sheet_name
        )

        values = worksheet.get_all_values()

        if not values:

            return None

        headers = values[0]

        if "PO No" not in headers:

            return None

        po_col_index = (
            headers.index(
                "PO No"
            )
        )

        for index, row in enumerate(
            values[1:],
            start=2
        ):

            if len(row) <= po_col_index:

                continue

            value = str(
                row[po_col_index]
            ).strip()

            if value == str(
                po_no
            ).strip():

                return index

        return None

    except Exception:

        return None


# ==============================================================================
# SAVE PO
# ==============================================================================

def save_po(
    po_data,
    items,
    edit_mode=False
):

    po_type = po_data[
        "PO Type"
    ]

    sheet = target_sheet(
        po_type
    )

    record, normalized_items = (
        build_record(
            po_data,
            items
        )
    )

    po_no = str(
        record["PO No"]
    )

    # --------------------------------------------------------------------------
    # CREATE
    # --------------------------------------------------------------------------

    if not edit_mode:

        # Cek nomor
        if not validate_unique_po(
            po_no
        ):

            po_no = get_next_po_number(
                po_type
            )

            record["PO No"] = (
                po_no
            )

            # Double check
            if not validate_unique_po(
                po_no
            ):

                raise RuntimeError(
                    "Nomor PO sedang digunakan "
                    "oleh user lain. "
                    "Silakan klik Save PO kembali."
                )

        append_record(
            sheet,
            record
        )

    # --------------------------------------------------------------------------
    # EDIT
    # --------------------------------------------------------------------------

    else:

        row_number = find_po_row(
            sheet,
            po_no
        )

        if row_number is None:

            raise RuntimeError(
                f"PO {po_no} tidak ditemukan."
            )

        df = read_sheet(
            sheet
        )

        if (
            not df.empty
            and "Approval Status"
            in df.columns
        ):

            rows = df[
                df["PO No"]
                .astype(str)
                .eq(po_no)
            ]

            if not rows.empty:

                current_status = str(
                    rows.iloc[0][
                        "Approval Status"
                    ]
                ).strip().lower()

                if (
                    current_status
                    == "approved"
                ):

                    raise PermissionError(
                        "PO sudah Approved oleh CEO "
                        "dan tidak dapat diedit."
                    )

        update_record(
            sheet,
            row_number,
            record
        )

    return (
        po_no,
        record,
        normalized_items
    )


# ==============================================================================
# SEARCH
# ==============================================================================

def search_po(
    po_type,
    keyword
):

    sheet = target_sheet(
        po_type
    )

    df = read_sheet(
        sheet
    )

    if (
        df.empty
        or "PO No" not in df.columns
    ):

        return pd.DataFrame()

    keyword = str(
        keyword
    ).strip().lower()

    if not keyword:

        return df

    mask = pd.Series(
        False,
        index=df.index
    )

    search_columns = [
        "PO No",
        "Supplier",
        "Project",
        "Approval Status",
    ]

    for col in search_columns:

        if col in df.columns:

            mask |= (
                df[col]
                .astype(str)
                .str.lower()
                .str.contains(
                    keyword,
                    regex=False,
                    na=False
                )
            )

    return df[
        mask
    ].copy()


# ==============================================================================
# LOAD PO
# ==============================================================================

def load_po(
    po_type,
    po_no
):

    df = read_sheet(
        target_sheet(
            po_type
        )
    )

    if (
        df.empty
        or "PO No" not in df.columns
    ):

        return None

    rows = df[
        df["PO No"]
        .astype(str)
        .eq(str(po_no))
    ]

    if rows.empty:

        return None

    row = rows.iloc[
        0
    ].to_dict()

    items = _items_from_json(
        row.get(
            "Items JSON",
            "[]"
        )
    )

    return (
        row,
        items
    )


# ==============================================================================
# CEO APPROVAL
# ==============================================================================

def set_po_approval(
    po_type,
    po_no,
    decision,
    approver,
    notes=""
):

    allowed = {
        "Approved",
        "Reject",
        "Revisi",
    }

    if decision not in allowed:

        raise ValueError(
            "Decision harus salah satu: "
            f"{sorted(allowed)}"
        )

    sheet = target_sheet(
        po_type
    )

    row_number = find_po_row(
        sheet,
        po_no
    )

    if row_number is None:

        raise RuntimeError(
            f"PO {po_no} tidak ditemukan."
        )

    df = read_sheet(
        sheet
    )

    if (
        df.empty
        or "PO No" not in df.columns
    ):

        raise RuntimeError(
            "Data PO belum tersedia."
        )

    rows = df[
        df["PO No"]
        .astype(str)
        .eq(str(po_no))
    ]

    if rows.empty:

        raise RuntimeError(
            f"PO {po_no} tidak ditemukan."
        )

    current = str(
        rows.iloc[0].get(
            "Approval Status",
            ""
        )
    ).strip()

    if (
        current.lower()
        == "approved"
        and decision != "Approved"
    ):

        raise PermissionError(
            "PO Approved tidak dapat "
            "diubah melalui form biasa."
        )

    record = rows.iloc[
        0
    ].to_dict()

    record[
        "Approval Status"
    ] = decision

    record[
        "Approval By"
    ] = approver

    record[
        "Approval Date"
    ] = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    record[
        "Approval Notes"
    ] = notes

    record[
        "Updated At"
    ] = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    update_record(
        sheet,
        row_number,
        record
    )

    return True


# ==============================================================================
# PENDING APPROVAL
# ==============================================================================

def get_pending_approval(
    po_type=None
):

    if po_type:

        sheets = [
            target_sheet(
                po_type
            )
        ]

    else:

        sheets = [
            MATERIAL_SHEET,
            GENERAL_SHEET,
        ]

    result = []

    for sheet in sheets:

        df = read_sheet(
            sheet
        )

        if df.empty:

            continue

        if (
            "Approval Status"
            not in df.columns
        ):

            continue

        part = df[
            df["Approval Status"]
            .astype(str)
            .str.lower()
            .eq("pending")
        ].copy()

        if not part.empty:

            result.append(
                part
            )

    if not result:

        return pd.DataFrame()

    return pd.concat(
        result,
        ignore_index=True
    )


# ==============================================================================
# ASSET
# ==============================================================================

def find_asset(*names):

    base_dirs = [

        os.path.join(
            os.getcwd(),
            "asset"
        ),

        os.path.join(
            os.getcwd(),
            "assets"
        ),

        os.path.dirname(
            os.path.abspath(
                __file__
            )
        ),

        os.path.dirname(
            os.path.dirname(
                os.path.abspath(
                    __file__
                )
            )
        ),
    ]

    for base in base_dirs:

        for name in names:

            path = os.path.join(
                base,
                name
            )

            if os.path.exists(
                path
            ):

                return path

    return None


# ==============================================================================
# REPORTLAB PARAGRAPH
# ==============================================================================

def _p(
    text,
    style
):

    text = str(
        text
    )

    text = (
        text
        .replace(
            "&",
            "&amp;"
        )
        .replace(
            "<",
            "&lt;"
        )
        .replace(
            ">",
            "&gt;"
        )
    )

    return Paragraph(
        text,
        style
    )


# ==============================================================================
# PDF
# ==============================================================================

def create_po_pdf(
    po_data,
    items
):

    if not REPORTLAB_OK:

        raise RuntimeError(
            "ReportLab belum terinstall. "
            "Tambahkan `reportlab` "
            "ke requirements.txt."
        )

    items = calculate_items(
        items
    )

    subtotal, tax_amount, grand_total = (
        calculate_totals(
            items,
            po_data.get(
                "Tax Rate",
                11
            )
        )
    )

    buf = io.BytesIO()

    doc = SimpleDocTemplate(

        buf,

        pagesize=A4,

        rightMargin=10 * mm,
        leftMargin=10 * mm,
        topMargin=8 * mm,
        bottomMargin=8 * mm,

        title=(
            f"Purchase Order "
            f"{po_data['PO No']}"
        ),

        author="PT CLX",
    )

    styles = (
        getSampleStyleSheet()
    )

    normal = ParagraphStyle(
        "PO_Normal",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=10.5,
    )

    small = ParagraphStyle(
        "PO_Small",
        parent=normal,
        fontSize=7.5,
        leading=9,
    )

    title = ParagraphStyle(
        "PO_Title",
        parent=normal,
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=18,
        alignment=TA_CENTER,
    )

    right = ParagraphStyle(
        "PO_Right",
        parent=normal,
        alignment=TA_RIGHT,
    )

    bold = ParagraphStyle(
        "PO_Bold",
        parent=normal,
        fontName="Helvetica-Bold",
    )

    story = []

    # --------------------------------------------------------------------------
    # LOGO
    # --------------------------------------------------------------------------

    logo = find_asset(
        "CLX.png",
        "CLX Logo.png",
        "logo.png",
        "logo_clx.png",
        "PT CLX.png",
    )

    if logo:

        try:

            logo_img = RLImage(
                logo,
                width=48 * mm,
                height=18 * mm
            )

            logo_img.hAlign = "LEFT"

            story.append(
                logo_img
            )

        except Exception:
            pass

    story.append(
        Spacer(
            1,
            2 * mm
        )
    )

    title_text = (
        "PURCHASE ORDER MATERIAL"
        if po_data["PO Type"]
        == "Material"
        else
        "PURCHASE ORDER"
    )

    story.append(
        _p(
            title_text,
            title
        )
    )

    story.append(
        Spacer(
            1,
            5 * mm
        )
    )

    # --------------------------------------------------------------------------
    # SUPPLIER
    # --------------------------------------------------------------------------

    supplier_data = [

        [
            _p(
                "<b>Kepada YTH</b>",
                normal
            ),
            "",
            _p(
                "<b>No. PO :</b>",
                normal
            ),
            _p(
                po_data["PO No"],
                normal
            ),
        ],

        [
            _p(
                "Supplier :",
                normal
            ),
            _p(
                po_data.get(
                    "Supplier",
                    ""
                ),
                normal
            ),
            _p(
                "Date :",
                normal
            ),
            _p(
                format_date(
                    po_data.get(
                        "PO Date",
                        ""
                    )
                ),
                normal
            ),
        ],

        [
            _p(
                "No. Tlpn :",
                normal
            ),
            _p(
                po_data.get(
                    "Supplier Phone",
                    ""
                ),
                normal
            ),
            _p(
                "Project :",
                normal
            ),
            _p(
                po_data.get(
                    "Project",
                    ""
                ),
                normal
            ),
        ],

        [
            _p(
                "Address :",
                normal
            ),
            _p(
                po_data.get(
                    "Supplier Address",
                    ""
                ),
                normal
            ),
            _p(
                "Payment :",
                normal
            ),
            _p(
                po_data.get(
                    "Payment Terms",
                    ""
                ),
                normal
            ),
        ],
    ]

    supplier_table = Table(
        supplier_data,
        colWidths=[
            24 * mm,
            76 * mm,
            24 * mm,
            55 * mm,
        ],
    )

    supplier_table.setStyle(
        TableStyle([
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP"
            ),
            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                1
            ),
            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                1
            ),
            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                1.5
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                1.5
            ),
        ])
    )

    story.append(
        supplier_table
    )

    story.append(
        Spacer(
            1,
            5 * mm
        )
    )

    # --------------------------------------------------------------------------
    # ITEM TABLE
    # --------------------------------------------------------------------------

    if (
        po_data["PO Type"]
        == "Material"
    ):

        table_data = [[

            _p(
                "<b>No</b>",
                small
            ),

            _p(
                "<b>Charger Type</b>",
                small
            ),

            _p(
                "<b>Site Name</b>",
                small
            ),

            _p(
                "<b>Qty</b>",
                small
            ),

            _p(
                "<b>UoM</b>",
                small
            ),

            _p(
                "<b>Unit Price (Rp)</b>",
                small
            ),

            _p(
                "<b>Total Amount (Rp)</b>",
                small
            ),
        ]]

        for item in items:

            table_data.append([

                str(
                    item.get(
                        "Item No",
                        ""
                    )
                ),

                str(
                    item.get(
                        "Charger Type",
                        ""
                    )
                ),

                str(
                    item.get(
                        "Site Name",
                        ""
                    )
                ),

                f"{_safe_float(item.get('Qty', 0)):g}",

                str(
                    item.get(
                        "UoM",
                        ""
                    )
                ),

                f"{_safe_float(item.get('Unit Price', 0)):,.0f}",

                f"{_safe_float(item.get('Total Amount', 0)):,.0f}",
            ])

        widths = [
            9 * mm,
            28 * mm,
            58 * mm,
            16 * mm,
            16 * mm,
            30 * mm,
            34 * mm,
        ]

    else:

        table_data = [[

            _p(
                "<b>No</b>",
                small
            ),

            _p(
                "<b>Description / Service</b>",
                small
            ),

            _p(
                "<b>Site / Location</b>",
                small
            ),

            _p(
                "<b>Qty</b>",
                small
            ),

            _p(
                "<b>UoM</b>",
                small
            ),

            _p(
                "<b>Unit Price (Rp)</b>",
                small
            ),

            _p(
                "<b>Total Amount (Rp)</b>",
                small
            ),
        ]]

        for item in items:

            table_data.append([

                str(
                    item.get(
                        "Item No",
                        ""
                    )
                ),

                str(
                    item.get(
                        "Description",
                        ""
                    )
                ),

                str(
                    item.get(
                        "Site / Location",
                        ""
                    )
                ),

                f"{_safe_float(item.get('Qty', 0)):g}",

                str(
                    item.get(
                        "UoM",
                        ""
                    )
                ),

                f"{_safe_float(item.get('Unit Price', 0)):,.0f}",

                f"{_safe_float(item.get('Total Amount', 0)):,.0f}",
            ])

        widths = [
            9 * mm,
            50 * mm,
            42 * mm,
            16 * mm,
            16 * mm,
            30 * mm,
            34 * mm,
        ]

    while len(
        table_data
    ) < 17:

        table_data.append(
            [
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ]
        )

    item_table = Table(
        table_data,
        colWidths=widths,
        repeatRows=1
    )

    item_table.setStyle(
        TableStyle([

            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.black
            ),

            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.white
            ),

            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold"
            ),

            (
                "ALIGN",
                (0, 0),
                (0, -1),
                "CENTER"
            ),

            (
                "ALIGN",
                (3, 1),
                (3, -1),
                "CENTER"
            ),

            (
                "ALIGN",
                (4, 1),
                (4, -1),
                "CENTER"
            ),

            (
                "ALIGN",
                (5, 1),
                (-1, -1),
                "RIGHT"
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),

            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                2
            ),

            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                2
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                3
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                3
            ),
        ])
    )

    story.append(
        item_table
    )

    # --------------------------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------------------------

    tax_rate = _safe_float(
        po_data.get(
            "Tax Rate",
            11
        )
    )

    summary = Table(

        [

            [
                _p(
                    "<b>TOTAL AMOUNT IN WORDS</b>",
                    small
                ),
                "",
                _p(
                    "<b>Total</b>",
                    small
                ),
                _p(
                    rupiah(subtotal),
                    right
                ),
            ],

            [
                _p(
                    english_amount_words(
                        grand_total
                    ),
                    bold
                ),
                "",
                _p(
                    f"Tax {tax_rate:g}%",
                    small
                ),
                _p(
                    rupiah(tax_amount),
                    right
                ),
            ],

            [
                "",
                "",
                _p(
                    "<b>Grand Total</b>",
                    small
                ),
                _p(
                    rupiah(grand_total),
                    right
                ),
            ],
        ],

        colWidths=[
            95 * mm,
            10 * mm,
            30 * mm,
            32 * mm,
        ]
    )

    summary.setStyle(
        TableStyle([

            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.black
            ),

            (
                "SPAN",
                (0, 1),
                (1, 1)
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),

            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                3
            ),

            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                3
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                4
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                4
            ),

            (
                "ALIGN",
                (3, 0),
                (3, -1),
                "RIGHT"
            ),
        ])
    )

    story.append(
        summary
    )

    story.append(
        Spacer(
            1,
            3 * mm
        )
    )

    # --------------------------------------------------------------------------
    # PAYMENT TERMS
    # --------------------------------------------------------------------------

    terms = (
        po_data.get(
            "Payment Terms",
            ""
        )
        or
        "100% 90 Days After Invoice received"
    )

    terms_table = Table(
        [
            [
                _p(
                    terms,
                    normal
                )
            ]
        ],
        colWidths=[
            167 * mm
        ]
    )

    terms_table.setStyle(
        TableStyle([

            (
                "BOX",
                (0, 0),
                (-1, -1),
                0.5,
                colors.black
            ),

            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                3
            ),

            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                3
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                4
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                4
            ),
        ])
    )

    story.append(
        terms_table
    )

    story.append(
        Spacer(
            1,
            10 * mm
        )
    )

    # --------------------------------------------------------------------------
    # SIGNATURE
    # --------------------------------------------------------------------------

    prepared_sign = find_asset(
        "Prepared By.png",
        "Prepared.png",
        "signature.png",
        "Sign Prepared By.png",
        "Tanda Tangan.png",
    )

    signature_cells = [

        _p(
            "Prepared By,",
            normal
        ),

        _p(
            "Approved By,",
            normal
        ),

        _p(
            "Mitra,",
            normal
        ),
    ]

    sign_table = Table(
        [
            signature_cells,
            ["", "", ""]
        ],
        colWidths=[
            55 * mm,
            55 * mm,
            55 * mm,
        ]
    )

    sign_table.setStyle(
        TableStyle([

            (
                "ALIGN",
                (0, 0),
                (-1, -1),
                "CENTER"
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP"
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, 0),
                8
            ),
        ])
    )

    if prepared_sign:

        try:

            sign_img = RLImage(
                prepared_sign,
                width=35 * mm,
                height=18 * mm
            )

            sign_img.hAlign = "CENTER"

            sign_table = Table(
                [
                    signature_cells,
                    [
                        sign_img,
                        "",
                        ""
                    ],
                ],
                colWidths=[
                    55 * mm,
                    55 * mm,
                    55 * mm,
                ]
            )

            sign_table.setStyle(
                TableStyle([

                    (
                        "ALIGN",
                        (0, 0),
                        (-1, -1),
                        "CENTER"
                    ),

                    (
                        "VALIGN",
                        (0, 0),
                        (-1, -1),
                        "TOP"
                    ),

                    (
                        "BOTTOMPADDING",
                        (0, 0),
                        (-1, 0),
                        5
                    ),
                ])
            )

        except Exception:
            pass

    story.append(
        sign_table
    )

    doc.build(
        story
    )

    buf.seek(0)

    return buf.getvalue()


# ==============================================================================
# PDF PREVIEW
# ==============================================================================

def pdf_preview(
    pdf_bytes,
    height=900
):

    encoded = base64.b64encode(
        pdf_bytes
    ).decode("utf-8")

    html = f"""
    <iframe
        src="data:application/pdf;base64,{encoded}"
        width="100%"
        height="{height}"
        type="application/pdf"
        style="
            border:1px solid #ddd;
            border-radius:8px;
        ">
    </iframe>
    """

    st.components.v1.html(
        html,
        height=height + 20,
        scrolling=True
    )


# ==============================================================================
# SESSION STATE
# ==============================================================================

def _init_state():

    if "po_items" not in st.session_state:

        st.session_state.po_items = []

    if "po_loaded" not in st.session_state:

        st.session_state.po_loaded = None

    if "po_pdf" not in st.session_state:

        st.session_state.po_pdf = None

    if "po_search_result" not in st.session_state:

        st.session_state.po_search_result = (
            pd.DataFrame()
        )

    if "po_active_type" not in st.session_state:

        st.session_state.po_active_type = None


# ==============================================================================
# EMPTY ITEM
# ==============================================================================

def _empty_item(
    po_type
):

    if po_type == "Material":

        return {

            "Charger Type": "",

            "Site Name": "",

            "Description": "",

            "Qty": 1,

            "UoM": "Unit",

            "Unit Price": 0,
        }

    return {

        "Description": "",

        "Site / Location": "",

        "Qty": 1,

        "UoM": "Unit",

        "Unit Price": 0,
    }


# ==============================================================================
# ITEM EDITOR
# ==============================================================================

def _render_item_editor(
    po_type
):

    items = (
        st.session_state.po_items
    )

    if not items:

        items.append(
            _empty_item(
                po_type
            )
        )

    st.subheader(
        "Detail PO"
    )

    remove_index = None

    for i in range(
        len(items)
    ):

        item = items[i]

        with st.container(
            border=True
        ):

            cols = st.columns(
                [
                    0.5,
                    1.2,
                    1.8,
                    1.0,
                    0.9,
                    1.4,
                    1.4,
                ]
            )

            with cols[0]:

                st.markdown(
                    f"**{i + 1}**"
                )

            # ------------------------------------------------------------------
            # MATERIAL
            # ------------------------------------------------------------------

            if po_type == "Material":

                with cols[1]:

                    item[
                        "Charger Type"
                    ] = st.text_input(

                        "Charger Type",

                        value=item.get(
                            "Charger Type",
                            ""
                        ),

                        key=f"po_charger_{i}",
                    )

                with cols[2]:

                    item[
                        "Site Name"
                    ] = st.text_input(

                        "Site Name",

                        value=item.get(
                            "Site Name",
                            ""
                        ),

                        key=f"po_site_{i}",
                    )

                with cols[3]:

                    item[
                        "Qty"
                    ] = st.number_input(

                        "Qty",

                        min_value=0.0,

                        value=float(
                            item.get(
                                "Qty",
                                1
                            )
                        ),

                        step=1.0,

                        key=f"po_qty_{i}",
                    )

                with cols[4]:

                    item[
                        "UoM"
                    ] = st.text_input(

                        "UoM",

                        value=item.get(
                            "UoM",
                            "Unit"
                        ),

                        key=f"po_uom_{i}",
                    )

                with cols[5]:

                    item[
                        "Unit Price"
                    ] = st.number_input(

                        "Unit Price",

                        min_value=0.0,

                        value=float(
                            item.get(
                                "Unit Price",
                                0
                            )
                        ),

                        step=1000.0,

                        key=f"po_price_{i}",
                    )

                with cols[6]:

                    st.metric(

                        "Total",

                        rupiah(
                            _safe_float(
                                item["Qty"]
                            )
                            *
                            _safe_float(
                                item["Unit Price"]
                            )
                        )
                    )

            # ------------------------------------------------------------------
            # GENERAL
            # ------------------------------------------------------------------

            else:

                with cols[1]:

                    item[
                        "Description"
                    ] = st.text_input(

                        "Description / Service",

                        value=item.get(
                            "Description",
                            ""
                        ),

                        key=f"po_desc_{i}",
                    )

                with cols[2]:

                    item[
                        "Site / Location"
                    ] = st.text_input(

                        "Site / Location",

                        value=item.get(
                            "Site / Location",
                            ""
                        ),

                        key=f"po_location_{i}",
                    )

                with cols[3]:

                    item[
                        "Qty"
                    ] = st.number_input(

                        "Qty",

                        min_value=0.0,

                        value=float(
                            item.get(
                                "Qty",
                                1
                            )
                        ),

                        step=1.0,

                        key=f"po_g_qty_{i}",
                    )

                with cols[4]:

                    item[
                        "UoM"
                    ] = st.text_input(

                        "UoM",

                        value=item.get(
                            "UoM",
                            "Unit"
                        ),

                        key=f"po_g_uom_{i}",
                    )

                with cols[5]:

                    item[
                        "Unit Price"
                    ] = st.number_input(

                        "Unit Price",

                        min_value=0.0,

                        value=float(
                            item.get(
                                "Unit Price",
                                0
                            )
                        ),

                        step=1000.0,

                        key=f"po_g_price_{i}",
                    )

                with cols[6]:

                    st.metric(

                        "Total",

                        rupiah(
                            _safe_float(
                                item["Qty"]
                            )
                            *
                            _safe_float(
                                item["Unit Price"]
                            )
                        )
                    )

            if st.button(
                "Remove",
                key=f"po_remove_{i}",
                disabled=len(items) <= 1,
            ):

                remove_index = i

    if remove_index is not None:

        items.pop(
            remove_index
        )

        st.rerun()

    if st.button(
        "＋ Add Item",
        use_container_width=True
    ):

        items.append(
            _empty_item(
                po_type
            )
        )

        st.rerun()


# ==============================================================================
# HEADER FORM
# ==============================================================================

def _render_header_form(
    po_type,
    edit_row=None
):

    edit_row = (
        edit_row or {}
    )

    po_no_default = (
        edit_row.get(
            "PO No"
        )
        or
        get_next_po_number(
            po_type
        )
    )

    po_date_default = (
        edit_row.get(
            "PO Date"
        )
        or
        datetime.now().date()
    )

    c1, c2, c3 = st.columns(
        [
            1.4,
            1.0,
            1.4,
        ]
    )

    with c1:

        po_no = st.text_input(
            "No. PO",
            value=str(
                po_no_default
            ),
            disabled=True,
        )

    with c2:

        po_date = st.date_input(
            "Date",
            value=pd.to_datetime(
                po_date_default
            ).date()
        )

    with c3:

        project = st.text_input(
            "Project",
            value=edit_row.get(
                "Project",
                ""
            )
        )

    c1, c2 = st.columns(
        2
    )

    with c1:

        supplier = st.text_input(
            "Supplier",
            value=edit_row.get(
                "Supplier",
                ""
            )
        )

        phone = st.text_input(
            "No. Tlpn",
            value=edit_row.get(
                "Supplier Phone",
                ""
            )
        )

    with c2:

        address = st.text_area(
            "Address",
            value=edit_row.get(
                "Supplier Address",
                ""
            ),
            height=68
        )

        payment = st.text_input(
            "Payment Terms",
            value=edit_row.get(
                "Payment Terms",
                "100% 90 Days After Invoice received"
            )
        )

    tax = st.number_input(
        "Tax (%)",

        min_value=0.0,

        max_value=100.0,

        value=float(
            edit_row.get(
                "Tax Rate",
                11
            )
            or 11
        ),

        step=0.5,
    )

    return {

        "PO No":
            po_no,

        "PO Type":
            po_type,

        "PO Date":
            po_date.strftime(
                "%Y-%m-%d"
            ),

        "Supplier":
            supplier,

        "Supplier Phone":
            phone,

        "Supplier Address":
            address,

        "Project":
            project,

        "Payment Terms":
            payment,

        "Tax Rate":
            tax,

        "Approval Status":
            edit_row.get(
                "Approval Status",
                "Pending"
            )
            or "Pending",

        "Approval By":
            edit_row.get(
                "Approval By",
                ""
            ),

        "Approval Date":
            edit_row.get(
                "Approval Date",
                ""
            ),

        "Approval Notes":
            edit_row.get(
                "Approval Notes",
                ""
            ),

        "Created By":
            st.session_state.get(
                "username",
                st.session_state.get(
                    "user_email",
                    ""
                )
            ),

        "Created At":
            edit_row.get(
                "Created At",
                ""
            ),
    }


# ==============================================================================
# LOAD TO FORM
# ==============================================================================

def _load_to_form(
    po_type,
    row,
    items
):

    st.session_state.po_loaded = (
        row
    )

    st.session_state.po_items = (
        items
        or
        [
            _empty_item(
                po_type
            )
        ]
    )

    st.session_state.po_active_type = (
        po_type
    )


# ==============================================================================
# SEARCH & EDIT
# ==============================================================================

def render_search_edit():

    st.subheader(
        "Search & Edit PO"
    )

    typ = st.selectbox(
        "PO Type",

        [
            "Material",
            "Non Material / General",
        ],

        key="po_search_type",
    )

    po_type = (
        "Material"
        if typ == "Material"
        else
        "Non Material"
    )

    keyword = st.text_input(
        "Search PO / Supplier / Project",
        key="po_search_keyword",
    )

    if st.button(
        "Search",
        use_container_width=True
    ):

        result = search_po(
            po_type,
            keyword
        )

        st.session_state.po_search_result = (
            result
        )

    result = (
        st.session_state.get(
            "po_search_result",
            pd.DataFrame()
        )
    )

    if result.empty:

        return

    display_cols = [

        c

        for c in [

            "PO No",
            "PO Date",
            "Supplier",
            "Project",
            "Grand Total",
            "Approval Status",

        ]

        if c in result.columns
    ]

    st.dataframe(
        result[display_cols],
        use_container_width=True,
        hide_index=True,
    )

    choices = (
        result["PO No"]
        .astype(str)
        .tolist()
    )

    selected = st.selectbox(
        "Select PO",
        choices
    )

    if st.button(
        "Load PO",
        use_container_width=True
    ):

        loaded = load_po(
            po_type,
            selected
        )

        if loaded:

            _load_to_form(
                po_type,
                loaded[0],
                loaded[1]
            )

            st.success(
                f"PO {selected} "
                "berhasil dipanggil."
            )

            st.rerun()


# ==============================================================================
# CREATE PO
# ==============================================================================

def render_create_po(
    po_type
):

    _init_state()

    loaded = (
        st.session_state.get(
            "po_loaded"
        )
    )

    edit_mode = bool(
        loaded
        and
        loaded.get(
            "PO Type"
        ) == po_type
    )

    if edit_mode:

        st.subheader(
            "Edit Purchase Order"
        )

    else:

        if po_type == "Material":

            st.subheader(
                "Create PO Material"
            )

        else:

            st.subheader(
                "Create PO Non Material / General"
            )

    # --------------------------------------------------------------------------
    # APPROVED LOCK
    # --------------------------------------------------------------------------

    if (
        edit_mode
        and
        str(
            loaded.get(
                "Approval Status",
                ""
            )
        ).lower()
        == "approved"
    ):

        st.error(
            "PO sudah Approved oleh CEO "
            "dan tidak dapat dilakukan edit."
        )

        return

    # --------------------------------------------------------------------------
    # RESET ITEM SAAT GANTI TYPE
    # --------------------------------------------------------------------------

    if (
        st.session_state.get(
            "po_active_type"
        )
        != po_type
        and not edit_mode
    ):

        st.session_state.po_active_type = (
            po_type
        )

        st.session_state.po_items = [
            _empty_item(
                po_type
            )
        ]

    # --------------------------------------------------------------------------
    # HEADER
    # --------------------------------------------------------------------------

    po_data = _render_header_form(
        po_type,
        loaded
        if edit_mode
        else None
    )

    # --------------------------------------------------------------------------
    # ITEMS
    # --------------------------------------------------------------------------

    _render_item_editor(
        po_type
    )

    items = calculate_items(
        st.session_state.po_items
    )

    subtotal, tax_amount, grand_total = (
        calculate_totals(
            items,
            po_data[
                "Tax Rate"
            ]
        )
    )

    # --------------------------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------------------------

    st.divider()

    c1, c2, c3 = st.columns(
        3
    )

    c1.metric(
        "Total",
        rupiah(
            subtotal
        )
    )

    c2.metric(
        f"Tax {po_data['Tax Rate']:g}%",
        rupiah(
            tax_amount
        )
    )

    c3.metric(
        "Grand Total",
        rupiah(
            grand_total
        )
    )

    st.info(
        "**Terbilang:** "
        f"{english_amount_words(grand_total)}"
    )

    # --------------------------------------------------------------------------
    # BUTTON
    # --------------------------------------------------------------------------

    b1, b2 = st.columns(
        2
    )

    with b1:

        save_label = (
            "Update PO"
            if edit_mode
            else
            "Save PO"
        )

        if st.button(
            save_label,
            type="primary",
            use_container_width=True
        ):

            # Validation
            if not po_data[
                "Supplier"
            ].strip():

                st.error(
                    "Supplier wajib diisi."
                )

                return

            if not po_data[
                "Project"
            ].strip():

                st.error(
                    "Project wajib diisi."
                )

                return

            if not items:

                st.error(
                    "Minimal 1 item."
                )

                return

            # Validate item
            invalid_item = False

            for index, item in enumerate(
                items,
                start=1
            ):

                if (
                    _safe_float(
                        item.get(
                            "Qty",
                            0
                        )
                    )
                    <= 0
                ):

                    st.error(
                        f"Qty item {index} "
                        "harus lebih dari 0."
                    )

                    invalid_item = True

                    break

                if (
                    _safe_float(
                        item.get(
                            "Unit Price",
                            0
                        )
                    )
                    < 0
                ):

                    st.error(
                        f"Unit Price item {index} "
                        "tidak valid."
                    )

                    invalid_item = True

                    break

            if invalid_item:

                return

            try:

                po_no, record, normalized = (
                    save_po(
                        po_data,
                        items,
                        edit_mode=edit_mode
                    )
                )

                # Generate PDF
                pdf_bytes = create_po_pdf(
                    record,
                    normalized
                )

                st.session_state.po_loaded = (
                    record
                )

                st.session_state.po_pdf = (
                    pdf_bytes
                )

                st.session_state.po_active_type = (
                    po_type
                )

                st.session_state.po_items = (
                    normalized
                )

                st.success(
                    "Create PO berhasil "
                    "dibuat dan disimpan."
                )

                st.toast(
                    "PO berhasil disimpan.",
                    icon="✅"
                )

            except PermissionError as e:

                st.error(
                    str(e)
                )

            except Exception as e:

                st.error(
                    f"Gagal menyimpan PO: {e}"
                )

    # --------------------------------------------------------------------------
    # NEW PO
    # --------------------------------------------------------------------------

    with b2:

        if st.session_state.get(
            "po_loaded"
        ):

            if st.button(
                "New PO",
                use_container_width=True
            ):

                st.session_state.po_loaded = (
                    None
                )

                st.session_state.po_items = [
                    _empty_item(
                        po_type
                    )
                ]

                st.session_state.po_pdf = (
                    None
                )

                st.session_state.po_active_type = (
                    po_type
                )

                st.rerun()

    # --------------------------------------------------------------------------
    # PDF
    # --------------------------------------------------------------------------

    if st.session_state.get(
        "po_pdf"
    ):

        st.divider()

        st.subheader(
            "Print & Download"
        )

        pdf_bytes = (
            st.session_state.po_pdf
        )

        po_no = (
            st.session_state
            .get(
                "po_loaded",
                {}
            )
            .get(
                "PO No",
                "PO"
            )
        )

        filename = re.sub(
            r"[^A-Za-z0-9_.-]+",
            "_",
            f"{po_no}.pdf"
        )

        st.download_button(

            "⬇ Download PDF",

            data=pdf_bytes,

            file_name=filename,

            mime="application/pdf",

            use_container_width=True,
        )

        pdf_preview(
            pdf_bytes
        )


# ==============================================================================
# MAIN PO MODULE
# ==============================================================================

def render_po_module():
    """
    Entry point yang dipanggil oleh app.py.
    """

    _init_state()

    st.title(
        "Purchase Order"
    )

    st.caption(
        "CLX ERP • SCM / Procurement"
    )

    mode = st.radio(
        "Menu",

        [
            "Create PO",
            "Search & Edit",
        ],

        horizontal=True,

        key="po_main_menu",
    )

    # --------------------------------------------------------------------------
    # SEARCH
    # --------------------------------------------------------------------------

    if mode == "Search & Edit":

        render_search_edit()

        return

    # --------------------------------------------------------------------------
    # CREATE
    # --------------------------------------------------------------------------

    po_choice = st.radio(

        "Pilih Form PO",

        [
            "Material",
            "Non Material / General",
        ],

        horizontal=True,

        key="po_create_type",
    )

    po_type = (
        "Material"
        if po_choice == "Material"
        else
        "Non Material"
    )

    render_create_po(
        po_type
    )


# ==============================================================================
# ALIAS
# Supaya app.py bisa memakai beberapa kemungkinan nama fungsi
# ==============================================================================

render = render_po_module
show_po_page = render_po_module
show = render_po_module


# ==============================================================================
# STANDALONE
# ==============================================================================

if __name__ == "__main__":

    render_po_module()
