# ==============================================================================
# CREATE INVOICE.PY
# CLX ERP SYSTEM
#
# VERSION:
# CREATE INVOICE V2.0
#
# UPDATE:
# 1. Site Name selector moved OUTSIDE st.form
# 2. Live Preview Site & BOQ saat Site Name dipilih
# 3. BOQ Amount mengambil DB BOQ kolom E
# 4. Termin / Percentage realtime
# 5. Invoice Amount realtime
# 6. Save tetap menggunakan form_submit_button
# 7. PDF Invoice tetap dipertahankan
# ==============================================================================

import os
import io
import streamlit as st
import pandas as pd
from datetime import datetime
import gspread

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


# ==============================================================================
# CONFIG
# ==============================================================================

SPREADSHEET_ID = "1FU1lL3ls3jP_hAxBdx_Fu35Z9Ap4ICdHmOpMvCyA3gY"

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


# ==============================================================================
# GENERAL HELPERS
# ==============================================================================

def format_currency(value):
    """
    Format angka menjadi format Rupiah.
    """
    try:
        value = float(value or 0)
    except Exception:
        value = 0

    return f"Rp {value:,.0f}".replace(",", ".")


def parse_amount(value):
    """
    Parse berbagai format nominal:
    1000000
    1.000.000
    Rp 1.000.000
    1,000,000
    1.000,50
    """

    if value is None:
        return 0.0

    if isinstance(value, (int, float)):
        try:
            return float(value)
        except Exception:
            return 0.0

    value = str(value).strip()

    if not value:
        return 0.0

    value = (
        value.replace("Rp", "")
        .replace("rp", "")
        .replace(" ", "")
    )

    # ------------------------------------------------------------------
    # Handle Indonesian / international number format
    # ------------------------------------------------------------------
    if "." in value and "," in value:

        # Example:
        # 1.234.567,89
        # => 1234567.89
        if value.rfind(",") > value.rfind("."):
            value = value.replace(".", "")
            value = value.replace(",", ".")

        # Example:
        # 1,234,567.89
        else:
            value = value.replace(",", "")

    elif "," in value:

        parts = value.split(",")

        # 1,234,567
        if len(parts) > 2:
            value = value.replace(",", "")

        # 100,50
        elif len(parts) == 2 and len(parts[1]) <= 2:
            value = value.replace(",", ".")

        else:
            value = value.replace(",", "")

    elif "." in value:

        parts = value.split(".")

        # 1.000.000
        if len(parts) > 2:
            value = value.replace(".", "")

        # 1.000
        elif len(parts) == 2 and len(parts[1]) == 3:
            value = value.replace(".", "")

    try:
        return float(value)
    except Exception:
        return 0.0


# ==============================================================================
# GOOGLE SHEETS CONNECTION
# ==============================================================================

@st.cache_resource
def init_gspread():
    """
    Initialize Google Sheets connection.

    Supports:
    1. Streamlit secrets
    2. Local credentials.json
    """

    gc = None

    # ------------------------------------------------------------------
    # Streamlit Cloud / Secrets
    # ------------------------------------------------------------------
    try:
        if "gcp_service_account" in st.secrets:

            credentials_dict = dict(
                st.secrets["gcp_service_account"]
            )

            gc = gspread.service_account_from_dict(
                credentials_dict
            )

    except Exception:
        gc = None

    # ------------------------------------------------------------------
    # Local credentials.json
    # ------------------------------------------------------------------
    if gc is None:

        credentials_path = os.path.join(
            os.getcwd(),
            "credentials.json"
        )

        if not os.path.exists(credentials_path):

            # Try project root
            project_root = os.path.dirname(
                os.path.dirname(
                    os.path.abspath(__file__)
                )
            )

            credentials_path = os.path.join(
                project_root,
                "credentials.json"
            )

        if not os.path.exists(credentials_path):

            raise FileNotFoundError(
                "credentials.json tidak ditemukan."
            )

        gc = gspread.service_account(
            filename=credentials_path
        )

    spreadsheet = gc.open_by_key(
        SPREADSHEET_ID
    )

    sheet_query = spreadsheet.worksheet(
        "Query"
    )

    sheet_master_dropdown = spreadsheet.worksheet(
        "Master Dropdown"
    )

    sheet_erp_project = spreadsheet.worksheet(
        "ERP Project"
    )

    sheet_db_invoice = spreadsheet.worksheet(
        "DB Invoice"
    )

    sheet_db_boq = spreadsheet.worksheet(
        "DB BOQ"
    )

    return (
        sheet_query,
        sheet_master_dropdown,
        sheet_erp_project,
        sheet_db_invoice,
        sheet_db_boq,
    )


# ==============================================================================
# GOOGLE SHEETS RAW DATA
# ==============================================================================

def get_raw_matrix(sheet):
    """
    Read worksheet as matrix.
    """

    try:
        return sheet.get_all_values()

    except Exception as e:

        st.error(
            f"Gagal membaca sheet `{sheet.title}`: {e}"
        )

        return []


# ==============================================================================
# AUTO INVOICE NUMBER
# ==============================================================================

def generate_auto_invoice_no(sheet_db):
    """
    Generate Invoice Number:

    0001/INV/CLX/IX/2026

    Existing sequence minimum starts at 462.
    """

    now = datetime.now()

    month_roman = MONTH_ROMAN.get(
        now.month,
        ""
    )

    try:
        records = sheet_db.get_all_records()
    except Exception:
        records = []

    try:

        next_seq = max(
            len(records) + 1,
            462
        )

    except Exception:

        next_seq = 462

    return (
        f"{next_seq:04d}/INV/CLX/"
        f"{month_roman}/{now.year}"
    )


# ==============================================================================
# QUERY DATA
# ==============================================================================

def get_site_data_from_query(
    raw_query,
    site_name
):
    """
    Query structure:

    C = Charging Type
    F = Site Name
    W = WO Number
    """

    if not raw_query:
        return {
            "site_name": site_name,
            "charging_type": "",
            "wo_number": "",
        }

    site_target = str(
        site_name
    ).strip().lower()

    for row in raw_query:

        if len(row) <= 22:
            continue

        query_site = str(
            row[5]
        ).strip().lower()

        if query_site == site_target:

            return {
                "site_name": row[5],
                "charging_type": row[2],
                "wo_number": row[22],
            }

    return {
        "site_name": site_name,
        "charging_type": "",
        "wo_number": "",
    }


# ==============================================================================
# DB BOQ DATA
# ==============================================================================

def get_boq_amount_from_db_boq(
    raw_db_boq,
    site_name,
    charging_type
):
    """
    DB BOQ structure:

    A = No
    B = BOQ No.
    C = Site Name
    D = Charger Type
    E = BOQ Amount Exc. PPN
    F = BOQ Amount inc. PPN
    G = EPC Name

    Return:
    BOQ Amount Exc. PPN
    """

    if not raw_db_boq:
        return 0.0

    site_target = str(
        site_name
    ).strip().lower()

    charger_target = str(
        charging_type
    ).strip().lower()

    for row in raw_db_boq:

        if len(row) <= 4:
            continue

        db_site = str(
            row[2]
        ).strip().lower()

        db_charger = str(
            row[3]
        ).strip().lower()

        if (
            db_site == site_target
            and db_charger == charger_target
        ):

            return parse_amount(
                row[4]
            )

    return 0.0


# ==============================================================================
# BUILD SITE DATA
# ==============================================================================

def build_site_data(
    raw_query,
    raw_db_boq,
    selected_sites
):
    """
    Build complete site data.

    Output:

    {
        site_name,
        charging_type,
        wo_number,
        boq_amount
    }
    """

    result = []

    if not selected_sites:
        return result

    for site in selected_sites:

        query_data = get_site_data_from_query(
            raw_query,
            site
        )

        site_name = query_data.get(
            "site_name",
            site
        )

        charging_type = query_data.get(
            "charging_type",
            ""
        )

        wo_number = query_data.get(
            "wo_number",
            ""
        )

        boq_amount = get_boq_amount_from_db_boq(
            raw_db_boq,
            site_name,
            charging_type
        )

        result.append(
            {
                "site_name": site_name,
                "charging_type": charging_type,
                "wo_number": wo_number,
                "boq_amount": boq_amount,
            }
        )

    return result


# ==============================================================================
# PREVIEW SITE & BOQ
# ==============================================================================

def render_site_boq_preview(
    site_data,
    show_invoice_amount=False,
    selected_pct=0
):
    """
    Live preview table.

    Dipanggil setelah Site Name dipilih.
    """

    if not site_data:

        st.info(
            "🔎 Pilih Site Name untuk melihat preview BOQ."
        )

        return

    st.markdown(
        "### 🔎 Preview Site & BOQ"
    )

    preview_rows = []

    missing_boq = []

    for index, item in enumerate(
        site_data,
        start=1
    ):

        boq_amount = float(
            item.get(
                "boq_amount",
                0
            ) or 0
        )

        if boq_amount <= 0:

            missing_boq.append(
                item.get(
                    "site_name",
                    ""
                )
            )

        invoice_amount = (
            boq_amount
            * float(selected_pct or 0)
            / 100
        )

        row = {
            "NO.": index,
            "SITE NAME": item.get(
                "site_name",
                ""
            ),
            "CHARGER TYPE": item.get(
                "charging_type",
                ""
            ),
            "WO NUMBER": item.get(
                "wo_number",
                ""
            ),
            "BOQ AMOUNT EXC. PPN": format_currency(
                boq_amount
            ),
        }

        if show_invoice_amount:

            row[
                "INVOICE AMOUNT"
            ] = format_currency(
                invoice_amount
            )

        preview_rows.append(
            row
        )

    preview_df = pd.DataFrame(
        preview_rows
    )

    st.dataframe(
        preview_df,
        use_container_width=True,
        hide_index=True,
    )

    if missing_boq:

        st.warning(
            "⚠️ BOQ Amount belum ditemukan "
            "untuk site berikut:\n\n"
            + "\n".join(
                [
                    f"- {site}"
                    for site in missing_boq
                ]
            )
        )

    else:

        st.success(
            "✅ Semua Site memiliki data "
            "BOQ Amount dari DB BOQ kolom E."
        )


# ==============================================================================
# PDF GENERATOR
# ==============================================================================

def create_invoice_pdf(data):
    """
    Generate Invoice PDF.
    """

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=90,
        bottomMargin=80,
    )

    styles = getSampleStyleSheet()

    normal_style = ParagraphStyle(
        "NormalInvoice",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
    )

    small_style = ParagraphStyle(
        "SmallInvoice",
        parent=styles["Normal"],
        fontSize=7,
        leading=9,
    )

    title_style = ParagraphStyle(
        "InvoiceTitle",
        parent=styles["Heading1"],
        fontSize=16,
        leading=20,
        alignment=1,
        spaceAfter=10,
    )

    story = []

    # ------------------------------------------------------------------
    # Header image
    # ------------------------------------------------------------------

    project_root = os.path.dirname(
        os.path.dirname(
            os.path.abspath(__file__)
        )
    )

    header_path = os.path.join(
        project_root,
        "assets",
        "header.png"
    )

    footer_path = os.path.join(
        project_root,
        "assets",
        "Footer.png"
    )

    if os.path.exists(header_path):

        try:

            from reportlab.platypus import Image

            header_img = Image(
                header_path
            )

            header_img.drawHeight = 0.55 * 72
            header_img.drawWidth = 7.0 * 72

            story.append(
                header_img
            )

            story.append(
                Spacer(1, 8)
            )

        except Exception:
            pass

    # ------------------------------------------------------------------
    # Title
    # ------------------------------------------------------------------

    story.append(
        Paragraph(
            "INVOICE",
            title_style
        )
    )

    story.append(
        Paragraph(
            "<b>Bill To:</b> "
            "PT Vgreen Global Charging "
            "Station Investment Indonesia",
            normal_style
        )
    )

    story.append(
        Spacer(1, 8)
    )

    # ------------------------------------------------------------------
    # Invoice Meta
    # ------------------------------------------------------------------

    meta_data = [
        [
            Paragraph(
                "<b>No Invoice</b>",
                normal_style
            ),
            Paragraph(
                str(
                    data.get(
                        "invoice_no",
                        ""
                    )
                ),
                normal_style
            ),
        ],
        [
            Paragraph(
                "<b>Invoice Date</b>",
                normal_style
            ),
            Paragraph(
                str(
                    data.get(
                        "invoice_date",
                        ""
                    )
                ),
                normal_style
            ),
        ],
        [
            Paragraph(
                "<b>Project Name</b>",
                normal_style
            ),
            Paragraph(
                str(
                    data.get(
                        "project_name",
                        ""
                    )
                ),
                normal_style
            ),
        ],
        [
            Paragraph(
                "<b>No Efaktur</b>",
                normal_style
            ),
            Paragraph(
                str(
                    data.get(
                        "efaktur_no",
                        ""
                    )
                ),
                normal_style
            ),
        ],
    ]

    meta_table = Table(
        meta_data,
        colWidths=[
            100,
            390
        ]
    )

    meta_table.setStyle(
        TableStyle(
            [
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
                    0
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    2
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    2
                ),
            ]
        )
    )

    story.append(
        meta_table
    )

    story.append(
        Spacer(1, 12)
    )

    # ------------------------------------------------------------------
    # Item Table
    # ------------------------------------------------------------------

    table_header = [
        "NO.",
        "ITEM DESCRIPTION",
        "QTY",
        "UOM",
        "CHARGER TYPE",
        "WO NUMBER",
        "UNIT PRICE",
        "AMOUNT",
    ]

    item_rows = [
        [
            Paragraph(
                f"<b>{x}</b>",
                small_style
            )
            for x in table_header
        ]
    ]

    for index, item in enumerate(
        data.get(
            "site_data",
            []
        ),
        start=1
    ):

        amount = float(
            item.get(
                "invoice_amount",
                0
            ) or 0
        )

        item_rows.append(
            [
                Paragraph(
                    str(index),
                    small_style
                ),
                Paragraph(
                    str(
                        item.get(
                            "site_name",
                            ""
                        )
                    ),
                    small_style
                ),
                Paragraph(
                    "1",
                    small_style
                ),
                Paragraph(
                    "Unit",
                    small_style
                ),
                Paragraph(
                    str(
                        item.get(
                            "charging_type",
                            ""
                        )
                    ),
                    small_style
                ),
                Paragraph(
                    str(
                        item.get(
                            "wo_number",
                            ""
                        )
                    ),
                    small_style
                ),
                Paragraph(
                    format_currency(
                        amount
                    ),
                    small_style
                ),
                Paragraph(
                    format_currency(
                        amount
                    ),
                    small_style
                ),
            ]
        )

    item_table = Table(
        item_rows,
        colWidths=[
            25,
            125,
            30,
            35,
            70,
            80,
            75,
            75,
        ],
        repeatRows=1,
    )

    item_table.setStyle(
        TableStyle(
            [
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
                    colors.lightgrey
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE"
                ),
                (
                    "ALIGN",
                    (0, 0),
                    (0, -1),
                    "CENTER"
                ),
                (
                    "ALIGN",
                    (2, 1),
                    (2, -1),
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
                    (6, 1),
                    (-1, -1),
                    "RIGHT"
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
            ]
        )
    )

    story.append(
        item_table
    )

    story.append(
        Spacer(1, 10)
    )

    # ------------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------------

    subtotal = float(
        data.get(
            "subtotal",
            0
        ) or 0
    )

    tax_rate = float(
        data.get(
            "tax_rate",
            11
        ) or 0
    )

    tax_amount = float(
        data.get(
            "tax_amount",
            0
        ) or 0
    )

    grand_total = float(
        data.get(
            "grand_total",
            0
        ) or 0
    )

    summary_rows = [
        [
            "",
            Paragraph(
                "<b>Subtotal</b>",
                normal_style
            ),
            Paragraph(
                format_currency(
                    subtotal
                ),
                normal_style
            ),
        ],
        [
            "",
            Paragraph(
                f"<b>PPN {tax_rate:.0f}%</b>",
                normal_style
            ),
            Paragraph(
                format_currency(
                    tax_amount
                ),
                normal_style
            ),
        ],
        [
            "",
            Paragraph(
                "<b>TOTAL</b>",
                normal_style
            ),
            Paragraph(
                format_currency(
                    grand_total
                ),
                normal_style
            ),
        ],
    ]

    summary_table = Table(
        summary_rows,
        colWidths=[
            270,
            120,
            125
        ]
    )

    summary_table.setStyle(
        TableStyle(
            [
                (
                    "ALIGN",
                    (2, 0),
                    (2, -1),
                    "RIGHT"
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
                    "LINEABOVE",
                    (1, 2),
                    (2, 2),
                    1,
                    colors.black
                ),
            ]
        )
    )

    story.append(
        summary_table
    )

    story.append(
        Spacer(1, 15)
    )

    # ------------------------------------------------------------------
    # Payment Info
    # ------------------------------------------------------------------

    story.append(
        Paragraph(
            "<b>Payment Information</b>",
            normal_style
        )
    )

    story.append(
        Paragraph(
            "Bank : BCA<br/>"
            "Account Number : 540-5282841<br/>"
            "Account Name : PT. CONNECTIVITY LEADS EXCELLENCE",
            normal_style
        )
    )

    story.append(
        Spacer(1, 20)
    )

    # ------------------------------------------------------------------
    # Signature
    # ------------------------------------------------------------------

    today_str = datetime.now().strftime(
        "%d %B %Y"
    )

    signature_data = [
        [
            Paragraph(
                f"Jakarta, {today_str}",
                normal_style
            )
        ],
        [
            Spacer(1, 40)
        ],
        [
            Paragraph(
                "<b>Christian</b>",
                normal_style
            )
        ],
        [
            Paragraph(
                "President Director",
                normal_style
            )
        ],
    ]

    signature_table = Table(
        signature_data,
        colWidths=[
            200
        ]
    )

    signature_table.setStyle(
        TableStyle(
            [
                (
                    "ALIGN",
                    (0, 0),
                    (-1, -1),
                    "CENTER"
                ),
            ]
        )
    )

    story.append(
        signature_table
    )

    # ------------------------------------------------------------------
    # Footer image
    # ------------------------------------------------------------------

    if os.path.exists(footer_path):

        try:

            from reportlab.platypus import Image

            story.append(
                Spacer(1, 20)
            )

            footer_img = Image(
                footer_path
            )

            footer_img.drawHeight = 0.35 * 72
            footer_img.drawWidth = 7.0 * 72

            story.append(
                footer_img
            )

        except Exception:
            pass

    doc.build(
        story
    )

    buffer.seek(0)

    return buffer.getvalue()


# ==============================================================================
# MAIN RENDER
# ==============================================================================

def render():

    # ------------------------------------------------------------------
    # PAGE HEADER
    # ------------------------------------------------------------------

    st.title(
        "🧾 Create Invoice"
    )

    st.caption(
        "Create Invoice berdasarkan Site Name, "
        "BOQ, WO Number dan Termin."
    )

    # ------------------------------------------------------------------
    # CONNECT GOOGLE SHEETS
    # ------------------------------------------------------------------

    try:

        (
            sheet_query,
            sheet_master_dropdown,
            sheet_erp_project,
            sheet_db_invoice,
            sheet_db_boq,
        ) = init_gspread()

    except Exception as e:

        st.error(
            f"❌ Gagal koneksi Google Sheets: {e}"
        )

        return

    # ------------------------------------------------------------------
    # LOAD DATA
    # ------------------------------------------------------------------

    raw_query = get_raw_matrix(
        sheet_query
    )

    raw_dropdown = get_raw_matrix(
        sheet_master_dropdown
    )

    raw_erp = get_raw_matrix(
        sheet_erp_project
    )

    raw_db = get_raw_matrix(
        sheet_db_invoice
    )

    raw_db_boq = get_raw_matrix(
        sheet_db_boq
    )

    # ------------------------------------------------------------------
    # AUTO INVOICE NUMBER
    # ------------------------------------------------------------------

    auto_inv_no = generate_auto_invoice_no(
        sheet_db_invoice
    )

    # ------------------------------------------------------------------
    # PROJECT OPTIONS
    # ------------------------------------------------------------------

    project_options = []

    if raw_dropdown:

        for row in raw_dropdown:

            if len(row) <= 6:
                continue

            value = str(
                row[6]
            ).strip()

            if value and value not in project_options:

                project_options.append(
                    value
                )

    # Fallback
    if not project_options:

        project_options = [
            "VGreen - Project",
            "VGreen - Operation",
            "SIP",
            "Charge Core",
        ]

    # Ensure known projects exist
    known_projects = [
        "VGreen - Project",
        "VGreen - Operation",
        "SIP",
        "Charge Core",
    ]

    for project in known_projects:

        if project not in project_options:

            project_options.append(
                project
            )

    # ------------------------------------------------------------------
    # BASIC INFORMATION
    #
    # IMPORTANT:
    # Semua widget di bawah ini berada OUTSIDE FORM
    # agar setiap perubahan langsung rerun.
    # ------------------------------------------------------------------

    st.markdown(
        "### 📋 Invoice Information"
    )

    col_left, col_right = st.columns(
        2
    )

    with col_left:

        selected_project = st.selectbox(
            "Project Name",
            project_options,
            key="invoice_project_name",
        )

        st.text_input(
            "Invoice No.",
            value=auto_inv_no,
            disabled=True,
            key="invoice_number_display",
        )

        inv_date = st.date_input(
            "Invoice Date",
            value=datetime.now().date(),
            key="invoice_date",
        )

    with col_right:

        efaktur_no = st.text_input(
            "No Efaktur",
            key="invoice_efaktur",
            placeholder="Masukkan nomor Efaktur",
        )

        mode_invoice = st.selectbox(
            "Mode",
            [
                "Invoice",
                "Proforma Invoice",
            ],
            key="invoice_mode",
        )

    st.divider()

    # ==============================================================================
    # SITE SELECTION
    # ==============================================================================

    st.markdown(
        "### 📍 Site Selection"
    )

    selected_sites = []

    if selected_project == "VGreen - Project":

        site_options = []

        if raw_query:

            for row in raw_query:

                if len(row) <= 5:
                    continue

                site_name = str(
                    row[5]
                ).strip()

                if (
                    site_name
                    and site_name not in site_options
                ):

                    site_options.append(
                        site_name
                    )

        site_options = sorted(
            site_options
        )

        selected_sites = st.multiselect(
            "Pilih Site Name",
            options=site_options,
            key="invoice_site_selector",
            placeholder="Pilih satu atau beberapa Site Name",
            help=(
                "Setelah Site Name dipilih, "
                "Preview BOQ akan langsung muncul."
            ),
        )

    else:

        manual_site = st.text_input(
            "Site Name",
            key="invoice_manual_site",
            placeholder="Masukkan Site Name",
        )

        if manual_site.strip():

            selected_sites = [
                manual_site.strip()
            ]

    # ==============================================================================
    # BUILD SITE DATA
    # ==============================================================================

    site_data = build_site_data(
        raw_query,
        raw_db_boq,
        selected_sites
    )

    # ==============================================================================
    # LIVE PREVIEW
    #
    # INI BAGIAN UTAMA PERBAIKAN
    # ==============================================================================

    if selected_sites:

        render_site_boq_preview(
            site_data=site_data,
            show_invoice_amount=False,
        )

    else:

        st.info(
            "👆 Silakan pilih Site Name terlebih dahulu. "
            "Preview BOQ akan muncul otomatis di sini."
        )

    st.divider()

    # ==============================================================================
    # TOP / TERMIN
    # ==============================================================================

    st.markdown(
        "### 💰 Invoice Calculation"
    )

    # ------------------------------------------------------------------
    # Determine used terms
    # ------------------------------------------------------------------

    used_terms = set()

    if raw_db and site_data:

        for row in raw_db:

            if len(row) <= 6:
                continue

            try:

                charging_db = str(
                    row[3]
                ).strip().lower()

                site_db = str(
                    row[4]
                ).strip().lower()

                term_db = str(
                    row[6]
                ).strip()

            except Exception:

                continue

            if not term_db:
                continue

            for item in site_data:

                item_site = str(
                    item.get(
                        "site_name",
                        ""
                    )
                ).strip().lower()

                item_charging = str(
                    item.get(
                        "charging_type",
                        ""
                    )
                ).strip().lower()

                if (
                    site_db == item_site
                    and charging_db == item_charging
                ):

                    used_terms.add(
                        term_db
                    )

    # ------------------------------------------------------------------
    # TOP Schema
    # ------------------------------------------------------------------

    top_schema = st.selectbox(
        "Pilih TOP",
        [
            "Standard 35/60/5",
            "New 10/85/5",
        ],
        key="invoice_top_schema",
    )

    if top_schema == "Standard 35/60/5":

        term_options = [
            "DP",
            "Termin 2",
            "Termin 3",
            "Retensi",
        ]

        term_percentages = {
            "DP": 35,
            "Termin 2": 60,
            "Termin 3": 0,
            "Retensi": 5,
        }

    else:

        term_options = [
            "DP",
            "Termin 2",
            "Termin 3",
            "Retensi",
        ]

        term_percentages = {
            "DP": 10,
            "Termin 2": 85,
            "Termin 3": 0,
            "Retensi": 5,
        }

    # ------------------------------------------------------------------
    # Available terms
    # ------------------------------------------------------------------

    available_terms = [
        term
        for term in term_options
        if term not in used_terms
    ]

    if not available_terms:

        available_terms = term_options

    selected_termin = st.selectbox(
        "Pilih Termin",
        available_terms,
        key="invoice_termin",
    )

    selected_pct = st.number_input(
        "Percentage (%)",
        min_value=0.0,
        max_value=100.0,
        value=float(
            term_percentages.get(
                selected_termin,
                0
            )
        ),
        step=1.0,
        key="invoice_percentage",
    )

    # ==============================================================================
    # CALCULATION
    # ==============================================================================

    total_boq_amount = 0.0
    total_invoice_amount = 0.0

    for item in site_data:

        boq_amount = float(
            item.get(
                "boq_amount",
                0
            ) or 0
        )

        invoice_amount = (
            boq_amount
            * float(selected_pct)
            / 100
        )

        item[
            "invoice_amount"
        ] = invoice_amount

        total_boq_amount += (
            boq_amount
        )

        total_invoice_amount += (
            invoice_amount
        )

    # ==============================================================================
    # SECOND LIVE PREVIEW
    # ==============================================================================

    if site_data:

        st.markdown(
            "### 💳 Preview Invoice Amount"
        )

        invoice_preview_rows = []

        for index, item in enumerate(
            site_data,
            start=1
        ):

            boq_amount = float(
                item.get(
                    "boq_amount",
                    0
                ) or 0
            )

            invoice_amount = float(
                item.get(
                    "invoice_amount",
                    0
                ) or 0
            )

            invoice_preview_rows.append(
                {
                    "NO.": index,
                    "SITE NAME": item.get(
                        "site_name",
                        ""
                    ),
                    "CHARGER TYPE": item.get(
                        "charging_type",
                        ""
                    ),
                    "WO NUMBER": item.get(
                        "wo_number",
                        ""
                    ),
                    "BOQ EXC. PPN": format_currency(
                        boq_amount
                    ),
                    f"INVOICE ({selected_pct:.0f}%)": format_currency(
                        invoice_amount
                    ),
                }
            )

        invoice_preview_df = pd.DataFrame(
            invoice_preview_rows
        )

        st.dataframe(
            invoice_preview_df,
            use_container_width=True,
            hide_index=True,
        )

    # ==============================================================================
    # PRICE INFORMATION
    # ==============================================================================

    if selected_project == "VGreen - Project":

        st.number_input(
            "Total Invoice Amount Before PPN",
            min_value=0.0,
            value=float(
                total_invoice_amount
            ),
            disabled=True,
            format="%.0f",
            key="invoice_total_before_ppn",
        )

        termin_amount = (
            total_invoice_amount
        )

    else:

        manual_unit_price = st.number_input(
            "Invoice Amount Before PPN",
            min_value=0.0,
            value=float(
                total_invoice_amount
            ),
            step=1000.0,
            format="%.0f",
            key="invoice_manual_amount",
        )

        termin_amount = (
            manual_unit_price
        )

    # ==============================================================================
    # TAX
    # ==============================================================================

    tax_rate = st.number_input(
        "PPN (%)",
        min_value=0.0,
        max_value=100.0,
        value=11.0,
        step=0.5,
        key="invoice_tax_rate",
    )

    tax_amount = (
        termin_amount
        * tax_rate
        / 100
    )

    grand_total = (
        termin_amount
        + tax_amount
    )

    # ==============================================================================
    # SUMMARY
    # ==============================================================================

    st.markdown(
        "### 📊 Invoice Summary"
    )

    summary_col1, summary_col2, summary_col3 = st.columns(
        3
    )

    with summary_col1:

        st.metric(
            "BOQ Amount Exc. PPN",
            format_currency(
                total_boq_amount
            )
        )

    with summary_col2:

        st.metric(
            "Invoice Before PPN",
            format_currency(
                termin_amount
            )
        )

    with summary_col3:

        st.metric(
            f"Grand Total + PPN {tax_rate:.0f}%",
            format_currency(
                grand_total
            )
        )

    # ==============================================================================
    # VALIDATION
    # ==============================================================================

    missing_boq_sites = []

    for item in site_data:

        if float(
            item.get(
                "boq_amount",
                0
            ) or 0
        ) <= 0:

            missing_boq_sites.append(
                item.get(
                    "site_name",
                    ""
                )
            )

    if (
        selected_project == "VGreen - Project"
        and missing_boq_sites
    ):

        st.error(
            "❌ Invoice belum dapat disimpan "
            "karena BOQ Amount belum ditemukan "
            "untuk Site berikut:\n\n"
            + "\n".join(
                [
                    f"- {x}"
                    for x in missing_boq_sites
                ]
            )
        )

    # ==============================================================================
    # SAVE FORM
    #
    # HANYA tombol Save berada di dalam FORM.
    # Semua input sebelumnya berada di luar form.
    # ==============================================================================

    st.divider()

    with st.form(
        "form_save_invoice",
        clear_on_submit=False,
    ):

        submit_btn = st.form_submit_button(
            "💾 Save Invoice to DB Invoice",
            use_container_width=True,
            disabled=(
                not selected_sites
                or (
                    selected_project
                    == "VGreen - Project"
                    and bool(
                        missing_boq_sites
                    )
                )
            ),
        )

    # ==============================================================================
    # SAVE PROCESS
    # ==============================================================================

    if submit_btn:

        # --------------------------------------------------------------
        # Validate Efaktur
        # --------------------------------------------------------------

        if not efaktur_no.strip():

            st.error(
                "❌ No Efaktur wajib diisi."
            )

            return

        # --------------------------------------------------------------
        # Validate Site
        # --------------------------------------------------------------

        if not selected_sites:

            st.error(
                "❌ Site Name wajib dipilih."
            )

            return

        # --------------------------------------------------------------
        # Validate BOQ
        # --------------------------------------------------------------

        if (
            selected_project
            == "VGreen - Project"
            and missing_boq_sites
        ):

            st.error(
                "❌ Tidak dapat menyimpan "
                "karena ada Site yang belum "
                "memiliki BOQ Amount."
            )

            return

        # --------------------------------------------------------------
        # Save each site
        # --------------------------------------------------------------

        save_rows = []

        for item in site_data:

            charging_type = item.get(
                "charging_type",
                ""
            )

            site_name = item.get(
                "site_name",
                ""
            )

            wo_number = item.get(
                "wo_number",
                ""
            )

            amount = float(
                item.get(
                    "invoice_amount",
                    0
                ) or 0
            )

            amount_inc_tax = (
                amount
                + (
                    amount
                    * tax_rate
                    / 100
                )
            )

            save_rows.append(
                [
                    selected_project,
                    auto_inv_no,
                    inv_date.strftime(
                        "%Y-%m-%d"
                    ),
                    charging_type,
                    site_name,
                    wo_number,
                    selected_termin,
                    efaktur_no,
                    amount,
                    amount_inc_tax,
                ]
            )

        # --------------------------------------------------------------
        # Append to DB Invoice
        # --------------------------------------------------------------

        try:

            for row in save_rows:

                sheet_db_invoice.append_row(
                    row,
                    value_input_option="USER_ENTERED"
                )

            st.success(
                f"✅ Invoice `{auto_inv_no}` "
                f"berhasil disimpan ke DB Invoice."
            )

            st.balloons()

        except Exception as e:

            st.error(
                f"❌ Gagal menyimpan Invoice "
                f"ke DB Invoice: {e}"
            )

            return

        # ==============================================================================
        # GENERATE PDF
        # ==============================================================================

        pdf_site_data = []

        for item in site_data:

            pdf_site_data.append(
                {
                    "site_name": item.get(
                        "site_name",
                        ""
                    ),
                    "charging_type": item.get(
                        "charging_type",
                        ""
                    ),
                    "wo_number": item.get(
                        "wo_number",
                        ""
                    ),
                    "invoice_amount": float(
                        item.get(
                            "invoice_amount",
                            0
                        ) or 0
                    ),
                }
            )

        pdf_payload = {
            "invoice_no": auto_inv_no,
            "invoice_date": inv_date.strftime(
                "%d-%m-%Y"
            ),
            "project_name": selected_project,
            "efaktur_no": efaktur_no,
            "site_data": pdf_site_data,
            "subtotal": termin_amount,
            "tax_rate": tax_rate,
            "tax_amount": tax_amount,
            "grand_total": grand_total,
        }

        try:

            pdf_bytes = create_invoice_pdf(
                pdf_payload
            )

            pdf_name = (
                f"{auto_inv_no.replace('/', '-')}"
                f".pdf"
            )

            st.session_state[
                "pdf_ready"
            ] = pdf_bytes

            st.session_state[
                "pdf_name"
            ] = pdf_name

            st.success(
                "📄 PDF Invoice berhasil dibuat."
            )

        except Exception as e:

            st.error(
                f"❌ Gagal membuat PDF Invoice: {e}"
            )

    # ==============================================================================
    # DOWNLOAD PDF
    # ==============================================================================

    if (
        "pdf_ready"
        in st.session_state
    ):

        st.divider()

        st.markdown(
            "### 📄 Download Invoice"
        )

        pdf_bytes = st.session_state[
            "pdf_ready"
        ]

        pdf_name = st.session_state.get(
            "pdf_name",
            "Invoice.pdf"
        )

        st.download_button(
            label="⬇️ Download Invoice PDF",
            data=pdf_bytes,
            file_name=pdf_name,
            mime="application/pdf",
            use_container_width=True,
        )


# ==============================================================================
# ENTRY POINT
# ==============================================================================

if __name__ == "__main__":
    render()
