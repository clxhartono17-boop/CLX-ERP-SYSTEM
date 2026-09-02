# ==============================================================================
# CREATE INVOICE.PY
# CLX ERP SYSTEM
#
# VERSION:
# CREATE INVOICE V2.1
#
# UPDATE V2.1:
# 1. Site Name menggunakan MULTISELECT
# 2. Bisa memilih lebih dari 1 Site
# 3. Site selection berada OUTSIDE st.form
# 4. Live Preview Site & BOQ
# 5. Session State untuk menjaga pilihan Site
# 6. BOQ Amount mengambil DB BOQ kolom E
# 7. Invoice Amount realtime
# 8. Total BOQ realtime
# 9. Total Invoice realtime
# 10. Save ke DB Invoice tetap berjalan
# 11. Generate PDF tetap berjalan
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
from reportlab.lib.styles import (
    getSampleStyleSheet,
    ParagraphStyle,
)


# ==============================================================================
# CONFIG
# ==============================================================================

SPREADSHEET_ID = (
    "1FU1lL3ls3jP_hAxBdx_Fu35Z9Ap4ICdHmOpMvCyA3gY"
)

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
# SESSION STATE INITIALIZATION
# ==============================================================================

def initialize_session_state():

    if "invoice_selected_sites" not in st.session_state:
        st.session_state[
            "invoice_selected_sites"
        ] = []

    if "pdf_ready" not in st.session_state:
        st.session_state[
            "pdf_ready"
        ] = None

    if "pdf_name" not in st.session_state:
        st.session_state[
            "pdf_name"
        ] = "Invoice.pdf"


# ==============================================================================
# FORMAT CURRENCY
# ==============================================================================

def format_currency(value):

    try:
        value = float(value or 0)
    except Exception:
        value = 0

    return (
        f"Rp {value:,.0f}"
        .replace(",", ".")
    )


# ==============================================================================
# PARSE AMOUNT
# ==============================================================================

def parse_amount(value):

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
        value
        .replace("Rp", "")
        .replace("rp", "")
        .replace(" ", "")
    )

    # --------------------------------------------------------------
    # Both "." and ","
    # --------------------------------------------------------------

    if "." in value and "," in value:

        # Indonesian:
        # 1.234.567,89
        if value.rfind(",") > value.rfind("."):

            value = value.replace(".", "")
            value = value.replace(",", ".")

        # International:
        # 1,234,567.89
        else:

            value = value.replace(",", "")

    # --------------------------------------------------------------
    # Comma only
    # --------------------------------------------------------------

    elif "," in value:

        parts = value.split(",")

        if len(parts) > 2:

            value = value.replace(",", "")

        elif (
            len(parts) == 2
            and len(parts[1]) <= 2
        ):

            value = value.replace(",", ".")

        else:

            value = value.replace(",", "")

    # --------------------------------------------------------------
    # Dot only
    # --------------------------------------------------------------

    elif "." in value:

        parts = value.split(".")

        if len(parts) > 2:

            value = value.replace(".", "")

        elif (
            len(parts) == 2
            and len(parts[1]) == 3
        ):

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

    gc = None

    # ------------------------------------------------------------------
    # Streamlit Secrets
    # ------------------------------------------------------------------

    try:

        if "gcp_service_account" in st.secrets:

            credentials_dict = dict(
                st.secrets[
                    "gcp_service_account"
                ]
            )

            gc = (
                gspread
                .service_account_from_dict(
                    credentials_dict
                )
            )

    except Exception:

        gc = None

    # ------------------------------------------------------------------
    # Local credentials.json
    # ------------------------------------------------------------------

    if gc is None:

        credentials_path = os.path.join(
            os.getcwd(),
            "credentials.json",
        )

        if not os.path.exists(
            credentials_path
        ):

            project_root = os.path.dirname(
                os.path.dirname(
                    os.path.abspath(__file__)
                )
            )

            credentials_path = os.path.join(
                project_root,
                "credentials.json",
            )

        if not os.path.exists(
            credentials_path
        ):

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

    sheet_master_dropdown = (
        spreadsheet.worksheet(
            "Master Dropdown"
        )
    )

    sheet_erp_project = (
        spreadsheet.worksheet(
            "ERP Project"
        )
    )

    sheet_db_invoice = (
        spreadsheet.worksheet(
            "DB Invoice"
        )
    )

    sheet_db_boq = (
        spreadsheet.worksheet(
            "DB BOQ"
        )
    )

    return (
        sheet_query,
        sheet_master_dropdown,
        sheet_erp_project,
        sheet_db_invoice,
        sheet_db_boq,
    )


# ==============================================================================
# GET RAW MATRIX
# ==============================================================================

def get_raw_matrix(sheet):

    try:

        return sheet.get_all_values()

    except Exception as e:

        st.error(
            f"Gagal membaca sheet "
            f"`{sheet.title}`: {e}"
        )

        return []


# ==============================================================================
# AUTO INVOICE NUMBER
# ==============================================================================

def generate_auto_invoice_no(
    sheet_db
):

    now = datetime.now()

    month_roman = MONTH_ROMAN.get(
        now.month,
        "",
    )

    try:

        records = (
            sheet_db
            .get_all_records()
        )

    except Exception:

        records = []

    try:

        next_seq = max(
            len(records) + 1,
            462,
        )

    except Exception:

        next_seq = 462

    return (
        f"{next_seq:04d}/INV/CLX/"
        f"{month_roman}/{now.year}"
    )


# ==============================================================================
# QUERY SITE DATA
# ==============================================================================

def get_site_data_from_query(
    raw_query,
    site_name,
):

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
# GET BOQ AMOUNT
# ==============================================================================

def get_boq_amount_from_db_boq(
    raw_db_boq,
    site_name,
    charging_type,
):

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
    selected_sites,
):

    result = []

    if not selected_sites:
        return result

    for site in selected_sites:

        query_data = (
            get_site_data_from_query(
                raw_query,
                site,
            )
        )

        site_name = query_data.get(
            "site_name",
            site,
        )

        charging_type = (
            query_data.get(
                "charging_type",
                "",
            )
        )

        wo_number = (
            query_data.get(
                "wo_number",
                "",
            )
        )

        boq_amount = (
            get_boq_amount_from_db_boq(
                raw_db_boq,
                site_name,
                charging_type,
            )
        )

        result.append(
            {
                "site_name": site_name,
                "charging_type": (
                    charging_type
                ),
                "wo_number": wo_number,
                "boq_amount": boq_amount,
            }
        )

    return result


# ==============================================================================
# SITE SELECTION CALLBACK
# ==============================================================================

def update_selected_sites():

    selected = st.session_state.get(
        "invoice_site_selector_multi_v2",
        [],
    )

    # Pastikan selalu list
    if selected is None:

        selected = []

    if not isinstance(
        selected,
        list,
    ):

        selected = list(selected)

    st.session_state[
        "invoice_selected_sites"
    ] = selected


# ==============================================================================
# RENDER SITE BOQ PREVIEW
# ==============================================================================

def render_site_boq_preview(
    site_data,
):

    if not site_data:

        st.info(
            "👆 Pilih Site Name terlebih dahulu "
            "untuk melihat Preview BOQ."
        )

        return

    st.markdown(
        "### 🔎 Preview Site & BOQ"
    )

    preview_rows = []

    missing_boq = []

    for index, item in enumerate(
        site_data,
        start=1,
    ):

        boq_amount = float(
            item.get(
                "boq_amount",
                0,
            )
            or 0
        )

        if boq_amount <= 0:

            missing_boq.append(
                item.get(
                    "site_name",
                    "",
                )
            )

        preview_rows.append(
            {
                "NO.": index,
                "SITE NAME": item.get(
                    "site_name",
                    "",
                ),
                "CHARGER TYPE": item.get(
                    "charging_type",
                    "",
                ),
                "WO NUMBER": item.get(
                    "wo_number",
                    "",
                ),
                "BOQ AMOUNT EXC. PPN": (
                    format_currency(
                        boq_amount
                    )
                ),
            }
        )

    preview_df = pd.DataFrame(
        preview_rows
    )

    st.dataframe(
        preview_df,
        use_container_width=True,
        hide_index=True,
    )

    # ------------------------------------------------------------------
    # Summary selected site
    # ------------------------------------------------------------------

    total_selected = len(
        site_data
    )

    total_boq = sum(
        float(
            item.get(
                "boq_amount",
                0,
            )
            or 0
        )
        for item in site_data
    )

    col1, col2 = st.columns(2)

    with col1:

        st.metric(
            "Jumlah Site",
            total_selected,
        )

    with col2:

        st.metric(
            "Total BOQ Exc. PPN",
            format_currency(
                total_boq
            ),
        )

    # ------------------------------------------------------------------
    # BOQ validation
    # ------------------------------------------------------------------

    if missing_boq:

        st.warning(
            "⚠️ BOQ Amount belum ditemukan "
            "untuk Site berikut:"
        )

        for site in missing_boq:

            st.write(
                f"- {site}"
            )

    else:

        st.success(
            "✅ Semua Site memiliki "
            "BOQ Amount dari DB BOQ kolom E."
        )


# ==============================================================================
# CREATE INVOICE PDF
# ==============================================================================

def create_invoice_pdf(
    data
):

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
    # PROJECT ROOT
    # ------------------------------------------------------------------

    project_root = os.path.dirname(
        os.path.dirname(
            os.path.abspath(__file__)
        )
    )

    header_path = os.path.join(
        project_root,
        "assets",
        "header.png",
    )

    footer_path = os.path.join(
        project_root,
        "assets",
        "Footer.png",
    )

    # ------------------------------------------------------------------
    # HEADER
    # ------------------------------------------------------------------

    if os.path.exists(
        header_path
    ):

        try:

            from reportlab.platypus import Image

            header_img = Image(
                header_path
            )

            header_img.drawHeight = (
                0.55 * 72
            )

            header_img.drawWidth = (
                7.0 * 72
            )

            story.append(
                header_img
            )

            story.append(
                Spacer(1, 8)
            )

        except Exception:

            pass

    # ------------------------------------------------------------------
    # TITLE
    # ------------------------------------------------------------------

    story.append(
        Paragraph(
            "INVOICE",
            title_style,
        )
    )

    story.append(
        Paragraph(
            "<b>Bill To:</b> "
            "PT Vgreen Global Charging "
            "Station Investment Indonesia",
            normal_style,
        )
    )

    story.append(
        Spacer(1, 8)
    )

    # ------------------------------------------------------------------
    # META
    # ------------------------------------------------------------------

    meta_data = [
        [
            Paragraph(
                "<b>No Invoice</b>",
                normal_style,
            ),
            Paragraph(
                str(
                    data.get(
                        "invoice_no",
                        "",
                    )
                ),
                normal_style,
            ),
        ],
        [
            Paragraph(
                "<b>Invoice Date</b>",
                normal_style,
            ),
            Paragraph(
                str(
                    data.get(
                        "invoice_date",
                        "",
                    )
                ),
                normal_style,
            ),
        ],
        [
            Paragraph(
                "<b>Project Name</b>",
                normal_style,
            ),
            Paragraph(
                str(
                    data.get(
                        "project_name",
                        "",
                    )
                ),
                normal_style,
            ),
        ],
        [
            Paragraph(
                "<b>No Efaktur</b>",
                normal_style,
            ),
            Paragraph(
                str(
                    data.get(
                        "efaktur_no",
                        "",
                    )
                ),
                normal_style,
            ),
        ],
    ]

    meta_table = Table(
        meta_data,
        colWidths=[
            100,
            390,
        ],
    )

    meta_table.setStyle(
        TableStyle(
            [
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "TOP",
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    4,
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    2,
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    2,
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
    # ITEM TABLE
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
                small_style,
            )
            for x in table_header
        ]
    ]

    for index, item in enumerate(
        data.get(
            "site_data",
            [],
        ),
        start=1,
    ):

        amount = float(
            item.get(
                "invoice_amount",
                0,
            )
            or 0
        )

        item_rows.append(
            [
                Paragraph(
                    str(index),
                    small_style,
                ),
                Paragraph(
                    str(
                        item.get(
                            "site_name",
                            "",
                        )
                    ),
                    small_style,
                ),
                Paragraph(
                    "1",
                    small_style,
                ),
                Paragraph(
                    "Unit",
                    small_style,
                ),
                Paragraph(
                    str(
                        item.get(
                            "charging_type",
                            "",
                        )
                    ),
                    small_style,
                ),
                Paragraph(
                    str(
                        item.get(
                            "wo_number",
                            "",
                        )
                    ),
                    small_style,
                ),
                Paragraph(
                    format_currency(
                        amount
                    ),
                    small_style,
                ),
                Paragraph(
                    format_currency(
                        amount
                    ),
                    small_style,
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
                    colors.black,
                ),
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.lightgrey,
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE",
                ),
                (
                    "ALIGN",
                    (0, 0),
                    (0, -1),
                    "CENTER",
                ),
                (
                    "ALIGN",
                    (2, 1),
                    (2, -1),
                    "CENTER",
                ),
                (
                    "ALIGN",
                    (3, 1),
                    (3, -1),
                    "CENTER",
                ),
                (
                    "ALIGN",
                    (6, 1),
                    (-1, -1),
                    "RIGHT",
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    3,
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    3,
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    4,
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    4,
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
    # SUMMARY
    # ------------------------------------------------------------------

    subtotal = float(
        data.get(
            "subtotal",
            0,
        )
        or 0
    )

    tax_rate = float(
        data.get(
            "tax_rate",
            11,
        )
        or 0
    )

    tax_amount = float(
        data.get(
            "tax_amount",
            0,
        )
        or 0
    )

    grand_total = float(
        data.get(
            "grand_total",
            0,
        )
        or 0
    )

    summary_rows = [
        [
            "",
            Paragraph(
                "<b>Subtotal</b>",
                normal_style,
            ),
            Paragraph(
                format_currency(
                    subtotal
                ),
                normal_style,
            ),
        ],
        [
            "",
            Paragraph(
                f"<b>PPN {tax_rate:.0f}%</b>",
                normal_style,
            ),
            Paragraph(
                format_currency(
                    tax_amount
                ),
                normal_style,
            ),
        ],
        [
            "",
            Paragraph(
                "<b>TOTAL</b>",
                normal_style,
            ),
            Paragraph(
                format_currency(
                    grand_total
                ),
                normal_style,
            ),
        ],
    ]

    summary_table = Table(
        summary_rows,
        colWidths=[
            270,
            120,
            125,
        ],
    )

    summary_table.setStyle(
        TableStyle(
            [
                (
                    "ALIGN",
                    (2, 0),
                    (2, -1),
                    "RIGHT",
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    4,
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    4,
                ),
                (
                    "LINEABOVE",
                    (1, 2),
                    (2, 2),
                    1,
                    colors.black,
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
    # PAYMENT INFO
    # ------------------------------------------------------------------

    story.append(
        Paragraph(
            "<b>Payment Information</b>",
            normal_style,
        )
    )

    story.append(
        Paragraph(
            "Bank : BCA<br/>"
            "Account Number : 540-5282841<br/>"
            "Account Name : "
            "PT. CONNECTIVITY LEADS EXCELLENCE",
            normal_style,
        )
    )

    story.append(
        Spacer(1, 20)
    )

    # ------------------------------------------------------------------
    # SIGNATURE
    # ------------------------------------------------------------------

    today_str = datetime.now().strftime(
        "%d %B %Y"
    )

    signature_data = [
        [
            Paragraph(
                f"Jakarta, {today_str}",
                normal_style,
            )
        ],
        [
            Spacer(1, 40)
        ],
        [
            Paragraph(
                "<b>Christian</b>",
                normal_style,
            )
        ],
        [
            Paragraph(
                "President Director",
                normal_style,
            )
        ],
    ]

    signature_table = Table(
        signature_data,
        colWidths=[
            200
        ],
    )

    signature_table.setStyle(
        TableStyle(
            [
                (
                    "ALIGN",
                    (0, 0),
                    (-1, -1),
                    "CENTER",
                ),
            ]
        )
    )

    story.append(
        signature_table
    )

    # ------------------------------------------------------------------
    # FOOTER
    # ------------------------------------------------------------------

    if os.path.exists(
        footer_path
    ):

        try:

            from reportlab.platypus import Image

            story.append(
                Spacer(1, 20)
            )

            footer_img = Image(
                footer_path
            )

            footer_img.drawHeight = (
                0.35 * 72
            )

            footer_img.drawWidth = (
                7.0 * 72
            )

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

    initialize_session_state()

    # ------------------------------------------------------------------
    # PAGE HEADER
    # ------------------------------------------------------------------

    st.title(
        "🧾 Create Invoice"
    )

    st.caption(
        "Create Invoice berdasarkan "
        "Site Name, BOQ, WO Number dan Termin."
    )

    # ------------------------------------------------------------------
    # CONNECT
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
    # AUTO INVOICE NO
    # ------------------------------------------------------------------

    auto_inv_no = generate_auto_invoice_no(
        sheet_db_invoice
    )

    # ==============================================================================
    # PROJECT OPTIONS
    # ==============================================================================

    project_options = []

    if raw_dropdown:

        for row in raw_dropdown:

            if len(row) <= 6:
                continue

            value = str(
                row[6]
            ).strip()

            if (
                value
                and value not in project_options
            ):

                project_options.append(
                    value
                )

    if not project_options:

        project_options = [
            "VGreen - Project",
            "VGreen - Operation",
            "SIP",
            "Charge Core",
        ]

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

    # ==============================================================================
    # INVOICE INFORMATION
    # ==============================================================================

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
            key="invoice_project_name_v2",
        )

        st.text_input(
            "Invoice No.",
            value=auto_inv_no,
            disabled=True,
            key="invoice_number_display_v2",
        )

        inv_date = st.date_input(
            "Invoice Date",
            value=datetime.now().date(),
            key="invoice_date_v2",
        )

    with col_right:

        efaktur_no = st.text_input(
            "No Efaktur",
            key="invoice_efaktur_v2",
            placeholder=(
                "Masukkan nomor Efaktur"
            ),
        )

        mode_invoice = st.selectbox(
            "Mode",
            [
                "Invoice",
                "Proforma Invoice",
            ],
            key="invoice_mode_v2",
        )

    st.divider()

    # ==============================================================================
    # SITE SELECTION
    # ==============================================================================

    st.markdown(
        "### 📍 Site Selection"
    )

    selected_sites = []

    # ==============================================================================
    # VGREEN PROJECT
    # ==============================================================================

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
            site_options,
            key=lambda x: x.lower(),
        )

        # ------------------------------------------------------------------
        # IMPORTANT:
        #
        # MULTISELECT
        # Bisa pilih banyak Site.
        #
        # TIDAK ADA max_selections=1
        # ------------------------------------------------------------------

        current_selection = (
            st.session_state.get(
                "invoice_selected_sites",
                [],
            )
        )

        # ------------------------------------------------------------------
        # Clean session state:
        # hanya pertahankan site yang masih
        # tersedia di options.
        # ------------------------------------------------------------------

        current_selection = [
            x
            for x in current_selection
            if x in site_options
        ]

        st.session_state[
            "invoice_selected_sites"
        ] = current_selection

        selected_sites = st.multiselect(
            "Pilih Site Name",
            options=site_options,

            # ----------------------------------------------------------
            # NO max_selections
            # ----------------------------------------------------------
            # Artinya unlimited selection.
            # ----------------------------------------------------------

            default=current_selection,

            key=(
                "invoice_site_selector_multi_v2"
            ),

            on_change=(
                update_selected_sites
            ),

            placeholder=(
                "Pilih satu atau beberapa "
                "Site Name"
            ),

            help=(
                "Bisa memilih lebih dari "
                "satu Site Name. "
                "Klik Site berikutnya untuk "
                "menambahkan ke pilihan."
            ),
        )

        # ------------------------------------------------------------------
        # Sinkronkan kembali session state
        # ------------------------------------------------------------------

        st.session_state[
            "invoice_selected_sites"
        ] = list(
            selected_sites
        )

        # ------------------------------------------------------------------
        # Selected site counter
        # ------------------------------------------------------------------

        if selected_sites:

            st.caption(
                f"📌 {len(selected_sites)} "
                f"Site dipilih"
            )

    # ==============================================================================
    # NON VGREEN PROJECT
    # ==============================================================================

    else:

        manual_site = st.text_input(
            "Site Name",
            key="invoice_manual_site_v2",
            placeholder=(
                "Masukkan Site Name"
            ),
        )

        if manual_site.strip():

            selected_sites = [
                manual_site.strip()
            ]

        else:

            selected_sites = []

    # ==============================================================================
    # BUILD SITE DATA
    # ==============================================================================

    site_data = build_site_data(
        raw_query,
        raw_db_boq,
        selected_sites,
    )

    # ==============================================================================
    # LIVE SITE & BOQ PREVIEW
    # ==============================================================================

    render_site_boq_preview(
        site_data
    )

    st.divider()

    # ==============================================================================
    # INVOICE CALCULATION
    # ==============================================================================

    st.markdown(
        "### 💰 Invoice Calculation"
    )

    # ==============================================================================
    # CHECK USED TERMS
    # ==============================================================================

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
                        "",
                    )
                ).strip().lower()

                item_charging = str(
                    item.get(
                        "charging_type",
                        "",
                    )
                ).strip().lower()

                if (
                    site_db == item_site
                    and charging_db
                    == item_charging
                ):

                    used_terms.add(
                        term_db
                    )

    # ==============================================================================
    # TOP SCHEMA
    # ==============================================================================

    top_schema = st.selectbox(
        "Pilih TOP",
        [
            "Standard 35/60/5",
            "New 10/85/5",
        ],
        key="invoice_top_schema_v2",
    )

    if (
        top_schema
        == "Standard 35/60/5"
    ):

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

    # ==============================================================================
    # AVAILABLE TERMS
    # ==============================================================================

    available_terms = [
        term
        for term in term_options
        if term not in used_terms
    ]

    if not available_terms:

        available_terms = term_options

    # ==============================================================================
    # TERMIN
    # ==============================================================================

    selected_termin = st.selectbox(
        "Pilih Termin",
        available_terms,
        key="invoice_termin_v2",
    )

    # ==============================================================================
    # PERCENTAGE
    # ==============================================================================

    default_pct = float(
        term_percentages.get(
            selected_termin,
            0,
        )
    )

    selected_pct = st.number_input(
        "Percentage (%)",
        min_value=0.0,
        max_value=100.0,
        value=default_pct,
        step=1.0,
        key="invoice_percentage_v2",
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
                0,
            )
            or 0
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
    # LIVE INVOICE PREVIEW
    # ==============================================================================

    if site_data:

        st.markdown(
            "### 💳 Preview Invoice Amount"
        )

        invoice_preview_rows = []

        for index, item in enumerate(
            site_data,
            start=1,
        ):

            boq_amount = float(
                item.get(
                    "boq_amount",
                    0,
                )
                or 0
            )

            invoice_amount = float(
                item.get(
                    "invoice_amount",
                    0,
                )
                or 0
            )

            invoice_preview_rows.append(
                {
                    "NO.": index,
                    "SITE NAME": item.get(
                        "site_name",
                        "",
                    ),
                    "CHARGER TYPE": item.get(
                        "charging_type",
                        "",
                    ),
                    "WO NUMBER": item.get(
                        "wo_number",
                        "",
                    ),
                    "BOQ EXC. PPN": (
                        format_currency(
                            boq_amount
                        )
                    ),
                    (
                        f"INVOICE "
                        f"({selected_pct:.0f}%)"
                    ): format_currency(
                        invoice_amount
                    ),
                }
            )

        invoice_preview_df = (
            pd.DataFrame(
                invoice_preview_rows
            )
        )

        st.dataframe(
            invoice_preview_df,
            use_container_width=True,
            hide_index=True,
        )

    # ==============================================================================
    # PRICE
    # ==============================================================================

    if (
        selected_project
        == "VGreen - Project"
    ):

        st.number_input(
            "Total Invoice Amount Before PPN",
            min_value=0.0,
            value=float(
                total_invoice_amount
            ),
            disabled=True,
            format="%.0f",
            key="invoice_total_before_ppn_v2",
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
            key="invoice_manual_amount_v2",
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
        key="invoice_tax_rate_v2",
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

    summary_col1, summary_col2, summary_col3 = (
        st.columns(3)
    )

    with summary_col1:

        st.metric(
            "Jumlah Site",
            len(site_data),
        )

    with summary_col2:

        st.metric(
            "Total BOQ Exc. PPN",
            format_currency(
                total_boq_amount
            ),
        )

    with summary_col3:

        st.metric(
            "Grand Total",
            format_currency(
                grand_total
            ),
        )

    # ==============================================================================
    # VALIDATION
    # ==============================================================================

    missing_boq_sites = []

    for item in site_data:

        if float(
            item.get(
                "boq_amount",
                0,
            )
            or 0
        ) <= 0:

            missing_boq_sites.append(
                item.get(
                    "site_name",
                    "",
                )
            )

    if (
        selected_project
        == "VGreen - Project"
        and missing_boq_sites
    ):

        st.error(
            "❌ Invoice belum dapat disimpan "
            "karena BOQ Amount belum ditemukan "
            "untuk Site berikut:"
        )

        for site in missing_boq_sites:

            st.write(
                f"- {site}"
            )

    # ==============================================================================
    # SAVE BUTTON
    # ==============================================================================

    st.divider()

    with st.form(
        "form_save_invoice_v2",
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

        # ------------------------------------------------------------------
        # Efaktur
        # ------------------------------------------------------------------

        if not efaktur_no.strip():

            st.error(
                "❌ No Efaktur wajib diisi."
            )

            return

        # ------------------------------------------------------------------
        # Site
        # ------------------------------------------------------------------

        if not selected_sites:

            st.error(
                "❌ Site Name wajib dipilih."
            )

            return

        # ------------------------------------------------------------------
        # BOQ
        # ------------------------------------------------------------------

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

        # ------------------------------------------------------------------
        # SAVE ROWS
        # ------------------------------------------------------------------

        save_rows = []

        for item in site_data:

            charging_type = item.get(
                "charging_type",
                "",
            )

            site_name = item.get(
                "site_name",
                "",
            )

            wo_number = item.get(
                "wo_number",
                "",
            )

            amount = float(
                item.get(
                    "invoice_amount",
                    0,
                )
                or 0
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

        # ------------------------------------------------------------------
        # APPEND DB INVOICE
        # ------------------------------------------------------------------

        try:

            for row in save_rows:

                sheet_db_invoice.append_row(
                    row,
                    value_input_option=(
                        "USER_ENTERED"
                    ),
                )

            st.success(
                f"✅ Invoice `{auto_inv_no}` "
                f"berhasil disimpan ke DB Invoice "
                f"untuk {len(save_rows)} Site."
            )

            st.balloons()

        except Exception as e:

            st.error(
                "❌ Gagal menyimpan Invoice "
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
                        "",
                    ),
                    "charging_type": item.get(
                        "charging_type",
                        "",
                    ),
                    "wo_number": item.get(
                        "wo_number",
                        "",
                    ),
                    "invoice_amount": float(
                        item.get(
                            "invoice_amount",
                            0,
                        )
                        or 0
                    ),
                }
            )

        pdf_payload = {
            "invoice_no": auto_inv_no,
            "invoice_date": inv_date.strftime(
                "%d-%m-%Y"
            ),
            "project_name": (
                selected_project
            ),
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
                "❌ Gagal membuat PDF Invoice: "
                f"{e}"
            )

    # ==============================================================================
    # DOWNLOAD PDF
    # ==============================================================================

    if (
        st.session_state.get(
            "pdf_ready"
        )
        is not None
    ):

        st.divider()

        st.markdown(
            "### 📄 Download Invoice"
        )

        st.download_button(
            label=(
                "⬇️ Download Invoice PDF"
            ),
            data=st.session_state[
                "pdf_ready"
            ],
            file_name=st.session_state.get(
                "pdf_name",
                "Invoice.pdf",
            ),
            mime="application/pdf",
            use_container_width=True,
        )


# ==============================================================================
# ENTRY POINT
# ==============================================================================

if __name__ == "__main__":

    render()
