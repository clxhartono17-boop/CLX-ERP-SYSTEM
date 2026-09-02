import os
import io
import json
import streamlit as st
import pandas as pd
from datetime import datetime
import gspread

# Import ReportLab untuk PDF
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PIL import Image as PILImage


# ==============================================================================
# 🎯 SPREADSHEET ID & KONSTANTA
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
    12: "XII"
}


# ==============================================================================
# 🔗 GOOGLE SHEETS CONNECTION
# ==============================================================================
@st.cache_resource
def init_gspread():
    """
    Fungsi inisialisasi gspread yang fleksibel
    (Cloud Secrets vs Local File)
    """

    gc = None

    # --------------------------------------------------------------------------
    # 1. Streamlit Secrets
    # --------------------------------------------------------------------------
    if "gcp_service_account" in st.secrets:
        try:
            creds_dict = dict(st.secrets["gcp_service_account"])
            gc = gspread.service_account_from_dict(creds_dict)

        except Exception as e:
            st.warning(
                f"⚠️ Gagal menghubungkan Secrets Cloud: {e}"
            )

    # --------------------------------------------------------------------------
    # 2. Local credentials.json
    # --------------------------------------------------------------------------
    if gc is None:

        cred_path = "credentials.json"

        if os.path.exists(cred_path):

            try:
                gc = gspread.service_account(
                    filename=cred_path
                )

            except Exception as e:

                st.error(
                    f"🚨 File credentials.json ditemukan "
                    f"tapi gagal dibaca: {e}"
                )

                return None, None, None, None

        else:

            st.error(
                "🚨 Kredensial tidak ditemukan! "
                "Harap set 'Secrets' di Streamlit Cloud "
                "atau sediakan file 'credentials.json' "
                "di localhost."
            )

            return None, None, None, None

    # --------------------------------------------------------------------------
    # 3. Open Spreadsheet
    # --------------------------------------------------------------------------
    try:

        sh = gc.open_by_key(SPREADSHEET_ID)

        sheet_query = sh.worksheet("Query")
        sheet_dropdown = sh.worksheet("Master Dropdown")
        sheet_erp = sh.worksheet("ERP Project")
        sheet_db = sh.worksheet("DB Invoice")

        return (
            sheet_query,
            sheet_dropdown,
            sheet_erp,
            sheet_db
        )

    except Exception as e:

        st.error(
            f"🚨 Gagal terhubung ke Google Sheets: {e}"
        )

        return None, None, None, None


# ==============================================================================
# 📥 GET RAW MATRIX
# ==============================================================================
def get_raw_matrix(sheet):
    """
    Membaca seluruh isi sheet menjadi 2D List
    """

    try:

        if sheet is None:
            return []

        return sheet.get_all_values()

    except Exception:
        return []


# ==============================================================================
# 🔢 GENERATE AUTO INVOICE NUMBER
# ==============================================================================
def generate_auto_invoice_no(sheet_db):
    """
    Generate No. Invoice otomatis.

    Format:
    0000/INV/CLX/BULAN_ROMAWI/TAHUN

    Nomor awal:
    0462

    Penting:
    Tidak menggunakan len(records) karena satu invoice sekarang
    dapat menghasilkan beberapa baris untuk beberapa site.
    """

    now = datetime.now()

    roman_month = MONTH_ROMAN.get(
        now.month,
        "I"
    )

    year = now.year

    next_seq = 462

    if sheet_db:

        try:

            records = sheet_db.get_all_records()

            existing_numbers = []

            for record in records:

                inv_no = str(
                    record.get(
                        "Invoice No.",
                        record.get(
                            "Invoice No",
                            ""
                        )
                    )
                ).strip()

                if not inv_no:
                    continue

                try:

                    seq_part = inv_no.split("/")[0].strip()

                    seq_number = int(seq_part)

                    existing_numbers.append(
                        seq_number
                    )

                except Exception:
                    continue

            if existing_numbers:

                next_seq = max(
                    max(existing_numbers) + 1,
                    462
                )

        except Exception:

            next_seq = 462

    formatted_seq = f"{next_seq:04d}"

    return (
        f"{formatted_seq}/INV/CLX/"
        f"{roman_month}/{year}"
    )


# ==============================================================================
# 🔎 FIND SITE INFORMATION
# ==============================================================================
def get_site_information(
    raw_query,
    raw_erp,
    site_name
):
    """
    Mengambil Charging Type, WO Number dan BOQ
    berdasarkan Site Name.

    Return:
        {
            charging_type,
            wo_number,
            boq_amount
        }
    """

    charging_type = ""
    wo_number = ""
    boq_amount = 0.0

    site_lower = str(
        site_name
    ).strip().lower()

    # --------------------------------------------------------------------------
    # QUERY → Charging Type + WO
    # --------------------------------------------------------------------------
    if len(raw_query) > 1:

        for row in raw_query[1:]:

            if len(row) <= 5:
                continue

            query_site = str(
                row[5]
            ).strip().lower()

            if query_site == site_lower:

                # Charging Type
                if len(row) > 2:

                    charging_type = str(
                        row[2]
                    ).strip()

                # WO Number
                if len(row) > 22:

                    wo_number = str(
                        row[22]
                    ).strip()

                break

    # --------------------------------------------------------------------------
    # ERP → BOQ Amount
    # --------------------------------------------------------------------------
    if len(raw_erp) > 1:

        for row in raw_erp[1:]:

            if len(row) <= 4:
                continue

            erp_site = str(
                row[4]
            ).strip().lower()

            erp_charge = ""

            if len(row) > 2:

                erp_charge = str(
                    row[2]
                ).strip().lower()

            if (
                erp_site == site_lower
                and (
                    not charging_type
                    or erp_charge
                    == charging_type.strip().lower()
                )
            ):

                if len(row) > 12:

                    raw_boq = str(
                        row[12]
                    )

                    raw_boq = (
                        raw_boq
                        .replace("Rp", "")
                        .replace("rp", "")
                        .replace(".", "")
                        .replace(",", ".")
                        .replace(" ", "")
                        .strip()
                    )

                    try:

                        boq_amount = float(
                            raw_boq
                        )

                    except ValueError:

                        boq_amount = 0.0

                break

    return {
        "charging_type": charging_type,
        "wo_number": wo_number,
        "boq_amount": boq_amount
    }


# ==============================================================================
# 📄 CREATE INVOICE PDF
# ==============================================================================
def create_invoice_pdf(data):
    """
    Membuat PDF Invoice.

    Mendukung:
    - 1 site
    - Multiple site
    """

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=90,
        bottomMargin=80
    )

    elements = []

    styles = getSampleStyleSheet()

    # --------------------------------------------------------------------------
    # CUSTOM STYLES
    # --------------------------------------------------------------------------
    title_style = ParagraphStyle(
        name="TitleStyle",
        parent=styles["Heading1"],
        fontSize=22,
        leading=26,
        textColor=colors.HexColor("#1A365D"),
        alignment=2
    )

    normal_bold = ParagraphStyle(
        name="NormalBold",
        parent=styles["Normal"],
        fontSize=8.5,
        leading=11,
        fontName="Helvetica-Bold"
    )

    normal_text = ParagraphStyle(
        name="NormalText",
        parent=styles["Normal"],
        fontSize=8.5,
        leading=11
    )

    pay_title_style = ParagraphStyle(
        name="PayTitle",
        parent=styles["Normal"],
        fontSize=9,
        leading=11,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#1A365D")
    )

    site_style = ParagraphStyle(
        name="SiteStyle",
        parent=styles["Normal"],
        fontSize=7.5,
        leading=9
    )

    # --------------------------------------------------------------------------
    # PATH HEADER & FOOTER
    # --------------------------------------------------------------------------
    base_dir = os.path.abspath(
        os.path.join(
            os.path.dirname(__file__),
            "..",
            ".."
        )
    )

    header_path = os.path.join(
        base_dir,
        "assets",
        "header.png"
    )

    footer_path = os.path.join(
        base_dir,
        "assets",
        "Footer.png"
    )

    if not os.path.exists(header_path):

        header_path = os.path.abspath(
            "assets/header.png"
        )

    if not os.path.exists(footer_path):

        footer_path = os.path.abspath(
            "assets/Footer.png"
        )

    # --------------------------------------------------------------------------
    # HEADER & FOOTER
    # --------------------------------------------------------------------------
    def draw_header_footer(canvas, doc):

        canvas.saveState()

        page_w, page_h = letter

        if os.path.exists(header_path):

            try:

                canvas.drawImage(
                    header_path,
                    0,
                    page_h - 75,
                    width=page_w,
                    height=75,
                    mask="auto"
                )

            except Exception:
                pass

        if os.path.exists(footer_path):

            try:

                canvas.drawImage(
                    footer_path,
                    0,
                    0,
                    width=page_w,
                    height=65,
                    mask="auto"
                )

            except Exception:
                pass

        canvas.restoreState()

    # --------------------------------------------------------------------------
    # 1. INVOICE TITLE
    # --------------------------------------------------------------------------
    inv_title = Paragraph(
        "<b>INVOICE</b>",
        title_style
    )

    elements.append(inv_title)

    elements.append(
        Spacer(1, 10)
    )

    # --------------------------------------------------------------------------
    # 2. BILL TO
    # --------------------------------------------------------------------------
    bill_to_text = Paragraph(
        """
        <b>Bill To :</b><br/>
        <b>PT Vgreen Global Charging Station Investment Indonesia</b><br/>
        Graha Binakarsa Lt.7, Jl. H.R Rasuna Said Kav C-18 RT 02 RW 005,<br/>
        Karet Kuningan, Kec. Setiabudi Jakarta Selatan
        """,
        normal_text
    )

    issuer_and_meta = [
        [
            Paragraph(
                "<b>Invoice No.</b>",
                normal_bold
            ),
            Paragraph(
                f": {data['inv_no']}",
                normal_text
            )
        ],
        [
            Paragraph(
                "<b>Invoice Date</b>",
                normal_bold
            ),
            Paragraph(
                f": {data['inv_date']}",
                normal_text
            )
        ],
        [
            Paragraph(
                "<b>Project Name</b>",
                normal_bold
            ),
            Paragraph(
                f": {data['project_name']}",
                normal_text
            )
        ],
        [
            Paragraph(
                "<b>No. Efaktur</b>",
                normal_bold
            ),
            Paragraph(
                f": {data['efaktur']}",
                normal_text
            )
        ],
    ]

    meta_table = Table(
        issuer_and_meta,
        colWidths=[80, 180]
    )

    meta_table.setStyle(
        TableStyle([
            (
                "PADDING",
                (0, 0),
                (-1, -1),
                2
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP"
            )
        ])
    )

    info_grid = Table(
        [[
            bill_to_text,
            meta_table
        ]],
        colWidths=[280, 260]
    )

    info_grid.setStyle(
        TableStyle([
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP"
            ),
            (
                "PADDING",
                (0, 0),
                (-1, -1),
                0
            )
        ])
    )

    elements.append(info_grid)

    elements.append(
        Spacer(1, 15)
    )

    # --------------------------------------------------------------------------
    # 3. ITEM TABLE - MULTIPLE SITE
    # --------------------------------------------------------------------------
    table_data = [
        [
            Paragraph(
                "<b>NO.</b>",
                normal_bold
            ),
            Paragraph(
                "<b>ITEM DESCRIPTION / SITE</b>",
                normal_bold
            ),
            Paragraph(
                "<b>CHARGING TYPE</b>",
                normal_bold
            ),
            Paragraph(
                "<b>WO NO.</b>",
                normal_bold
            ),
            Paragraph(
                "<b>TERMIN</b>",
                normal_bold
            ),
            Paragraph(
                "<b>AMOUNT (IDR)</b>",
                normal_bold
            )
        ]
    ]

    site_rows = data.get(
        "site_rows",
        []
    )

    for idx, site in enumerate(
        site_rows,
        start=1
    ):

        table_data.append([
            str(idx),

            Paragraph(
                str(
                    site.get(
                        "site_name",
                        ""
                    )
                ),
                site_style
            ),

            Paragraph(
                str(
                    site.get(
                        "charging_type",
                        ""
                    )
                ),
                site_style
            ),

            Paragraph(
                str(
                    site.get(
                        "wo_number",
                        ""
                    )
                ),
                site_style
            ),

            str(
                site.get(
                    "termin",
                    ""
                )
            ),

            f"{site.get('termin_amount', 0):,.2f}"
        ])

    item_table = Table(
        table_data,
        colWidths=[
            25,
            190,
            75,
            90,
            55,
            105
        ],
        repeatRows=1
    )

    item_table.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor("#9FA5AD")
            ),
            (
                "TEXTCOLOR",
                (0, 0),
                (-1, 0),
                colors.white
            ),
            (
                "ALIGN",
                (0, 0),
                (0, -1),
                "CENTER"
            ),
            (
                "ALIGN",
                (4, 0),
                (4, -1),
                "CENTER"
            ),
            (
                "ALIGN",
                (5, 0),
                (5, -1),
                "RIGHT"
            ),
            (
                "PADDING",
                (0, 0),
                (-1, -1),
                5
            ),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.HexColor("#CBD5E0")
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            )
        ])
    )

    elements.append(item_table)

    elements.append(
        Spacer(1, 10)
    )

    # --------------------------------------------------------------------------
    # 4. SUMMARY TOTALS
    # --------------------------------------------------------------------------
    summary_data = [
        [
            "Subtotal (Excl. Tax) :",
            f"Rp {data['subtotal']:,.2f}"
        ],
        [
            f"PPN ({data['tax_rate']}%) :",
            f"Rp {data['tax_amount']:,.2f}"
        ],
        [
            "TOTAL AMOUNT :",
            f"Rp {data['grand_total']:,.2f}"
        ]
    ]

    summary_table = Table(
        summary_data,
        colWidths=[380, 160]
    )

    summary_table.setStyle(
        TableStyle([
            (
                "ALIGN",
                (0, 0),
                (-1, -1),
                "RIGHT"
            ),
            (
                "FONTNAME",
                (0, -1),
                (-1, -1),
                "Helvetica-Bold"
            ),
            (
                "LINEABOVE",
                (0, -1),
                (-1, -1),
                1,
                colors.HexColor("#2B6CB0")
            ),
            (
                "PADDING",
                (0, 0),
                (-1, -1),
                4
            )
        ])
    )

    elements.append(summary_table)

    elements.append(
        Spacer(1, 15)
    )

    # --------------------------------------------------------------------------
    # 5. PAYMENT INFO
    # --------------------------------------------------------------------------
    pay_table_data = [
        [
            Paragraph(
                "<b>PAYMENT INFO</b>",
                pay_title_style
            ),
            "",
            ""
        ],
        [
            Paragraph(
                "Bank Name",
                normal_bold
            ),
            ":",
            Paragraph(
                "BANK CENTRAL ASIA (BCA)",
                normal_text
            )
        ],
        [
            Paragraph(
                "Bank Account",
                normal_bold
            ),
            ":",
            Paragraph(
                "540-5282841",
                normal_text
            )
        ],
        [
            Paragraph(
                "Account Name",
                normal_bold
            ),
            ":",
            Paragraph(
                "PT. CONNECTIVITY LEADS EXCELLENCE",
                normal_text
            )
        ],
    ]

    payment_info_table = Table(
        pay_table_data,
        colWidths=[85, 10, 205]
    )

    payment_info_table.setStyle(
        TableStyle([
            (
                "SPAN",
                (0, 0),
                (2, 0)
            ),
            (
                "BACKGROUND",
                (0, 0),
                (2, 0),
                colors.HexColor("#E2E8F0")
            ),
            (
                "PADDING",
                (0, 0),
                (-1, -1),
                4
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.HexColor("#A0AEC0")
            ),
            (
                "BOX",
                (0, 0),
                (-1, -1),
                1,
                colors.HexColor("#2B6CB0")
            )
        ])
    )

    # --------------------------------------------------------------------------
    # SIGNATURE
    # --------------------------------------------------------------------------
    now_str = datetime.now().strftime(
        "%d %B %Y"
    )

    sig_style = ParagraphStyle(
        name="SigStyle",
        parent=styles["Normal"],
        fontSize=8.5,
        leading=12,
        alignment=1
    )

    sig_block = Paragraph(
        f"""
        Jakarta, {now_str}<br/><br/><br/><br/>
        <b><u>Christian</u></b><br/>
        President Director
        """,
        sig_style
    )

    bottom_grid = Table(
        [[
            payment_info_table,
            sig_block
        ]],
        colWidths=[300, 240]
    )

    bottom_grid.setStyle(
        TableStyle([
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "BOTTOM"
            ),
            (
                "ALIGN",
                (1, 0),
                (1, 0),
                "CENTER"
            ),
            (
                "PADDING",
                (0, 0),
                (-1, -1),
                0
            )
        ])
    )

    elements.append(bottom_grid)

    # --------------------------------------------------------------------------
    # BUILD PDF
    # --------------------------------------------------------------------------
    doc.build(
        elements,
        onFirstPage=draw_header_footer,
        onLaterPages=draw_header_footer
    )

    buffer.seek(0)

    return buffer.getvalue()


# ==============================================================================
# 🖥️ RENDER
# ==============================================================================
def render():

    st.title("📄 Create Invoice")

    st.caption(
        "Modul Pembuatan Invoice - CLX ERP System"
    )

    st.markdown("---")

    # --------------------------------------------------------------------------
    # CONNECTION
    # --------------------------------------------------------------------------
    (
        sheet_query,
        sheet_dropdown,
        sheet_erp,
        sheet_db
    ) = init_gspread()

    if not sheet_query:
        return

    # --------------------------------------------------------------------------
    # LOAD SHEETS
    # --------------------------------------------------------------------------
    raw_query = get_raw_matrix(
        sheet_query
    )

    raw_dropdown = get_raw_matrix(
        sheet_dropdown
    )

    raw_erp = get_raw_matrix(
        sheet_erp
    )

    raw_db = get_raw_matrix(
        sheet_db
    )

    # --------------------------------------------------------------------------
    # AUTO INVOICE NO
    # --------------------------------------------------------------------------
    auto_inv_no = generate_auto_invoice_no(
        sheet_db
    )

    # ==========================================================================
    # PROJECT OPTIONS
    # ==========================================================================
    project_options = []

    if len(raw_dropdown) > 1:

        for row in raw_dropdown[1:]:

            if len(row) > 6 and str(
                row[6]
            ).strip():

                val = str(
                    row[6]
                ).strip()

                if val not in project_options:

                    project_options.append(
                        val
                    )

    if not project_options:

        project_options = [
            "VGreen - Project",
            "VGreen - Operation",
            "SIP",
            "Charge Core"
        ]

    # ==========================================================================
    # FORM UTAMA
    # ==========================================================================
    with st.form(
        "form_create_invoice",
        clear_on_submit=False
    ):

        col1, col2 = st.columns(2)

        # ======================================================================
        # LEFT COLUMN
        # ======================================================================
        with col1:

            selected_project = st.selectbox(
                "Project Name",
                options=project_options
            )

            st.text_input(
                "No. Invoice (Auto Generated)",
                value=auto_inv_no,
                disabled=True
            )

            inv_date = st.date_input(
                "Invoice Date",
                datetime.now()
            )

            # ------------------------------------------------------------------
            # MULTIPLE SITE
            # ------------------------------------------------------------------
            if selected_project == "VGreen - Project":

                site_options = []

                if len(raw_query) > 1:

                    for row in raw_query[1:]:

                        if (
                            len(row) > 5
                            and str(row[5]).strip()
                        ):

                            s_name = str(
                                row[5]
                            ).strip()

                            if (
                                s_name
                                not in site_options
                            ):

                                site_options.append(
                                    s_name
                                )

                selected_sites = st.multiselect(
                    "Item Description / Site Name",
                    options=site_options,
                    placeholder="Pilih satu atau lebih site..."
                    ,
                    help=(
                        "Anda dapat memilih lebih dari "
                        "satu site untuk satu Invoice."
                    )
                )

            else:

                manual_site = st.text_input(
                    "Item Description / Site Name",
                    placeholder=(
                        "Ketik Site Name secara manual..."
                    )
                )

                selected_sites = (
                    [manual_site]
                    if manual_site.strip()
                    else []
                )

        # ======================================================================
        # SITE INFORMATION
        # ======================================================================
        site_details = []

        for site_name in selected_sites:

            info = get_site_information(
                raw_query,
                raw_erp,
                site_name
            )

            site_details.append({
                "site_name": site_name,
                "charging_type": info[
                    "charging_type"
                ],
                "wo_number": info[
                    "wo_number"
                ],
                "boq_amount": info[
                    "boq_amount"
                ]
            })

        # ======================================================================
        # RIGHT COLUMN
        # ======================================================================
        with col2:

            if selected_project == "VGreen - Project":

                if site_details:

                    charging_display = ", ".join(
                        [
                            d["charging_type"]
                            for d in site_details
                            if d["charging_type"]
                        ]
                    )

                    wo_display = ", ".join(
                        [
                            d["wo_number"]
                            for d in site_details
                            if d["wo_number"]
                        ]
                    )

                else:

                    charging_display = ""
                    wo_display = ""

                st.text_input(
                    "Charging Type (Auto)",
                    value=charging_display,
                    disabled=True
                )

                st.text_input(
                    "WO No. (Auto)",
                    value=wo_display,
                    disabled=True
                )

            else:

                st.text_input(
                    "Charging Type (Auto)",
                    value="",
                    disabled=True
                )

                st.text_input(
                    "WO No. (Auto)",
                    value="",
                    disabled=True
                )

            efaktur_no = st.text_input(
                "No. Efaktur",
                placeholder=(
                    "Contoh: 04002600290355647"
                )
            )

        # ==========================================================================
        # 💳 DETAIL PEMBAYARAN
        # ==========================================================================
        st.markdown(
            "### 💳 Detail Pembayaran & Termin"
        )

        # --------------------------------------------------------------------------
        # TOP SCHEMA
        # --------------------------------------------------------------------------
        top_schema_options = {

            "Skema Standar (35% - 60% - 5%)": [
                "35%",
                "60%",
                "5%"
            ],

            "Skema Baru (10% - 85% - 5%)": [
                "10%",
                "85%",
                "5%"
            ]
        }

        selected_schema = st.selectbox(
            "Skema TOP (Terms of Payment)",
            options=list(
                top_schema_options.keys()
            ),
            help=(
                "Pilih skema TOP yang berlaku "
                "sesuai kesepakatan customer."
            )
        )

        all_possible_termins = (
            top_schema_options[
                selected_schema
            ]
        )

        # ==========================================================================
        # 🔒 DETEKSI TERMIN YANG SUDAH TERBIT
        # ==========================================================================
        used_termins_by_site = {}

        if (
            selected_project == "VGreen - Project"
            and len(raw_db) > 1
        ):

            for site in site_details:

                current_site = site[
                    "site_name"
                ].strip().lower()

                current_charge = site[
                    "charging_type"
                ].strip().lower()

                used_for_site = []

                for r in raw_db[1:]:

                    if len(r) <= 6:
                        continue

                    db_charging = str(
                        r[3]
                    ).strip().lower()

                    db_site = str(
                        r[4]
                    ).strip().lower()

                    db_term = str(
                        r[6]
                    ).strip()

                    if (
                        db_site == current_site
                        and db_charging
                        == current_charge
                    ):

                        if db_term:

                            used_for_site.append(
                                db_term
                            )

                used_termins_by_site[
                    current_site
                ] = used_for_site

        # --------------------------------------------------------------------------
        # Termin yang tidak boleh digunakan apabila sudah terpakai pada
        # SALAH SATU site yang dipilih.
        # --------------------------------------------------------------------------
        unavailable_termins = set()

        for used_list in (
            used_termins_by_site.values()
        ):

            unavailable_termins.update(
                used_list
            )

        available_termins = [
            t
            for t in all_possible_termins
            if t not in unavailable_termins
        ]

        # ==========================================================================
        # WARNING
        # ==========================================================================
        if selected_project == "VGreen - Project":

            if selected_sites:

                for site in site_details:

                    site_key = site[
                        "site_name"
                    ].strip().lower()

                    used = (
                        used_termins_by_site
                        .get(site_key, [])
                    )

                    if used:

                        st.warning(
                            f"ℹ️ Termin yang sudah pernah "
                            f"dibuat untuk site "
                            f"**{site['site_name']}** "
                            f"({site['charging_type']}): "
                            f"**{', '.join(set(used))}**"
                        )

            if (
                selected_sites
                and not available_termins
            ):

                st.error(
                    "⛔ Tidak ada termin yang tersedia "
                    "untuk seluruh site yang dipilih."
                )

        # ==========================================================================
        # TERMIN / UNIT PRICE / TAX
        # ==========================================================================
        c1, c2, c3 = st.columns(3)

        with c1:

            if available_termins:

                selected_termin = st.selectbox(
                    "Termin Pembayaran",
                    options=available_termins
                )

                try:

                    pct_val = (
                        float(
                            selected_termin
                            .replace("%", "")
                            .strip()
                        )
                        / 100.0
                    )

                except ValueError:

                    pct_val = 1.0

            else:

                selected_termin = "N/A"

                st.selectbox(
                    "Termin Pembayaran",
                    options=[
                        "Penuh / Lunas"
                    ],
                    disabled=True
                )

                pct_val = 0.0

        # ==========================================================================
        # CALCULATE TOTAL BOQ
        # ==========================================================================
        total_boq_amount = sum(
            float(
                site.get(
                    "boq_amount",
                    0.0
                )
            )
            for site in site_details
        )

        calculated_unit_price = (
            total_boq_amount
            * pct_val
        )

        with c2:

            if selected_project == "VGreen - Project":

                unit_price = st.number_input(
                    "Unit Price (IDR - Auto Calculated)",
                    value=calculated_unit_price,
                    disabled=True,
                    format="%.2f"
                )

            else:

                unit_price = st.number_input(
                    "Unit Price (IDR)",
                    min_value=0.0,
                    value=0.0,
                    step=1000.0
                )

        with c3:

            tax_rate = st.number_input(
                "PPN / Tax (%)",
                min_value=0.0,
                value=11.0,
                step=0.5
            )

        # ==========================================================================
        # AMOUNT
        # ==========================================================================
        termin_amount = unit_price

        tax_amount = (
            termin_amount
            * (tax_rate / 100)
        )

        grand_total = (
            termin_amount
            + tax_amount
        )

        # ==========================================================================
        # SITE BREAKDOWN
        # ==========================================================================
        if site_details:

            st.markdown(
                "#### 📋 Detail Nilai per Site"
            )

            breakdown_rows = []

            for site in site_details:

                site_termin_amount = (
                    site["boq_amount"]
                    * pct_val
                )

                breakdown_rows.append({
                    "Site Name": site[
                        "site_name"
                    ],
                    "Charging Type": site[
                        "charging_type"
                    ],
                    "WO No.": site[
                        "wo_number"
                    ],
                    "BOQ Amount": (
                        site["boq_amount"]
                    ),
                    "Invoice Amount": (
                        site_termin_amount
                    )
                })

            breakdown_df = pd.DataFrame(
                breakdown_rows
            )

            st.dataframe(
                breakdown_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "BOQ Amount": st.column_config.NumberColumn(
                        "BOQ Amount",
                        format="Rp %,.2f"
                    ),
                    "Invoice Amount": st.column_config.NumberColumn(
                        "Invoice Amount",
                        format="Rp %,.2f"
                    )
                }
            )

        # ==========================================================================
        # SUMMARY
        # ==========================================================================
        st.info(
            f"""
**Ringkasan Perhitungan:**

* **Jumlah Site:** {len(site_details)}
* **Skema TOP Terpilih:** {selected_schema}
* **BOQ Amount Base (Sheet ERP):** Rp {total_boq_amount:,.2f}
* **Nilai Invoice ({selected_termin}):** Rp {termin_amount:,.2f}
* **PPN ({tax_rate}%):** Rp {tax_amount:,.2f}
* **GRAND TOTAL:** Rp {grand_total:,.2f}
"""
        )

        # ==========================================================================
        # SUBMIT
        # ==========================================================================
        is_disabled = (
            True
            if (
                selected_project == "VGreen - Project"
                and (
                    not available_termins
                    or not selected_sites
                )
            )
            else False
        )

        submit_btn = st.form_submit_button(
            "💾 Save to DB Invoice",
            type="primary",
            disabled=is_disabled
        )

    # ==========================================================================
    # 💾 SAVE TO DB & GENERATE PDF
    # ==========================================================================
    if submit_btn:

        # --------------------------------------------------------------------------
        # VALIDATION
        # --------------------------------------------------------------------------
        if not efaktur_no:

            st.error(
                "⚠️ No. Efaktur wajib diisi!"
            )

            return

        if not selected_sites:

            st.error(
                "⚠️ Minimal pilih 1 Site!"
            )

            return

        # --------------------------------------------------------------------------
        # VALIDATE MULTIPLE SITE
        # --------------------------------------------------------------------------
        if selected_project == "VGreen - Project":

            if not site_details:

                st.error(
                    "⚠️ Data site tidak ditemukan."
                )

                return

            # Pastikan semua site mempunyai BOQ
            missing_boq = [
                site["site_name"]
                for site in site_details
                if site["boq_amount"] <= 0
            ]

            if missing_boq:

                st.warning(
                    "⚠️ BOQ Amount tidak ditemukan "
                    "untuk site berikut:\n\n"
                    + "\n".join(
                        [
                            f"- {x}"
                            for x in missing_boq
                        ]
                    )
                )

                return

        # ==========================================================================
        # CREATE SITE ROWS
        # ==========================================================================
        pdf_site_rows = []

        db_rows = []

        for site in site_details:

            site_boq = float(
                site.get(
                    "boq_amount",
                    0.0
                )
            )

            # Untuk VGreen Project:
            # Site Amount = BOQ × Termin %
            #
            # Untuk project manual:
            # karena sebelumnya Unit Price adalah nilai invoice,
            # maka seluruh nilai diberikan ke site tersebut.
            if selected_project == "VGreen - Project":

                site_invoice_amount = (
                    site_boq
                    * pct_val
                )

            else:

                site_invoice_amount = (
                    unit_price
                )

            site_tax_amount = (
                site_invoice_amount
                * (tax_rate / 100)
            )

            site_grand_total = (
                site_invoice_amount
                + site_tax_amount
            )

            # ----------------------------------------------------------------------
            # PDF DATA
            # ----------------------------------------------------------------------
            pdf_site_rows.append({
                "site_name": site[
                    "site_name"
                ],
                "charging_type": site[
                    "charging_type"
                ],
                "wo_number": site[
                    "wo_number"
                ],
                "termin": selected_termin,
                "boq_amount": site_boq,
                "termin_amount": site_invoice_amount,
                "tax_amount": site_tax_amount,
                "grand_total": site_grand_total
            })

            # ----------------------------------------------------------------------
            # DB ROW
            #
            # Struktur tetap:
            # Project Name
            # Invoice No.
            # Invoice Date
            # Charging Type
            # Site Name
            # WO No.
            # Termin
            # Efaktur
            # Subtotal
            # Grand Total
            # ----------------------------------------------------------------------
            db_rows.append([
                selected_project,
                auto_inv_no,
                inv_date.strftime(
                    "%d %B %Y"
                ),
                site[
                    "charging_type"
                ],
                site[
                    "site_name"
                ],
                site[
                    "wo_number"
                ],
                selected_termin,
                efaktur_no,
                f"{site_invoice_amount:.2f}",
                f"{site_grand_total:.2f}"
            ])

        # ==========================================================================
        # SAVE TO GOOGLE SHEETS
        # ==========================================================================
        if sheet_db:

            try:

                # ------------------------------------------------------------------
                # Append setiap site sebagai 1 row
                # ------------------------------------------------------------------
                for row in db_rows:

                    sheet_db.append_row(
                        row,
                        value_input_option="USER_ENTERED"
                    )

                # ------------------------------------------------------------------
                # SUCCESS
                # ------------------------------------------------------------------
                st.success(
                    f"✅ Invoice **{auto_inv_no}** "
                    f"berhasil disimpan ke Sheet "
                    f"'DB Invoice'!"
                )

                st.success(
                    f"📍 {len(db_rows)} site "
                    f"berhasil dimasukkan dalam "
                    f"1 Invoice."
                )

                st.balloons()

                # ------------------------------------------------------------------
                # PDF PAYLOAD
                # ------------------------------------------------------------------
                pdf_payload = {

                    "project_name":
                        selected_project,

                    "inv_no":
                        auto_inv_no,

                    "inv_date":
                        inv_date.strftime(
                            "%d %B %Y"
                        ),

                    "efaktur":
                        efaktur_no,

                    "termin":
                        selected_termin,

                    "site_rows":
                        pdf_site_rows,

                    "subtotal":
                        termin_amount,

                    "tax_rate":
                        tax_rate,

                    "tax_amount":
                        tax_amount,

                    "grand_total":
                        grand_total
                }

                # ------------------------------------------------------------------
                # CREATE PDF
                # ------------------------------------------------------------------
                pdf_file_bytes = (
                    create_invoice_pdf(
                        pdf_payload
                    )
                )

                st.session_state[
                    "pdf_ready"
                ] = pdf_file_bytes

                st.session_state[
                    "pdf_name"
                ] = (
                    f"Invoice_"
                    f"{auto_inv_no.replace('/', '_')}"
                    f".pdf"
                )

                # ------------------------------------------------------------------
                # CLEAR CACHE / REFRESH
                # ------------------------------------------------------------------
                st.cache_resource.clear()

            except Exception as e:

                st.error(
                    "🚨 Gagal menyimpan ke Google Sheets: "
                    f"{e}"
                )

        else:

            st.error(
                "🚨 Tidak terhubung ke Google Sheets."
            )

    # ==========================================================================
    # 📥 DOWNLOAD PDF
    # ==========================================================================
    if (
        "pdf_ready"
        in st.session_state
        and st.session_state[
            "pdf_ready"
        ]
    ):

        st.markdown("---")

        st.subheader(
            "📥 Download Berkas Invoice"
        )

        st.download_button(
            label="📄 Download Invoice PDF",
            data=st.session_state[
                "pdf_ready"
            ],
            file_name=st.session_state[
                "pdf_name"
            ],
            mime="application/pdf",
            type="primary",
            key="dl_btn_invoice"
        )
