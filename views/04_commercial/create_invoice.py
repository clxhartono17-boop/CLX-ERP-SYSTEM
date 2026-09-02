import os
import io
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
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


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
    12: "XII",
}


# ==============================================================================
# GOOGLE SHEETS CONNECTION
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
            sheet_db,
        )

    except Exception as e:

        st.error(
            f"🚨 Gagal terhubung ke Google Sheets: {e}"
        )

        return None, None, None, None


# ==============================================================================
# GET RAW MATRIX
# ==============================================================================
def get_raw_matrix(sheet):
    """
    Membaca seluruh isi sheet menjadi 2D List.
    """

    try:
        return sheet.get_all_values()

    except Exception:
        return []


# ==============================================================================
# GENERATE AUTO INVOICE NUMBER
# ==============================================================================
def generate_auto_invoice_no(sheet_db):
    """
    Generate No. Invoice Otomatis

    Format:
    0000/INV/CLX/BULAN_ROMAWI/TAHUN

    Nomor awal minimal 462.
    """

    now = datetime.now()

    roman_month = MONTH_ROMAN.get(
        now.month,
        "I"
    )

    year = now.year

    if sheet_db:

        try:

            records = sheet_db.get_all_records()

            next_seq = max(
                len(records) + 1,
                462
            )

        except Exception:

            next_seq = 462

    else:

        next_seq = 462

    formatted_seq = f"{next_seq:04d}"

    return (
        f"{formatted_seq}/INV/CLX/"
        f"{roman_month}/{year}"
    )


# ==============================================================================
# SAFE FLOAT
# ==============================================================================
def safe_float(value):
    """
    Mengubah berbagai format angka menjadi float.
    """

    if value is None:
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    value = str(value).strip()

    if not value:
        return 0.0

    try:

        # Format Indonesia:
        # 47.562.500
        if "." in value and "," not in value:

            value = (
                value
                .replace("Rp", "")
                .replace(".", "")
                .replace(" ", "")
                .strip()
            )

            return float(value)

        # Format:
        # 47.562.500,50
        value = (
            value
            .replace("Rp", "")
            .replace(".", "")
            .replace(",", ".")
            .replace(" ", "")
            .strip()
        )

        return float(value)

    except Exception:

        return 0.0


# ==============================================================================
# GET SITE DATA
# ==============================================================================
def get_site_data(
    raw_query,
    raw_erp,
    site_name
):
    """
    Mengambil data site dari Query dan ERP Project.

    Query:
        Charger Type -> Column C / index 2
        Site Name    -> Column F / index 5
        WO Number    -> Column W / index 22

    ERP Project:
        Site Name    -> Column E / index 4
        Charger Type -> Column C / index 2
        BOQ Amount   -> Column M / index 12
    """

    result = {
        "site_name": site_name,
        "charging_type": "",
        "wo_number": "",
        "boq_amount": 0.0,
        "qty": 1,
        "uom": "Unit",
    }

    # --------------------------------------------------------------------------
    # QUERY
    # --------------------------------------------------------------------------
    if len(raw_query) > 1:

        target_site = site_name.strip().lower()

        for row in raw_query[1:]:

            if len(row) > 5:

                current_site = (
                    row[5]
                    .strip()
                    .lower()
                )

                if current_site == target_site:

                    # Charging Type
                    if len(row) > 2:
                        result["charging_type"] = (
                            row[2].strip()
                        )

                    # WO Number
                    if len(row) > 22:
                        result["wo_number"] = (
                            row[22].strip()
                        )

                    break

    # --------------------------------------------------------------------------
    # ERP PROJECT
    # --------------------------------------------------------------------------
    if len(raw_erp) > 1:

        target_site = site_name.strip().lower()

        target_charge = (
            result["charging_type"]
            .strip()
            .lower()
        )

        for row in raw_erp[1:]:

            if len(row) > 4:

                erp_site = (
                    row[4]
                    .strip()
                    .lower()
                )

                erp_charge = (
                    row[2].strip().lower()
                    if len(row) > 2
                    else ""
                )

                if (
                    erp_site == target_site
                    and (
                        not target_charge
                        or erp_charge == target_charge
                    )
                ):

                    if len(row) > 12:

                        result["boq_amount"] = safe_float(
                            row[12]
                        )

                    break

    return result


# ==============================================================================
# CREATE MULTI SITE INVOICE PDF
# ==============================================================================
def create_invoice_pdf(data):
    """
    Fungsi pembuat PDF Invoice.

    Mendukung multiple site.
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

    elements = []

    styles = getSampleStyleSheet()

    # --------------------------------------------------------------------------
    # STYLES
    # --------------------------------------------------------------------------
    title_style = ParagraphStyle(
        name="TitleStyle",
        parent=styles["Heading1"],
        fontSize=22,
        leading=26,
        textColor=colors.HexColor("#1A365D"),
        alignment=2,
    )

    normal_bold = ParagraphStyle(
        name="NormalBold",
        parent=styles["Normal"],
        fontSize=8.5,
        leading=11,
        fontName="Helvetica-Bold",
    )

    normal_text = ParagraphStyle(
        name="NormalText",
        parent=styles["Normal"],
        fontSize=8.5,
        leading=11,
    )

    small_text = ParagraphStyle(
        name="SmallText",
        parent=styles["Normal"],
        fontSize=7.5,
        leading=9,
    )

    pay_title_style = ParagraphStyle(
        name="PayTitle",
        parent=styles["Normal"],
        fontSize=9,
        leading=11,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#1A365D"),
    )

    # --------------------------------------------------------------------------
    # HEADER & FOOTER
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
                    mask="auto",
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
                    mask="auto",
                )

            except Exception:
                pass

        canvas.restoreState()

    # --------------------------------------------------------------------------
    # TITLE
    # --------------------------------------------------------------------------
    elements.append(
        Paragraph(
            "<b>INVOICE</b>",
            title_style
        )
    )

    elements.append(
        Spacer(1, 10)
    )

    # --------------------------------------------------------------------------
    # BILL TO
    # --------------------------------------------------------------------------
    bill_to_text = Paragraph(
        """
        <b>Bill To :</b><br/>
        <b>PT Vgreen Global Charging Station Investment Indonesia</b><br/>
        Graha Binakarsa Lt.7, Jl. H.R Rasuna Said Kav C-18
        RT 02 RW 005,<br/>
        Karet Kuningan, Kec. Setiabudi Jakarta Selatan
        """,
        normal_text,
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
            ),
        ],

        [
            Paragraph(
                "<b>Invoice Date</b>",
                normal_bold
            ),
            Paragraph(
                f": {data['inv_date']}",
                normal_text
            ),
        ],

        [
            Paragraph(
                "<b>Project Name</b>",
                normal_bold
            ),
            Paragraph(
                f": {data['project_name']}",
                normal_text
            ),
        ],

        [
            Paragraph(
                "<b>No. Efaktur</b>",
                normal_bold
            ),
            Paragraph(
                f": {data['efaktur']}",
                normal_text
            ),
        ],
    ]

    meta_table = Table(
        issuer_and_meta,
        colWidths=[80, 180]
    )

    meta_table.setStyle(
        TableStyle(
            [
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
                ),
            ]
        )
    )

    info_grid = Table(
        [[bill_to_text, meta_table]],
        colWidths=[280, 260]
    )

    info_grid.setStyle(
        TableStyle(
            [
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
                ),
            ]
        )
    )

    elements.append(info_grid)

    elements.append(
        Spacer(1, 15)
    )

    # --------------------------------------------------------------------------
    # MULTI SITE ITEM TABLE
    # --------------------------------------------------------------------------
    table_data = [

        [
            Paragraph(
                "<b>NO.</b>",
                small_text
            ),

            Paragraph(
                "<b>ITEM DESCRIPTION</b>",
                small_text
            ),

            Paragraph(
                "<b>QTY</b>",
                small_text
            ),

            Paragraph(
                "<b>UOM</b>",
                small_text
            ),

            Paragraph(
                "<b>CHARGER TYPE</b>",
                small_text
            ),

            Paragraph(
                "<b>WO NUMBER</b>",
                small_text
            ),

            Paragraph(
                "<b>UNIT PRICE</b>",
                small_text
            ),

            Paragraph(
                "<b>AMOUNT</b>",
                small_text
            ),
        ]
    ]

    for idx, item in enumerate(
        data["items"],
        start=1
    ):

        table_data.append(
            [
                str(idx),

                Paragraph(
                    str(item["site_name"]),
                    small_text
                ),

                str(item.get("qty", 1)),

                str(item.get("uom", "Unit")),

                Paragraph(
                    str(item["charging_type"]),
                    small_text
                ),

                Paragraph(
                    str(item["wo_number"]),
                    small_text
                ),

                f"{item['unit_price']:,.2f}",

                f"{item['amount']:,.2f}",
            ]
        )

    item_table = Table(
        table_data,
        colWidths=[
            25,
            145,
            30,
            35,
            60,
            75,
            75,
            75,
        ],
        repeatRows=1,
    )

    item_table.setStyle(
        TableStyle(
            [
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
                    (2, 1),
                    (3, -1),
                    "CENTER"
                ),

                (
                    "ALIGN",
                    (6, 1),
                    (7, -1),
                    "RIGHT"
                ),

                (
                    "PADDING",
                    (0, 0),
                    (-1, -1),
                    4
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
                ),
            ]
        )
    )

    elements.append(item_table)

    elements.append(
        Spacer(1, 10)
    )

    # --------------------------------------------------------------------------
    # SUMMARY TOTALS
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
        ],
    ]

    summary_table = Table(
        summary_data,
        colWidths=[380, 160]
    )

    summary_table.setStyle(
        TableStyle(
            [
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
                ),
            ]
        )
    )

    elements.append(summary_table)

    elements.append(
        Spacer(1, 15)
    )

    # --------------------------------------------------------------------------
    # PAYMENT INFO
    # --------------------------------------------------------------------------
    pay_table_data = [

        [
            Paragraph(
                "<b>PAYMENT INFO</b>",
                pay_title_style
            ),
            "",
            "",
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
            ),
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
            ),
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
            ),
        ],
    ]

    payment_info_table = Table(
        pay_table_data,
        colWidths=[85, 10, 205]
    )

    payment_info_table.setStyle(
        TableStyle(
            [
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
                ),
            ]
        )
    )

    # --------------------------------------------------------------------------
    # SIGNATURE
    # --------------------------------------------------------------------------
    now_str = datetime.now().strftime(
        "%d %B %Y"
    )

    sig_block = Paragraph(
        f"""
        Jakarta, {now_str}<br/><br/><br/><br/>
        <b><u>Christian</u></b><br/>
        President Director
        """,
        ParagraphStyle(
            name="SigStyle",
            parent=styles["Normal"],
            fontSize=8.5,
            leading=12,
            alignment=1,
        )
    )

    bottom_grid = Table(
        [[payment_info_table, sig_block]],
        colWidths=[300, 240]
    )

    bottom_grid.setStyle(
        TableStyle(
            [
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
                ),
            ]
        )
    )

    elements.append(bottom_grid)

    # --------------------------------------------------------------------------
    # BUILD PDF
    # --------------------------------------------------------------------------
    doc.build(
        elements,
        onFirstPage=draw_header_footer,
        onLaterPages=draw_header_footer,
    )

    buffer.seek(0)

    return buffer.getvalue()


# ==============================================================================
# RENDER
# ==============================================================================
def render():

    st.title("📄 Create Invoice")

    st.caption(
        "Modul Pembuatan Invoice - CLX ERP System"
    )

    st.markdown("---")

    # ==========================================================================
    # CONNECT GOOGLE SHEETS
    # ==========================================================================
    (
        sheet_query,
        sheet_dropdown,
        sheet_erp,
        sheet_db,
    ) = init_gspread()

    if not sheet_query:
        return

    # ==========================================================================
    # LOAD SHEETS
    # ==========================================================================
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

    # ==========================================================================
    # AUTO INVOICE NUMBER
    # ==========================================================================
    auto_inv_no = generate_auto_invoice_no(
        sheet_db
    )

    # ==========================================================================
    # PROJECT OPTIONS
    # ==========================================================================
    project_options = []

    if len(raw_dropdown) > 1:

        for row in raw_dropdown[1:]:

            if len(row) > 6:

                if row[6].strip():

                    val = row[6].strip()

                    if val not in project_options:

                        project_options.append(val)

    if not project_options:

        project_options = [
            "VGreen - Project",
            "VGreen - Operation",
            "SIP",
            "Charge Core",
        ]

    # ==========================================================================
    # PROJECT NAME
    #
    # PENTING:
    # Diletakkan DI LUAR FORM agar perubahan selection
    # langsung melakukan rerun.
    # ==========================================================================
    selected_project = st.selectbox(
        "Project Name",
        options=project_options,
        key="invoice_project_name",
    )

    # ==========================================================================
    # SITE SELECTION
    # ==========================================================================
    selected_sites = []

    if selected_project == "VGreen - Project":

        site_options = []

        if len(raw_query) > 1:

            for row in raw_query[1:]:

                if len(row) > 5:

                    s_name = row[5].strip()

                    if s_name:

                        if s_name not in site_options:

                            site_options.append(
                                s_name
                            )

        selected_sites = st.multiselect(
            "Item Description / Site Name",
            options=site_options,
            placeholder="Pilih satu atau beberapa site...",
            key="invoice_selected_sites",
            help=(
                "Anda dapat memilih lebih dari satu site. "
                "Setelah site dipilih, tabel detail akan "
                "langsung muncul di bawah."
            ),
        )

    else:

        manual_site = st.text_input(
            "Item Description / Site Name",
            placeholder=(
                "Ketik Site Name secara manual..."
            ),
            key="invoice_manual_site",
        )

        if manual_site.strip():

            selected_sites = [
                manual_site.strip()
            ]

    # ==========================================================================
    # BUILD SITE DATA
    # ==========================================================================
    site_data_list = []

    for site in selected_sites:

        item = get_site_data(
            raw_query,
            raw_erp,
            site,
        )

        site_data_list.append(item)

    # ==========================================================================
    # PREVIEW TABLE
    # ==========================================================================
    if selected_project == "VGreen - Project":

        st.markdown("---")

        st.subheader(
            "🔎 Preview Detail Site"
        )

        if site_data_list:

            preview_rows = []

            for idx, item in enumerate(
                site_data_list,
                start=1
            ):

                preview_rows.append(
                    {
                        "NO.": idx,
                        "ITEM DESCRIPTION": item[
                            "site_name"
                        ],
                        "QTY": item.get(
                            "qty",
                            1
                        ),
                        "UOM": item.get(
                            "uom",
                            "Unit"
                        ),
                        "CHARGER TYPE": item[
                            "charging_type"
                        ],
                        "WO NUMBER": item[
                            "wo_number"
                        ],
                        "BOQ AMOUNT": item[
                            "boq_amount"
                        ],
                    }
                )

            preview_df = pd.DataFrame(
                preview_rows
            )

            # ------------------------------------------------------------------
            # FORMAT CURRENCY UNTUK DISPLAY
            # ------------------------------------------------------------------
            display_df = preview_df.copy()

            display_df[
                "BOQ AMOUNT"
            ] = display_df[
                "BOQ AMOUNT"
            ].apply(
                lambda x: f"Rp {x:,.2f}"
            )

            st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True,
            )

            st.caption(
                f"📌 Total site dipilih: "
                f"**{len(site_data_list)} site**"
            )

        else:

            st.info(
                "👆 Silakan pilih satu atau beberapa site "
                "untuk melihat detail dan nominalnya."
            )

    # ==========================================================================
    # FORM DETAIL INVOICE
    # ==========================================================================
    with st.form(
        "form_create_invoice",
        clear_on_submit=False
    ):

        # ----------------------------------------------------------------------
        # INFO INVOICE
        # ----------------------------------------------------------------------
        col1, col2 = st.columns(2)

        with col1:

            st.text_input(
                "No. Invoice (Auto Generated)",
                value=auto_inv_no,
                disabled=True,
            )

            inv_date = st.date_input(
                "Invoice Date",
                datetime.now(),
            )

        with col2:

            # ------------------------------------------------------------------
            # CHARGING TYPE
            # ------------------------------------------------------------------
            if (
                selected_project == "VGreen - Project"
                and site_data_list
            ):

                charging_types = []

                for item in site_data_list:

                    ct = item[
                        "charging_type"
                    ]

                    if ct not in charging_types:

                        charging_types.append(
                            ct
                        )

                charging_type_display = (
                    ", ".join(charging_types)
                )

            else:

                charging_type_display = ""

            st.text_input(
                "Charging Type (Auto)",
                value=charging_type_display,
                disabled=True,
            )

            # ------------------------------------------------------------------
            # WO
            # ------------------------------------------------------------------
            if (
                selected_project == "VGreen - Project"
                and site_data_list
            ):

                wo_numbers = []

                for item in site_data_list:

                    wo = item[
                        "wo_number"
                    ]

                    if wo and wo not in wo_numbers:

                        wo_numbers.append(
                            wo
                        )

                wo_display = ", ".join(
                    wo_numbers
                )

            else:

                wo_display = ""

            st.text_input(
                "WO No. (Auto)",
                value=wo_display,
                disabled=True,
            )

            efaktur_no = st.text_input(
                "No. Efaktur",
                placeholder=(
                    "Contoh: 04002600290355647"
                ),
            )

        # ======================================================================
        # DETAIL PEMBAYARAN
        # ======================================================================
        st.markdown(
            "### 💳 Detail Pembayaran & Termin"
        )

        # ----------------------------------------------------------------------
        # TOP SCHEMA
        # ----------------------------------------------------------------------
        top_schema_options = {

            "Skema Standar (35% - 60% - 5%)": [
                "35%",
                "60%",
                "5%",
            ],

            "Skema Baru (10% - 85% - 5%)": [
                "10%",
                "85%",
                "5%",
            ],
        }

        selected_schema = st.selectbox(
            "Skema TOP (Terms of Payment)",
            options=list(
                top_schema_options.keys()
            ),
            help=(
                "Pilih skema TOP yang berlaku "
                "sesuai kesepakatan customer"
            ),
        )

        all_possible_termins = (
            top_schema_options[
                selected_schema
            ]
        )

        # ======================================================================
        # DETEKSI TERMIN YANG SUDAH TERPAKAI
        # ======================================================================
        used_termins_by_site = {}

        if (
            selected_project == "VGreen - Project"
            and len(raw_db) > 1
        ):

            for site in selected_sites:

                used_termins_by_site[
                    site.strip().lower()
                ] = []

            for r in raw_db[1:]:

                if len(r) > 6:

                    db_charging = (
                        r[3]
                        .strip()
                        .lower()
                    )

                    db_site = (
                        r[4]
                        .strip()
                        .lower()
                    )

                    db_term = (
                        r[6].strip()
                    )

                    if (
                        db_site
                        in used_termins_by_site
                    ):

                        site_info = get_site_data(
                            raw_query,
                            raw_erp,
                            db_site,
                        )

                        current_charge = (
                            site_info[
                                "charging_type"
                            ]
                            .strip()
                            .lower()
                        )

                        if (
                            current_charge
                            == db_charging
                        ):

                            used_termins_by_site[
                                db_site
                            ].append(
                                db_term
                            )

        # ======================================================================
        # AVAILABLE TERMIN
        # ======================================================================
        available_termins = (
            all_possible_termins.copy()
        )

        if selected_project == "VGreen - Project":

            # Sebuah termin hanya tersedia apabila
            # termin tersebut belum digunakan oleh
            # SEMUA site yang dipilih.
            #
            # Jika salah satu site sudah menggunakan termin,
            # maka termin tidak dapat digunakan untuk invoice
            # gabungan tersebut.

            for site in selected_sites:

                used_for_site = (
                    used_termins_by_site.get(
                        site.strip().lower(),
                        []
                    )
                )

                available_termins = [
                    t
                    for t in available_termins
                    if t not in used_for_site
                ]

            # --------------------------------------------------------------
            # WARNING
            # --------------------------------------------------------------
            warning_parts = []

            for site in selected_sites:

                used_for_site = (
                    used_termins_by_site.get(
                        site.strip().lower(),
                        []
                    )
                )

                if used_for_site:

                    warning_parts.append(
                        f"**{site}**: "
                        f"{', '.join(set(used_for_site))}"
                    )

            if warning_parts:

                st.warning(
                    "ℹ️ Termin yang sudah pernah "
                    "dibuat:\n\n"
                    + "\n\n".join(
                        warning_parts
                    )
                )

            if (
                selected_sites
                and not available_termins
            ):

                st.error(
                    "⛔ Tidak ada termin yang tersedia "
                    "untuk seluruh site yang dipilih."
                )

        # ======================================================================
        # TERMIN + TAX
        # ======================================================================
        c1, c2, c3 = st.columns(3)

        with c1:

            if available_termins:

                selected_termin = st.selectbox(
                    "Termin Pembayaran",
                    options=available_termins,
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
                    disabled=True,
                )

                pct_val = 0.0

        # ======================================================================
        # CALCULATE PER SITE
        # ======================================================================
        calculated_items = []

        if selected_project == "VGreen - Project":

            for item in site_data_list:

                unit_price = (
                    item["boq_amount"]
                    * pct_val
                )

                calculated_item = {
                    "site_name": item[
                        "site_name"
                    ],

                    "charging_type": item[
                        "charging_type"
                    ],

                    "wo_number": item[
                        "wo_number"
                    ],

                    "boq_amount": item[
                        "boq_amount"
                    ],

                    "qty": item.get(
                        "qty",
                        1
                    ),

                    "uom": item.get(
                        "uom",
                        "Unit"
                    ),

                    "unit_price": unit_price,

                    "amount": unit_price,
                }

                calculated_items.append(
                    calculated_item
                )

        else:

            manual_unit_price = 0.0

            calculated_items = [
                {
                    "site_name": (
                        selected_sites[0]
                        if selected_sites
                        else ""
                    ),

                    "charging_type": "",

                    "wo_number": "",

                    "boq_amount": 0.0,

                    "qty": 1,

                    "uom": "Unit",

                    "unit_price": 0.0,

                    "amount": 0.0,
                }
            ]

        # ======================================================================
        # UNIT PRICE
        # ======================================================================
        with c2:

            if (
                selected_project == "VGreen - Project"
            ):

                total_boq = sum(
                    item["boq_amount"]
                    for item in site_data_list
                )

                total_unit_price = sum(
                    item["unit_price"]
                    for item in calculated_items
                )

                st.number_input(
                    "Unit Price Total "
                    "(IDR - Auto Calculated)",
                    value=float(
                        total_unit_price
                    ),
                    disabled=True,
                    format="%.2f",
                )

            else:

                manual_unit_price = (
                    st.number_input(
                        "Unit Price (IDR)",
                        min_value=0.0,
                        value=0.0,
                        step=1000.0,
                    )
                )

                if calculated_items:

                    calculated_items[0][
                        "unit_price"
                    ] = manual_unit_price

                    calculated_items[0][
                        "amount"
                    ] = manual_unit_price

        # ======================================================================
        # TAX
        # ======================================================================
        with c3:

            tax_rate = st.number_input(
                "PPN / Tax (%)",
                min_value=0.0,
                value=11.0,
                step=0.5,
            )

        # ======================================================================
        # TOTAL
        # ======================================================================
        if selected_project == "VGreen - Project":

            termin_amount = sum(
                item["amount"]
                for item in calculated_items
            )

        else:

            termin_amount = (
                manual_unit_price
                if calculated_items
                else 0.0
            )

        tax_amount = (
            termin_amount
            * (tax_rate / 100)
        )

        grand_total = (
            termin_amount
            + tax_amount
        )

        # ======================================================================
        # DETAIL CROSS CHECK TABLE
        # ======================================================================
        if (
            selected_project == "VGreen - Project"
            and calculated_items
        ):

            st.markdown(
                "### 💰 Detail Nilai Invoice"
            )

            invoice_preview_rows = []

            for idx, item in enumerate(
                calculated_items,
                start=1
            ):

                invoice_preview_rows.append(
                    {
                        "NO.": idx,

                        "ITEM DESCRIPTION": item[
                            "site_name"
                        ],

                        "QTY": item[
                            "qty"
                        ],

                        "UOM": item[
                            "uom"
                        ],

                        "CHARGER TYPE": item[
                            "charging_type"
                        ],

                        "WO NUMBER": item[
                            "wo_number"
                        ],

                        "UNIT PRICE": item[
                            "unit_price"
                        ],

                        "AMOUNT": item[
                            "amount"
                        ],
                    }
                )

            invoice_preview_df = pd.DataFrame(
                invoice_preview_rows
            )

            display_invoice_df = (
                invoice_preview_df.copy()
            )

            display_invoice_df[
                "UNIT PRICE"
            ] = display_invoice_df[
                "UNIT PRICE"
            ].apply(
                lambda x: f"{x:,.0f}"
            )

            display_invoice_df[
                "AMOUNT"
            ] = display_invoice_df[
                "AMOUNT"
            ].apply(
                lambda x: f"{x:,.0f}"
            )

            st.dataframe(
                display_invoice_df,
                use_container_width=True,
                hide_index=True,
            )

        # ======================================================================
        # SUMMARY
        # ======================================================================
        st.info(
            f"""
**Ringkasan Perhitungan:**

* **Skema TOP Terpilih:** {selected_schema}
* **Jumlah Site:** {len(calculated_items)}
* **BOQ Amount Total:** Rp {sum(item['boq_amount'] for item in calculated_items):,.2f}
* **Nilai Invoice ({selected_termin}):** Rp {termin_amount:,.2f}
* **PPN ({tax_rate}%):** Rp {tax_amount:,.2f}
* **GRAND TOTAL:** Rp {grand_total:,.2f}
"""
        )

        # ======================================================================
        # SUBMIT
        # ======================================================================
        is_disabled = (
            True
            if (
                selected_project
                == "VGreen - Project"
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
            disabled=is_disabled,
        )

    # ==========================================================================
    # SAVE TO DB
    # ==========================================================================
    if submit_btn:

        if not efaktur_no:

            st.error(
                "⚠️ No. Efaktur wajib diisi!"
            )

        elif not selected_sites:

            st.error(
                "⚠️ Site Name / Item Description "
                "wajib diisi!"
            )

        else:

            try:

                # =================================================================
                # SAVE SATU ROW UNTUK SETIAP SITE
                # =================================================================
                if sheet_db:

                    for item in calculated_items:

                        new_row = [

                            selected_project,

                            auto_inv_no,

                            inv_date.strftime(
                                "%d %B %Y"
                            ),

                            item[
                                "charging_type"
                            ],

                            item[
                                "site_name"
                            ],

                            item[
                                "wo_number"
                            ],

                            selected_termin,

                            efaktur_no,

                            f"{item['amount']:.2f}",

                            f"{item['amount'] + (item['amount'] * tax_rate / 100):.2f}",
                        ]

                        sheet_db.append_row(
                            new_row
                        )

                    # =============================================================
                    # SUCCESS
                    # =============================================================
                    st.success(
                        f"✅ Invoice **{auto_inv_no}** "
                        f"berhasil disimpan ke Sheet "
                        f"'DB Invoice'!"
                    )

                    st.success(
                        f"📍 Total **{len(calculated_items)} site** "
                        f"berhasil disimpan."
                    )

                    st.balloons()

                    # =============================================================
                    # PDF PAYLOAD
                    # =============================================================
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

                        "items":
                            calculated_items,

                        "subtotal":
                            termin_amount,

                        "tax_rate":
                            tax_rate,

                        "tax_amount":
                            tax_amount,

                        "grand_total":
                            grand_total,
                    }

                    # =============================================================
                    # GENERATE PDF
                    # =============================================================
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
                        "Invoice_"
                        + auto_inv_no.replace(
                            "/",
                            "_"
                        )
                        + ".pdf"
                    )

                else:

                    st.error(
                        "🚨 Tidak terhubung ke "
                        "Google Sheets."
                    )

            except Exception as e:

                st.error(
                    f"🚨 Gagal menyimpan ke "
                    f"Google Sheets: {e}"
                )

    # ==========================================================================
    # DOWNLOAD PDF
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

            key="dl_btn_invoice",
        )
