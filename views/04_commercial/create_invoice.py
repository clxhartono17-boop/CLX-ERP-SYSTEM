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
# 🔌 GOOGLE SHEETS CONNECTION
# ==============================================================================

@st.cache_resource
def init_gspread():
    """
    Inisialisasi koneksi Google Sheets.

    Mendukung:
    1. Streamlit Cloud Secrets
    2. credentials.json untuk localhost
    """

    gc = None

    # --------------------------------------------------------------------------
    # 1. STREAMLIT SECRETS
    # --------------------------------------------------------------------------
    if "gcp_service_account" in st.secrets:

        try:

            creds_dict = dict(
                st.secrets["gcp_service_account"]
            )

            gc = gspread.service_account_from_dict(
                creds_dict
            )

        except Exception as e:

            st.warning(
                f"⚠️ Gagal menghubungkan Secrets Cloud: {e}"
            )

    # --------------------------------------------------------------------------
    # 2. LOCAL CREDENTIALS
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

                return (
                    None,
                    None,
                    None,
                    None,
                    None,
                )

        else:

            st.error(
                "🚨 Kredensial tidak ditemukan! "
                "Harap set 'Secrets' di Streamlit Cloud "
                "atau sediakan file 'credentials.json' "
                "di localhost."
            )

            return (
                None,
                None,
                None,
                None,
                None,
            )

    # --------------------------------------------------------------------------
    # 3. OPEN GOOGLE SPREADSHEET
    # --------------------------------------------------------------------------
    try:

        sh = gc.open_by_key(
            SPREADSHEET_ID
        )

        sheet_query = sh.worksheet(
            "Query"
        )

        sheet_dropdown = sh.worksheet(
            "Master Dropdown"
        )

        sheet_erp = sh.worksheet(
            "ERP Project"
        )

        sheet_db = sh.worksheet(
            "DB Invoice"
        )

        # ==========================================================================
        # ⭐ DB BOQ
        #
        # Struktur:
        # A = No
        # B = BOQ No.
        # C = Site Name
        # D = Charger Type
        # E = BOQ Amount Exc. PPN
        # F = BOQ Amount inc. PPN
        # G = EPC Name
        # ==========================================================================

        sheet_db_boq = sh.worksheet(
            "DB BOQ"
        )

        return (
            sheet_query,
            sheet_dropdown,
            sheet_erp,
            sheet_db,
            sheet_db_boq,
        )

    except Exception as e:

        st.error(
            f"🚨 Gagal terhubung ke Google Sheets: {e}"
        )

        return (
            None,
            None,
            None,
            None,
            None,
        )


# ==============================================================================
# 📥 GET RAW MATRIX
# ==============================================================================

def get_raw_matrix(sheet):
    """
    Membaca seluruh isi sheet menjadi 2D List.
    """

    try:

        if sheet is None:
            return []

        return sheet.get_all_values()

    except Exception:

        return []


# ==============================================================================
# 🔢 GENERATE INVOICE NUMBER
# ==============================================================================

def generate_auto_invoice_no(sheet_db):
    """
    Generate No. Invoice Otomatis.

    Format:
    0000/INV/CLX/BULAN_ROMAWI/TAHUN

    Minimum sequence:
    462
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
# 💰 PARSE CURRENCY / NUMBER
# ==============================================================================

def parse_amount(value):
    """
    Mengubah berbagai format nominal menjadi float.

    Contoh:

    47.562.500
    Rp 47.562.500
    47,562,500
    47562500
    47.562.500,00

    menjadi:

    47562500.0
    """

    if value is None:
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()

    if not text:
        return 0.0

    # Hilangkan Rp
    text = (
        text
        .replace("Rp", "")
        .replace("rp", "")
        .replace(" ", "")
    )

    # --------------------------------------------------------------------------
    # Jika terdapat titik dan koma:
    #
    # 47.562.500,00
    #
    # titik = pemisah ribuan
    # koma = decimal
    # --------------------------------------------------------------------------

    if "." in text and "," in text:

        if text.rfind(",") > text.rfind("."):

            text = text.replace(".", "")
            text = text.replace(",", ".")

        else:

            text = text.replace(",", "")

    # --------------------------------------------------------------------------
    # Hanya koma
    # --------------------------------------------------------------------------

    elif "," in text:

        parts = text.split(",")

        if (
            len(parts) == 2
            and len(parts[1]) <= 2
        ):

            text = text.replace(",", ".")

        else:

            text = text.replace(",", "")

    # --------------------------------------------------------------------------
    # Hanya titik
    #
    # 47.562.500
    #
    # Jika lebih dari satu titik → ribuan
    # --------------------------------------------------------------------------

    elif "." in text:

        if text.count(".") > 1:

            text = text.replace(".", "")

        else:

            parts = text.split(".")

            if (
                len(parts) == 2
                and len(parts[1]) == 3
            ):

                text = text.replace(".", "")

    try:

        return float(text)

    except Exception:

        return 0.0


# ==============================================================================
# 🔍 GET SITE DATA FROM QUERY
# ==============================================================================

def get_site_data_from_query(
    raw_query,
    site_name
):
    """
    Mengambil data site dari sheet Query.

    Struktur:

    C / index 2  = Charging Type
    F / index 5  = Site Name
    W / index 22 = WO Number
    """

    result = {
        "site_name": site_name,
        "charging_type": "",
        "wo_number": "",
    }

    if not site_name:
        return result

    if not raw_query or len(raw_query) <= 1:
        return result

    target = site_name.strip().lower()

    for row in raw_query[1:]:

        if len(row) <= 5:
            continue

        current_site = str(
            row[5]
        ).strip()

        if current_site.lower() != target:
            continue

        # Charging Type
        if len(row) > 2:

            result["charging_type"] = str(
                row[2]
            ).strip()

        # WO Number
        if len(row) > 22:

            result["wo_number"] = str(
                row[22]
            ).strip()

        break

    return result


# ==============================================================================
# ⭐ GET BOQ FROM DB BOQ
# ==============================================================================

def get_boq_amount_from_db_boq(
    raw_db_boq,
    site_name,
    charging_type
):
    """
    Mencari BOQ Amount dari sheet DB BOQ.

    Struktur DB BOQ:

    A = No
    B = BOQ No.
    C = Site Name
    D = Charger Type
    E = BOQ Amount Exc. PPN
    F = BOQ Amount inc. PPN
    G = EPC Name

    Matching:

        Site Name + Charger Type

    Return:

        BOQ Amount Exc. PPN
    """

    if not raw_db_boq:
        return 0.0

    if len(raw_db_boq) <= 1:
        return 0.0

    target_site = (
        str(site_name)
        .strip()
        .lower()
    )

    target_charger = (
        str(charging_type)
        .strip()
        .lower()
    )

    for row in raw_db_boq[1:]:

        # Minimal sampai kolom E
        if len(row) <= 4:
            continue

        db_site = (
            str(row[2])
            .strip()
            .lower()
        )

        db_charger = (
            str(row[3])
            .strip()
            .lower()
        )

        # Cocokkan Site Name + Charger Type
        if (
            db_site == target_site
            and db_charger == target_charger
        ):

            # ==============================================================
            # KOLOM E = INDEX 4
            # BOQ Amount Exc. PPN
            # ==============================================================

            return parse_amount(
                row[4]
            )

    return 0.0


# ==============================================================================
# 🔎 BUILD SITE DATA
# ==============================================================================

def build_site_data(
    raw_query,
    raw_db_boq,
    selected_sites
):
    """
    Membuat data lengkap setiap site.
    """

    site_data = []

    for site in selected_sites:

        query_data = get_site_data_from_query(
            raw_query,
            site
        )

        charging_type = query_data[
            "charging_type"
        ]

        wo_number = query_data[
            "wo_number"
        ]

        # ==============================================================
        # BOQ DARI DB BOQ
        # ==============================================================

        boq_amount = get_boq_amount_from_db_boq(
            raw_db_boq,
            site,
            charging_type
        )

        site_data.append(
            {
                "site_name": site,
                "charging_type": charging_type,
                "wo_number": wo_number,
                "boq_amount": boq_amount,
            }
        )

    return site_data


# ==============================================================================
# 🔎 SITE PREVIEW / CROSS CHECK
# ==============================================================================

def render_site_preview(
    site_data,
    selected_termin="Belum Dipilih",
    pct_val=0.0,
):
    """
    Menampilkan preview data site secara live.

    Preview ini muncul setelah user memilih Site Name,
    sehingga user dapat melakukan cross-check sebelum
    menyimpan invoice.
    """

    if not site_data:
        return

    st.markdown(
        "### 🔎 Preview & Cross Check Site"
    )

    st.caption(
        "Pastikan Site Name, Charger Type, WO Number, "
        "dan BOQ Amount sudah sesuai sebelum menyimpan Invoice."
    )

    preview_rows = []

    total_boq = 0.0
    total_invoice = 0.0

    for index, item in enumerate(
        site_data,
        start=1
    ):

        boq_amount = float(
            item.get(
                "boq_amount",
                0.0
            )
            or 0.0
        )

        invoice_amount = (
            boq_amount * pct_val
        )

        total_boq += boq_amount
        total_invoice += invoice_amount

        preview_rows.append(
            {
                "NO.": index,

                "SITE NAME": (
                    item.get(
                        "site_name",
                        ""
                    )
                ),

                "CHARGER TYPE": (
                    item.get(
                        "charging_type",
                        ""
                    )
                    or "-"
                ),

                "WO NUMBER": (
                    item.get(
                        "wo_number",
                        ""
                    )
                    or "-"
                ),

                "BOQ AMOUNT EXC. PPN": (
                    boq_amount
                ),

                f"INVOICE AMOUNT ({selected_termin})": (
                    invoice_amount
                ),
            }
        )

    preview_df = pd.DataFrame(
        preview_rows
    )

    # --------------------------------------------------------------------------
    # FORMAT NOMINAL
    # --------------------------------------------------------------------------

    display_df = preview_df.copy()

    display_df[
        "BOQ AMOUNT EXC. PPN"
    ] = display_df[
        "BOQ AMOUNT EXC. PPN"
    ].apply(
        lambda x: f"Rp {x:,.2f}"
    )

    invoice_col = (
        f"INVOICE AMOUNT ({selected_termin})"
    )

    display_df[
        invoice_col
    ] = display_df[
        invoice_col
    ].apply(
        lambda x: f"Rp {x:,.2f}"
    )

    # --------------------------------------------------------------------------
    # DISPLAY TABLE
    # --------------------------------------------------------------------------

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
    )

    # --------------------------------------------------------------------------
    # SUMMARY PREVIEW
    # --------------------------------------------------------------------------

    p1, p2, p3 = st.columns(3)

    with p1:

        st.metric(
            "Jumlah Site",
            len(site_data)
        )

    with p2:

        st.metric(
            "Total BOQ Exc. PPN",
            f"Rp {total_boq:,.2f}"
        )

    with p3:

        st.metric(
            f"Invoice {selected_termin}",
            f"Rp {total_invoice:,.2f}"
        )

    # --------------------------------------------------------------------------
    # WARNING BOQ KOSONG
    # --------------------------------------------------------------------------

    missing_boq = [
        item["site_name"]
        for item in site_data
        if float(
            item.get(
                "boq_amount",
                0.0
            )
            or 0.0
        ) <= 0
    ]

    if missing_boq:

        st.warning(
            "⚠️ BOQ Amount tidak ditemukan "
            "di DB BOQ untuk:"
        )

        for site in missing_boq:

            st.write(
                f"- {site}"
            )

        st.caption(
            "Pencarian BOQ menggunakan kombinasi "
            "Site Name + Charger Type pada sheet DB BOQ."
        )

    else:

        st.success(
            "✅ Seluruh Site yang dipilih memiliki "
            "BOQ Amount pada DB BOQ."
        )


# ==============================================================================
# 📄 CREATE INVOICE PDF
# ==============================================================================

def create_invoice_pdf(data):
    """
    Membuat PDF Invoice.
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

    pay_title_style = ParagraphStyle(
        name="PayTitle",
        parent=styles["Normal"],
        fontSize=9,
        leading=11,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#1A365D"),
    )

    table_text = ParagraphStyle(
        name="TableText",
        parent=styles["Normal"],
        fontSize=7.5,
        leading=9,
    )

    table_text_center = ParagraphStyle(
        name="TableTextCenter",
        parent=table_text,
        alignment=1,
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

    def draw_header_footer(
        canvas,
        doc
    ):

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
        colWidths=[
            80,
            180
        ]
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
        [[
            bill_to_text,
            meta_table
        ]],
        colWidths=[
            280,
            260
        ]
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

    elements.append(
        info_grid
    )

    elements.append(
        Spacer(1, 15)
    )

    # --------------------------------------------------------------------------
    # ITEM TABLE MULTI SITE
    # --------------------------------------------------------------------------

    table_data = [
        [
            Paragraph(
                "<b>NO.</b>",
                normal_bold
            ),
            Paragraph(
                "<b>ITEM DESCRIPTION</b>",
                normal_bold
            ),
            Paragraph(
                "<b>QTY</b>",
                normal_bold
            ),
            Paragraph(
                "<b>UOM</b>",
                normal_bold
            ),
            Paragraph(
                "<b>CHARGER TYPE</b>",
                normal_bold
            ),
            Paragraph(
                "<b>WO NUMBER</b>",
                normal_bold
            ),
            Paragraph(
                "<b>UNIT PRICE</b>",
                normal_bold
            ),
            Paragraph(
                "<b>AMOUNT</b>",
                normal_bold
            ),
        ]
    ]

    for index, item in enumerate(
        data["items"],
        start=1
    ):

        table_data.append(
            [
                Paragraph(
                    str(index),
                    table_text_center
                ),
                Paragraph(
                    str(
                        item["site_name"]
                    ),
                    table_text
                ),
                Paragraph(
                    "1",
                    table_text_center
                ),
                Paragraph(
                    "Unit",
                    table_text_center
                ),
                Paragraph(
                    str(
                        item["charging_type"]
                    ),
                    table_text_center
                ),
                Paragraph(
                    str(
                        item["wo_number"]
                    ),
                    table_text_center
                ),
                Paragraph(
                    f"{item['unit_price']:,.2f}",
                    table_text_center
                ),
                Paragraph(
                    f"{item['amount']:,.2f}",
                    table_text_center
                ),
            ]
        )

    item_table = Table(
        table_data,
        colWidths=[
            25,
            155,
            30,
            40,
            70,
            75,
            70,
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
                    (-1, -1),
                    "CENTER"
                ),
                (
                    "ALIGN",
                    (1, 1),
                    (1, -1),
                    "LEFT"
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
                ),
            ]
        )
    )

    elements.append(
        item_table
    )

    elements.append(
        Spacer(1, 10)
    )

    # --------------------------------------------------------------------------
    # SUMMARY
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
        colWidths=[
            380,
            160
        ]
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

    elements.append(
        summary_table
    )

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
        colWidths=[
            85,
            10,
            205
        ]
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

    sig_style = ParagraphStyle(
        name="SigStyle",
        parent=styles["Normal"],
        fontSize=8.5,
        leading=12,
        alignment=1,
    )

    sig_block = Paragraph(
        f"""
        Jakarta, {now_str}<br/><br/><br/><br/>
        <b><u>Christian</u></b><br/>
        President Director
        """,
        sig_style,
    )

    bottom_grid = Table(
        [[
            payment_info_table,
            sig_block
        ]],
        colWidths=[
            300,
            240
        ]
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

    elements.append(
        bottom_grid
    )

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
# 🖥️ RENDER
# ==============================================================================

def render():

    st.title(
        "📄 Create Invoice"
    )

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
        sheet_db_boq,
    ) = init_gspread()

    if not sheet_query:
        return

    # ==========================================================================
    # LOAD DATA
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

    # ⭐ DB BOQ
    raw_db_boq = get_raw_matrix(
        sheet_db_boq
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

                val = str(
                    row[6]
                ).strip()

                if (
                    val
                    and val not in project_options
                ):

                    project_options.append(
                        val
                    )

    if not project_options:

        project_options = [
            "VGreen - Project",
            "VGreen - Operation",
            "SIP",
            "Charge Core",
        ]

    # ==========================================================================
    # FORM
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
                options=project_options,
            )

            st.text_input(
                "No. Invoice (Auto Generated)",
                value=auto_inv_no,
                disabled=True,
            )

            inv_date = st.date_input(
                "Invoice Date",
                datetime.now()
            )

            # ==================================================================
            # SITE SELECTION
            # ==================================================================

            site_options = []

            if len(raw_query) > 1:

                for row in raw_query[1:]:

                    if len(row) > 5:

                        s_name = str(
                            row[5]
                        ).strip()

                        if (
                            s_name
                            and s_name not in site_options
                        ):

                            site_options.append(
                                s_name
                            )

            if selected_project == "VGreen - Project":

                if site_options:

                    selected_sites = st.multiselect(
                        "Item Description / Site Name",
                        options=site_options,
                        help=(
                            "Pilih satu atau lebih site. "
                            "Data Charger Type, WO Number "
                            "dan BOQ Amount akan otomatis "
                            "ditampilkan."
                        ),
                    )

                else:

                    selected_sites = []

                    st.warning(
                        "⚠️ Tidak ada Site Name "
                        "yang ditemukan pada sheet Query."
                    )

            else:

                manual_site = st.text_input(
                    "Item Description / Site Name",
                    placeholder=(
                        "Ketik Site Name secara manual..."
                    ),
                )

                selected_sites = (
                    [manual_site]
                    if manual_site.strip()
                    else []
                )

        # ======================================================================
        # RIGHT COLUMN
        # ======================================================================

        with col2:

            st.text_input(
                "Mode",
                value=(
                    "Multi Site"
                    if selected_project
                    == "VGreen - Project"
                    else "Manual"
                ),
                disabled=True,
            )

            efaktur_no = st.text_input(
                "No. Efaktur",
                placeholder=(
                    "Contoh: 04002600290355647"
                ),
            )

        # ======================================================================
        # SITE DATA
        # ======================================================================

        site_data = build_site_data(
            raw_query,
            raw_db_boq,
            selected_sites,
        )

        # ======================================================================
        # ⭐ LIVE SITE PREVIEW
        # ======================================================================
        #
        # Preview langsung muncul setelah Site Name dipilih.
        # Nilai invoice akan mengikuti Termin yang dipilih
        # di bawahnya.
        #
        # Untuk tahap awal sebelum Termin dipilih, pct_val = 0.
        # Preview akan diperbarui lagi setelah Termin dipilih.
        # ======================================================================

        if selected_project == "VGreen - Project":

            if site_data:

                st.markdown(
                    "### 🔎 Preview Site yang Dipilih"
                )

                st.caption(
                    "Cross-check data Site Name, Charger Type, "
                    "WO Number, dan BOQ Amount sebelum menentukan "
                    "Termin pembayaran."
                )

                early_preview_rows = []

                early_total_boq = 0.0

                for index, item in enumerate(
                    site_data,
                    start=1
                ):

                    boq_value = float(
                        item.get(
                            "boq_amount",
                            0.0
                        )
                        or 0.0
                    )

                    early_total_boq += (
                        boq_value
                    )

                    early_preview_rows.append(
                        {
                            "NO.": index,

                            "SITE NAME": (
                                item["site_name"]
                            ),

                            "CHARGER TYPE": (
                                item["charging_type"]
                                or "-"
                            ),

                            "WO NUMBER": (
                                item["wo_number"]
                                or "-"
                            ),

                            "BOQ AMOUNT EXC. PPN": (
                                boq_value
                            ),
                        }
                    )

                early_preview_df = pd.DataFrame(
                    early_preview_rows
                )

                early_display_df = (
                    early_preview_df.copy()
                )

                early_display_df[
                    "BOQ AMOUNT EXC. PPN"
                ] = early_display_df[
                    "BOQ AMOUNT EXC. PPN"
                ].apply(
                    lambda x: f"Rp {x:,.2f}"
                )

                st.dataframe(
                    early_display_df,
                    use_container_width=True,
                    hide_index=True,
                )

                ep1, ep2 = st.columns(2)

                with ep1:

                    st.metric(
                        "Jumlah Site Dipilih",
                        len(site_data)
                    )

                with ep2:

                    st.metric(
                        "Total BOQ Exc. PPN",
                        f"Rp {early_total_boq:,.2f}"
                    )

                early_missing_boq = [
                    item["site_name"]
                    for item in site_data
                    if item["boq_amount"] <= 0
                ]

                if early_missing_boq:

                    st.warning(
                        "⚠️ Ada Site yang belum memiliki "
                        "BOQ Amount:"
                    )

                    for site in early_missing_boq:

                        st.write(
                            f"- {site}"
                        )

                else:

                    st.success(
                        "✅ Data BOQ seluruh site ditemukan."
                    )

        # ======================================================================
        # TOP
        # ======================================================================

        st.markdown(
            "### 💳 Detail Pembayaran & Termin"
        )

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
        # USED TERMIN
        # ======================================================================

        used_termins = []

        if (
            selected_project
            == "VGreen - Project"
            and len(raw_db) > 1
            and selected_sites
        ):

            selected_site_lower = [
                s.strip().lower()
                for s in selected_sites
            ]

            for r in raw_db[1:]:

                if len(r) > 6:

                    db_charging = (
                        str(r[3])
                        .strip()
                        .lower()
                    )

                    db_site = (
                        str(r[4])
                        .strip()
                        .lower()
                    )

                    db_term = str(
                        r[6]
                    ).strip()

                    if (
                        db_site
                        in selected_site_lower
                        and any(
                            db_site
                            == s["site_name"]
                            .strip()
                            .lower()
                            and db_charging
                            == s["charging_type"]
                            .strip()
                            .lower()
                            for s in site_data
                        )
                    ):

                        if db_term:

                            used_termins.append(
                                db_term
                            )

        available_termins = [
            t
            for t in all_possible_termins
            if t not in used_termins
        ]

        if (
            selected_project
            == "VGreen - Project"
            and used_termins
        ):

            st.warning(
                "ℹ️ Beberapa termin sudah pernah "
                "dibuat untuk site yang dipilih: "
                + ", ".join(
                    sorted(
                        set(used_termins)
                    )
                )
            )

        if (
            selected_project
            == "VGreen - Project"
            and selected_sites
            and not available_termins
        ):

            st.error(
                "⛔ Semua termin pada skema "
                f"{selected_schema} sudah terpakai."
            )

        # ======================================================================
        # TERMIN / TAX
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
        # ⭐ HITUNG BOQ PER SITE DARI DB BOQ
        # ======================================================================

        total_boq_amount = 0.0
        total_invoice_amount = 0.0

        for item in site_data:

            item["unit_price"] = (
                item["boq_amount"]
                * pct_val
            )

            item["amount"] = (
                item["unit_price"]
            )

            total_boq_amount += (
                item["boq_amount"]
            )

            total_invoice_amount += (
                item["amount"]
            )

        # ======================================================================
        # ⭐ SECOND / UPDATED PREVIEW
        # ======================================================================
        #
        # Setelah Termin dipilih, preview akan menampilkan
        # nominal invoice per site berdasarkan persentase termin.
        # ======================================================================

        if (
            selected_project
            == "VGreen - Project"
            and site_data
        ):

            render_site_preview(
                site_data=site_data,
                selected_termin=selected_termin,
                pct_val=pct_val,
            )

        # ======================================================================
        # UNIT PRICE
        # ======================================================================

        with c2:

            if selected_project == "VGreen - Project":

                unit_price = st.number_input(
                    "Total Unit Price / Invoice Amount",
                    value=float(
                        total_invoice_amount
                    ),
                    disabled=True,
                    format="%.2f",
                )

            else:

                unit_price = st.number_input(
                    "Unit Price (IDR)",
                    min_value=0.0,
                    value=0.0,
                    step=1000.0,
                )

                total_invoice_amount = (
                    unit_price
                )

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
        # TAX CALCULATION
        # ======================================================================

        termin_amount = (
            total_invoice_amount
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
        # SUMMARY
        # ======================================================================

        st.info(
            f"""
**Ringkasan Perhitungan:**

* **Skema TOP Terpilih:** {selected_schema}
* **Jumlah Site:** {len(site_data)}
* **Total BOQ Amount Exc. PPN:** Rp {total_boq_amount:,.2f}
* **Nilai Invoice ({selected_termin}):** Rp {termin_amount:,.2f}
* **PPN ({tax_rate}%):** Rp {tax_amount:,.2f}
* **GRAND TOTAL:** Rp {grand_total:,.2f}
"""
        )

        # ======================================================================
        # SUBMIT
        # ======================================================================

        is_disabled = (
            selected_project
            == "VGreen - Project"
            and (
                not available_termins
                or not selected_sites
            )
        )

        submit_btn = st.form_submit_button(
            "💾 Save to DB Invoice",
            type="primary",
            disabled=is_disabled,
        )

    # ==========================================================================
    # SAVE
    # ==========================================================================

    if submit_btn:

        # ----------------------------------------------------------------------
        # VALIDATION
        # ----------------------------------------------------------------------

        if not efaktur_no:

            st.error(
                "⚠️ No. Efaktur wajib diisi!"
            )

            return

        if not selected_sites:

            st.error(
                "⚠️ Site Name / Item Description "
                "wajib diisi!"
            )

            return

        # ----------------------------------------------------------------------
        # VALIDASI BOQ
        # ----------------------------------------------------------------------

        if (
            selected_project
            == "VGreen - Project"
        ):

            missing_boq = [
                item["site_name"]
                for item in site_data
                if item["boq_amount"] <= 0
            ]

            if missing_boq:

                st.error(
                    "🚨 Invoice tidak dapat disimpan "
                    "karena BOQ Amount tidak ditemukan "
                    "untuk site berikut:"
                )

                for site in missing_boq:

                    st.write(
                        f"- {site}"
                    )

                st.info(
                    "Pastikan Site Name dan Charger Type "
                    "pada DB BOQ sesuai dengan data Query."
                )

                return

        # ----------------------------------------------------------------------
        # SAVE EACH SITE
        # ----------------------------------------------------------------------

        if sheet_db:

            try:

                # ==============================================================
                # SATU INVOICE NUMBER
                # DIGUNAKAN UNTUK SEMUA SITE
                # ==============================================================

                for item in site_data:

                    tax_inclusive_amount = (
                        item["amount"]
                        + (
                            item["amount"]
                            * tax_rate
                            / 100
                        )
                    )

                    new_row = [
                        selected_project,

                        auto_inv_no,

                        inv_date.strftime(
                            "%d %B %Y"
                        ),

                        item["charging_type"],

                        item["site_name"],

                        item["wo_number"],

                        selected_termin,

                        efaktur_no,

                        f"{item['amount']:.2f}",

                        f"{tax_inclusive_amount:.2f}",
                    ]

                    sheet_db.append_row(
                        new_row
                    )

                st.success(
                    f"✅ Invoice **{auto_inv_no}** "
                    f"berhasil disimpan ke Sheet "
                    f"'DB Invoice' untuk "
                    f"**{len(site_data)} site**!"
                )

                st.balloons()

                # ==============================================================
                # PDF PAYLOAD
                # ==============================================================

                pdf_payload = {
                    "project_name": selected_project,

                    "inv_no": auto_inv_no,

                    "inv_date": inv_date.strftime(
                        "%d %B %Y"
                    ),

                    "efaktur": efaktur_no,

                    "items": site_data,

                    "subtotal": termin_amount,

                    "tax_rate": tax_rate,

                    "tax_amount": tax_amount,

                    "grand_total": grand_total,
                }

                # ==============================================================
                # CREATE PDF
                # ==============================================================

                pdf_file_bytes = create_invoice_pdf(
                    pdf_payload
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

            except Exception as e:

                st.error(
                    "🚨 Gagal menyimpan ke "
                    f"Google Sheets: {e}"
                )

        else:

            st.error(
                "🚨 Tidak terhubung ke "
                "Google Sheets."
            )

    # ==========================================================================
    # DOWNLOAD PDF
    # ==========================================================================

    if (
        "pdf_ready" in st.session_state
        and st.session_state["pdf_ready"]
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
