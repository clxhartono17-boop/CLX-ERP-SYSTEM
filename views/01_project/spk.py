import datetime
import io
import os
import re

import pandas as pd
import streamlit as st

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    Image,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# Import Koneksi Database sesuai Struktur ERP
from core.database import get_google_sheet_connection


# ==============================================================================
# CONFIGURATION
# ==============================================================================

COO_PIN_SECRET = "1234"

SIGNATURE_PATH = os.path.join(
    "assets",
    "Approved CFO.png"
)

COMPANY_NAME = "PT. Connectivity Leads eXcellence"

DEFAULT_PHONE = "0851-8259-6296"


# ==============================================================================
# CSS
# ==============================================================================

PASTEL_ORANGE_CSS = """
<style>

    .stApp,
    [data-testid="stAppViewContainer"] {
        background-color: #FAFAFA !important;
        color: #2D3748 !important;
    }

    div[data-baseweb="select"] > div,
    div[data-baseweb="select"] > div * {
        background-color: #FFE5D9 !important;
        color: #1A202C !important;
        border-color: #FFCAD4 !important;
        font-weight: 600 !important;
    }

    div[data-baseweb="select"] svg {
        fill: #C05621 !important;
    }

    ul[data-baseweb="menu"],
    div[data-baseweb="popover"] {
        background-color: #FFF0EB !important;
    }

    li[data-baseweb="option"] {
        color: #2D3748 !important;
        background-color: #FFF0EB !important;
    }

    li[data-baseweb="option"]:hover,
    li[aria-selected="true"] {
        background-color: #FFD7C2 !important;
        color: #9C4221 !important;
    }

    div[data-baseweb="input"] > div,
    div[data-baseweb="input"] input {
        background-color: #FFFFFF !important;
        color: #1A202C !important;
        border-color: #FFCAD4 !important;
    }

    code {
        background-color: #FFE5D9 !important;
        color: #C05621 !important;
        border: 1px solid #FFCAD4 !important;
        font-weight: bold !important;
        padding: 3px 8px !important;
        border-radius: 6px !important;
    }

    div[data-testid="stDataEditor"],
    div[data-testid="stDataEditor"] > div,
    .dgb-grid,
    canvas {
        background-color: #FFFFFF !important;
    }

</style>
"""


# ==============================================================================
# HELPER - HEADER / DATAFRAME
# ==============================================================================

def normalize_headers(headers):
    """
    Membersihkan header Google Sheet:
    - menghilangkan spasi
    - menangani header kosong
    - menghindari duplicate columns
    """

    result = []
    used = {}

    for i, header in enumerate(headers):

        name = str(header).strip()

        if not name:
            name = f"Column_{i + 1}"

        if name in used:
            used[name] += 1
            name = f"{name}_{used[name]}"
        else:
            used[name] = 1

        result.append(name)

    return result


def dataframe_from_sheet_rows(rows):
    """
    Convert Google Sheet rows menjadi DataFrame yang aman.
    """

    if not rows or len(rows) <= 1:
        return pd.DataFrame()

    headers = normalize_headers(rows[0])

    data = rows[1:]

    normalized_data = []

    for row in data:

        row = list(row)

        if len(row) < len(headers):
            row += [""] * (len(headers) - len(row))

        elif len(row) > len(headers):
            row = row[:len(headers)]

        normalized_data.append(row)

    return pd.DataFrame(
        normalized_data,
        columns=headers
    )


def find_column(df, candidates, default=None):
    """
    Mencari nama kolom secara fleksibel.
    """

    if df is None or df.empty:
        return default

    normalized = {}

    for col in df.columns:

        key = re.sub(
            r"[^a-z0-9]",
            "",
            str(col).lower()
        )

        normalized[key] = col

    for candidate in candidates:

        key = re.sub(
            r"[^a-z0-9]",
            "",
            str(candidate).lower()
        )

        if key in normalized:
            return normalized[key]

    return default


def ensure_column(sheet, column_name):
    """
    Memastikan kolom tersedia pada Google Sheet.
    """

    rows = sheet.get_all_values()

    if not rows:
        sheet.append_row([column_name])
        return

    headers = normalize_headers(rows[0])

    if column_name not in headers:

        next_col = len(headers) + 1

        # Update header
        sheet.update_cell(
            1,
            next_col,
            column_name
        )


def ensure_database_columns(sheet):
    """
    Memastikan struktur database SPK mempunyai
    kolom yang diperlukan.
    """

    required_columns = [
        "Date SPK",
        "No. SPK",
        "No. WO",
        "Pekerjaan",
        "EPC",
        "Site Name",
        "Charger Type",
        "WO Release",
        "WO End",
        "Mitra",
        "Date SPK (Take Over)",
        "No. SPK (Take Over)",
        "Mitra (Take Over)",
        "Status Approval",
        "Sign COO",
    ]

    rows = sheet.get_all_values()

    if not rows:

        sheet.append_row(required_columns)
        return

    headers = normalize_headers(rows[0])

    for col in required_columns:

        if col not in headers:

            next_col = len(headers) + 1

            sheet.update_cell(
                1,
                next_col,
                col
            )

            headers.append(col)


# ==============================================================================
# GENERATE NOMOR SPK
# ==============================================================================

def generate_spk_number(
    sow_type="GENERAL",
    sequence_num=1
):

    now = datetime.datetime.now()

    roman_months = [
        "I",
        "II",
        "III",
        "IV",
        "V",
        "VI",
        "VII",
        "VIII",
        "IX",
        "X",
        "XI",
        "XII",
    ]

    month_roman = roman_months[now.month - 1]

    sow_code = (
        "SURVEY"
        if "survey" in str(sow_type).lower()
        else "CONS"
    )

    seq_str = f"{sequence_num:04d}"

    return (
        f"{seq_str}/CLX/SPK/"
        f"{sow_code}/{month_roman}/{now.year}"
    )


# ==============================================================================
# SIGNATURE
# ==============================================================================

def get_signature_path():

    if os.path.exists(SIGNATURE_PATH):
        return SIGNATURE_PATH

    return None


# ==============================================================================
# GENERATE PDF
# ==============================================================================

def generate_spk_pdf_bytes(
    selected_wo,
    selected_sites,
    spk_metadata,
    matched_sow_df,
    approved=False,
):

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=portrait(letter),
        leftMargin=0.4 * inch,
        rightMargin=0.4 * inch,
        topMargin=0.4 * inch,
        bottomMargin=0.4 * inch,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Heading1"],
        fontSize=12,
        leading=14,
        alignment=1,
        fontName="Helvetica-Bold",
        textColor=colors.black,
        spaceAfter=2,
    )

    subtitle_style = ParagraphStyle(
        "DocSubtitle",
        parent=styles["Normal"],
        fontSize=9,
        leading=11,
        alignment=1,
        fontName="Helvetica-Bold",
        textColor=colors.black,
        spaceAfter=10,
    )

    header_left = ParagraphStyle(
        "HeaderLeft",
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#1B365D"),
    )

    header_right = ParagraphStyle(
        "HeaderRight",
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#333333"),
    )

    meta_label = ParagraphStyle(
        "MetaLabel",
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
    )

    meta_val = ParagraphStyle(
        "MetaVal",
        fontName="Helvetica",
        fontSize=8,
        leading=10,
    )

    cell_head = ParagraphStyle(
        "CellHead",
        fontName="Helvetica-Bold",
        fontSize=6.5,
        leading=8,
        alignment=1,
        textColor=colors.white,
    )

    cell_body = ParagraphStyle(
        "CellBody",
        fontName="Helvetica",
        fontSize=6,
        leading=7.5,
    )

    sign_style_bold = ParagraphStyle(
        "SignBold",
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
        alignment=1,
    )

    sign_style_norm = ParagraphStyle(
        "SignNorm",
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        alignment=1,
    )

    elements = []

    # ==========================================================================
    # HEADER
    # ==========================================================================

    logo_path = "assets/logo.png"

    if os.path.exists(logo_path):

        company_logo = Image(
            logo_path,
            width=1.5 * inch,
            height=1.0 * inch,
        )

    else:

        company_logo = Paragraph(
            COMPANY_NAME,
            header_left
        )

    header_data = [
        [
            company_logo,
            Paragraph(
                f"<b>{COMPANY_NAME}</b><br/>"
                "<b>Jakarta Office:</b> "
                "Jl. M. Ali 2 No. 19 RT 007 RW 004 "
                "Tanah Baru Beji Kota Depok Jawa Barat<br/>"
                "<b>E:</b> clx.central@gmail.com",
                header_right,
            ),
        ]
    ]

    t_header = Table(
        header_data,
        colWidths=[
            2.2 * inch,
            5.1 * inch
        ],
    )

    t_header.setStyle(
        TableStyle([
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE",
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                4,
            ),
            (
                "LINEBELOW",
                (0, 0),
                (-1, -1),
                1,
                colors.HexColor("#1B365D"),
            ),
        ])
    )

    elements.append(t_header)
    elements.append(Spacer(1, 8))

    # ==========================================================================
    # TITLE
    # ==========================================================================

    no_spk = spk_metadata.get(
        "no_spk",
        "0001/CLX/SPK/CONS/VII/2026"
    )

    date_value = spk_metadata.get(
        "date_spk",
        datetime.datetime.now()
    )

    if isinstance(date_value, datetime.datetime):

        date_str = date_value.strftime(
            "%d %B %Y"
        )

    else:

        date_str = str(date_value)

    elements.append(
        Paragraph(
            "SURAT PERINTAH KERJA (SPK)",
            title_style,
        )
    )

    elements.append(
        Paragraph(
            f"No. {no_spk}<br/>"
            f"Jakarta, {date_str}",
            subtitle_style,
        )
    )

    # ==========================================================================
    # APPROVED STAMP
    # ==========================================================================

    if approved:

        approved_style = ParagraphStyle(
            "ApprovedStyle",
            fontName="Helvetica-Bold",
            fontSize=9,
            alignment=1,
            textColor=colors.HexColor("#16803C"),
        )

        elements.append(
            Paragraph(
                "✓ APPROVED BY COO",
                approved_style,
            )
        )

        elements.append(
            Spacer(1, 6)
        )

    # ==========================================================================
    # METADATA
    # ==========================================================================

    meta_table_data = [

        [
            Paragraph("Proyek:", meta_label),
            Paragraph(
                str(
                    spk_metadata.get(
                        "proyek",
                        "-"
                    )
                ),
                meta_val,
            ),

            Paragraph("No. WO:", meta_label),
            Paragraph(
                str(selected_wo),
                meta_val,
            ),
        ],

        [
            Paragraph("Pekerjaan:", meta_label),
            Paragraph(
                str(
                    spk_metadata.get(
                        "pekerjaan",
                        "-"
                    )
                ),
                meta_val,
            ),

            Paragraph("Penanggung Jawab:", meta_label),
            Paragraph(
                f"{spk_metadata.get('pic_name', '-')}"
                f" ({spk_metadata.get('pic_phone', '-')})",
                meta_val,
            ),
        ],

        [
            Paragraph("Lokasi:", meta_label),
            Paragraph(
                str(
                    spk_metadata.get(
                        "lokasi",
                        "-"
                    )
                ),
                meta_val,
            ),

            Paragraph(
                "Penerbit (PT. CLX):",
                meta_label,
            ),

            Paragraph(
                "Wikantiyoso S. (0878-8855-0300)<br/>"
                "Hartono (0818-0690-9317)",
                meta_val,
            ),
        ],

        [
            Paragraph("Detail Site:", meta_label),

            Paragraph(
                "Terlampir pada tabel di bawah",
                meta_val,
            ),

            Paragraph("SOW:", meta_label),

            Paragraph(
                str(
                    spk_metadata.get(
                        "sow_type",
                        "-"
                    )
                ),
                meta_val,
            ),
        ],
    ]

    t_meta = Table(
        meta_table_data,
        colWidths=[
            1.0 * inch,
            2.7 * inch,
            1.2 * inch,
            2.4 * inch,
        ],
    )

    t_meta.setStyle(
        TableStyle([
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP",
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
            (
                "BACKGROUND",
                (0, 0),
                (-1, -1),
                colors.HexColor("#F8F9FA"),
            ),
            (
                "BOX",
                (0, 0),
                (-1, -1),
                0.5,
                colors.HexColor("#CCCCCC"),
            ),
            (
                "INNERGRID",
                (0, 0),
                (-1, -1),
                0.25,
                colors.HexColor("#E5E7EB"),
            ),
        ])
    )

    elements.append(t_meta)
    elements.append(Spacer(1, 10))

    # ==========================================================================
    # SOW
    # ==========================================================================

    sow_headers = [

        Paragraph(
            "<b>No.</b>",
            cell_head
        ),

        Paragraph(
            "<b>Uraian Pekerjaan (SoW)</b>",
            cell_head
        ),

        Paragraph(
            "<b>Target Penyelesaian</b>",
            cell_head
        ),
    ]

    sow_rows = [sow_headers]

    if (
        matched_sow_df is not None
        and not matched_sow_df.empty
    ):

        for idx, (_, row) in enumerate(
            matched_sow_df.iterrows(),
            1
        ):

            deskripsi = (
                str(row.iloc[2])
                if row.shape[0] > 2
                else "-"
            )

            target = (
                str(row.iloc[3])
                if row.shape[0] > 3
                else "1 Hari"
            )

            sow_rows.append([
                Paragraph(
                    str(idx),
                    cell_body
                ),

                Paragraph(
                    deskripsi,
                    cell_body
                ),

                Paragraph(
                    target,
                    cell_body
                ),
            ])

    else:

        sow_rows.append([

            Paragraph(
                "1",
                cell_body
            ),

            Paragraph(
                "SOW Pekerjaan Standar",
                cell_body
            ),

            Paragraph(
                "1 Hari",
                cell_body
            ),
        ])

    t_sow = Table(
        sow_rows,
        colWidths=[
            0.4 * inch,
            5.0 * inch,
            1.9 * inch,
        ],
    )

    t_sow.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor("#1B365D"),
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE",
            ),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.HexColor("#BDC3C7"),
            ),
            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                3,
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                3,
            ),
        ])
    )

    elements.append(t_sow)
    elements.append(Spacer(1, 10))

    # ==========================================================================
    # SITE LIST
    # ==========================================================================

    elements.append(
        Paragraph(
            "<b>Detail Site List:</b>",
            ParagraphStyle(
                "SubHeader",
                fontName="Helvetica-Bold",
                fontSize=9,
                spaceAfter=4,
            ),
        )
    )

    site_headers = [

        Paragraph(
            "<b>No</b>",
            cell_head
        ),

        Paragraph(
            "<b>Site Name</b>",
            cell_head
        ),

        Paragraph(
            "<b>Charging Type</b>",
            cell_head
        ),

        Paragraph(
            "<b>WO Number</b>",
            cell_head
        ),

        Paragraph(
            "<b>Province</b>",
            cell_head
        ),

        Paragraph(
            "<b>PIC + Contact</b>",
            cell_head
        ),

        Paragraph(
            "<b>Gmaps</b>",
            cell_head
        ),
    ]

    site_rows = [site_headers]

    if selected_sites is not None:

        for idx, (_, row) in enumerate(
            selected_sites.iterrows(),
            1
        ):

            site_rows.append([

                Paragraph(
                    str(idx),
                    cell_body
                ),

                Paragraph(
                    str(
                        row.get(
                            "col_site",
                            row.get(
                                "Site Name",
                                "-"
                            )
                        )
                    ),
                    cell_body
                ),

                Paragraph(
                    str(
                        row.get(
                            "col_charge",
                            row.get(
                                "Charger Type",
                                "-"
                            )
                        )
                    ),
                    cell_body
                ),

                Paragraph(
                    str(selected_wo),
                    cell_body
                ),

                Paragraph(
                    str(
                        row.get(
                            "col_province",
                            row.get(
                                "Province",
                                "-"
                            )
                        )
                    ),
                    cell_body
                ),

                Paragraph(
                    str(
                        row.get(
                            "col_pic",
                            row.get(
                                "PIC",
                                "-"
                            )
                        )
                    ),
                    cell_body
                ),

                Paragraph(
                    str(
                        row.get(
                            "col_gmaps",
                            row.get(
                                "Gmaps",
                                "-"
                            )
                        )
                    ),
                    cell_body
                ),
            ])

    col_widths = [
        0.3 * inch,
        1.6 * inch,
        1.0 * inch,
        1.3 * inch,
        0.9 * inch,
        1.1 * inch,
        1.1 * inch,
    ]

    t_site = Table(
        site_rows,
        colWidths=col_widths,
    )

    t_site.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor("#1B365D"),
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP",
            ),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.4,
                colors.HexColor("#BDC3C7"),
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
        ])
    )

    elements.append(t_site)
    elements.append(Spacer(1, 15))

    # ==========================================================================
    # SIGNATURE
    # ==========================================================================

    if approved and get_signature_path():

        try:

            digital_signature = Image(
                get_signature_path(),
                width=1.65 * inch,
                height=0.75 * inch,
            )

        except Exception:

            digital_signature = Spacer(
                1,
                0.75 * inch
            )

    else:

        digital_signature = Spacer(
            1,
            0.75 * inch
        )

    sign_data = [

        [
            Paragraph(
                "<b>PELAKSANA PEKERJAAN</b>",
                sign_style_bold
            ),

            Paragraph(
                "<b>PEMBERI PERINTAH KERJA</b><br/>"
                f"<b>{COMPANY_NAME}</b>",
                sign_style_bold
            ),
        ],

        [
            Spacer(1, 25),

            digital_signature
            if approved
            else Spacer(1, 25),
        ],

        [

            Paragraph(
                f"<b>"
                f"{spk_metadata.get('pic_name', 'Edy')}"
                f"</b><br/>"
                f"Contact: "
                f"{spk_metadata.get('pic_phone', '-')}",
                sign_style_norm
            ),

            Paragraph(
                "<b>Wikantiyoso Suyono</b><br/>"
                "Chief Operating Officer",
                sign_style_norm
            ),
        ],
    ]

    t_sign = Table(
        sign_data,
        colWidths=[
            3.6 * inch,
            3.7 * inch
        ],
    )

    t_sign.setStyle(
        TableStyle([
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP",
            ),
            (
                "ALIGN",
                (0, 0),
                (-1, -1),
                "CENTER",
            ),
        ])
    )

    elements.append(
        KeepTogether([
            t_sign
        ])
    )

    # ==========================================================================
    # APPROVAL FOOTER
    # ==========================================================================

    if approved:

        elements.append(
            Spacer(1, 10)
        )

        approved_footer = ParagraphStyle(
            "ApprovedFooter",
            fontName="Helvetica-Bold",
            fontSize=7,
            alignment=1,
            textColor=colors.HexColor("#16803C"),
        )

        elements.append(
            Paragraph(
                "Digitally Approved by Chief Operating Officer",
                approved_footer,
            )
        )

    doc.build(elements)

    buffer.seek(0)

    return no_spk, buffer.getvalue()


# ==============================================================================
# BUILD APPROVED PDF FROM DATABASE ROWS
# ==============================================================================

def build_approved_pdf(
    sh,
    db_sheet_name,
    selected_spk,
):

    sheet = sh.worksheet(
        db_sheet_name
    )

    rows = sheet.get_all_values()

    df = dataframe_from_sheet_rows(
        rows
    )

    if df.empty:
        raise ValueError(
            "Database SPK masih kosong."
        )

    spk_col = find_column(
        df,
        [
            "No. SPK",
            "No SPK",
            "SPK",
        ]
    )

    if not spk_col:
        raise ValueError(
            "Kolom No. SPK tidak ditemukan."
        )

    selected_rows = df[
        df[spk_col]
        .astype(str)
        .str.strip()
        == str(selected_spk).strip()
    ].copy()

    if selected_rows.empty:
        raise ValueError(
            f"SPK {selected_spk} tidak ditemukan."
        )

    # ==========================================================================
    # Cek approval
    # ==========================================================================

    status_col = find_column(
        selected_rows,
        [
            "Status Approval",
            "Approval Status",
        ]
    )

    if status_col:

        statuses = (
            selected_rows[status_col]
            .astype(str)
            .str.strip()
            .str.lower()
            .unique()
            .tolist()
        )

        if not any(
            "approved" in status
            for status in statuses
        ):

            raise ValueError(
                "SPK ini belum Approved oleh COO."
            )

    # ==========================================================================
    # Ambil metadata
    # ==========================================================================

    first = selected_rows.iloc[0]

    def val(candidates, default="-"):

        col = find_column(
            selected_rows,
            candidates
        )

        if col:
            value = first[col]

            if pd.isna(value):
                return default

            return str(value)

        return default

    selected_wo = val(
        [
            "No. WO",
            "WO Number",
            "WO"
        ]
    )

    pekerjaan = val(
        [
            "Pekerjaan",
            "Work"
        ]
    )

    mitra = val(
        [
            "Mitra",
            "PIC",
            "Penanggung Jawab"
        ]
    )

    date_spk = val(
        [
            "Date SPK"
        ],
        "-"
    )

    # ==========================================================================
    # Site dataframe
    # ==========================================================================

    site_col = find_column(
        selected_rows,
        [
            "Site Name",
            "Site"
        ]
    )

    charger_col = find_column(
        selected_rows,
        [
            "Charger Type",
            "Charging Type"
        ]
    )

    site_rows = []

    for _, row in selected_rows.iterrows():

        site_rows.append({
            "col_site": (
                row[site_col]
                if site_col
                else "-"
            ),

            "col_charge": (
                row[charger_col]
                if charger_col
                else "-"
            ),

            "col_province": "-",

            "col_pic": mitra,

            "col_gmaps": "-",
        })

    selected_sites = pd.DataFrame(
        site_rows
    )

    # ==========================================================================
    # Ambil Master SOW
    # ==========================================================================

    matched_sow_df = pd.DataFrame()

    try:

        sow_sheet = sh.worksheet(
            "Master SOW"
        )

        sow_rows = sow_sheet.get_all_values()

        df_sow = dataframe_from_sheet_rows(
            sow_rows
        )

        if not df_sow.empty:

            sow_col = df_sow.columns[0]

            sow_value = val(
                [
                    "WO Release",
                    "SOW",
                    "Jenis SOW"
                ]
            )

            matched_sow_df = df_sow[
                df_sow[sow_col]
                .astype(str)
                .str.strip()
                .str.lower()
                ==
                str(sow_value)
                .strip()
                .lower()
            ]

    except Exception:
        matched_sow_df = pd.DataFrame()

    # ==========================================================================
    # Metadata
    # ==========================================================================

    spk_metadata = {

        "no_spk": selected_spk,

        "proyek": "V-Green",

        "pekerjaan": pekerjaan,

        "lokasi": "-",

        "pic_name": mitra,

        "pic_phone": "-",

        "sow_type": val(
            [
                "WO Release",
                "SOW"
            ]
        ),

        "date_spk": date_spk,
    }

    return generate_spk_pdf_bytes(
        selected_wo,
        selected_sites,
        spk_metadata,
        matched_sow_df,
        approved=True,
    )


# ==============================================================================
# MAIN PAGE
# ==============================================================================

def show_spk_page():

    st.markdown(
        PASTEL_ORANGE_CSS,
        unsafe_allow_html=True
    )

    sh = get_google_sheet_connection()

    if not sh:

        st.error(
            "❌ Koneksi ke Database tidak tersedia."
        )

        return

    # ==========================================================================
    # MAIN MENU
    # ==========================================================================

    main_menu = st.radio(

        "📌 **PILIH MENU OPERASIONAL:**",

        options=[
            "📄 Generate SPK Baru",
            "🔄 Take Over Site",
            "🔒 COO Approval Dashboard",
            "📥 Download SPK",
        ],

        horizontal=True,

        key="main_spk_menu",
    )

    st.markdown("---")

    # ==========================================================================
    # MENU 1 - GENERATE SPK
    # ==========================================================================

    if main_menu == "📄 Generate SPK Baru":

        st.subheader(
            "📄 Generator Surat Perintah Kerja (SPK) & PDF"
        )

        st.markdown(
            "Pilih Jenis SOW & Nomor WO dari **Sheet Query**, "
            "centang site yang diinginkan, lalu klik "
            "**Generate SPK**."
        )

        try:

            query_sheet = sh.worksheet(
                "Query"
            )

            query_rows = query_sheet.get_all_values()

            # ------------------------------------------------------------------
            # MASTER DROPDOWN
            # ------------------------------------------------------------------

            try:

                dropdown_sheet = sh.worksheet(
                    "Master Dropdown"
                )

                dropdown_data = (
                    dropdown_sheet
                    .get_all_values()
                )

                df_dropdown = (
                    dataframe_from_sheet_rows(
                        dropdown_data
                    )
                )

            except Exception:

                df_dropdown = pd.DataFrame()

            # ------------------------------------------------------------------
            # MASTER SOW
            # ------------------------------------------------------------------

            try:

                sow_sheet = sh.worksheet(
                    "Master SOW"
                )

                sow_data = (
                    sow_sheet
                    .get_all_values()
                )

                df_master_sow = (
                    dataframe_from_sheet_rows(
                        sow_data
                    )
                )

            except Exception:

                df_master_sow = pd.DataFrame()

            if not query_rows or len(query_rows) <= 1:

                st.warning(
                    "⚠️ Belum ada data pada sheet 'Query'."
                )

            else:

                df_query = (
                    dataframe_from_sheet_rows(
                        query_rows
                    )
                )

                # ------------------------------------------------------------------
                # SOW DROPDOWN
                # ------------------------------------------------------------------

                sow_col = find_column(
                    df_dropdown,
                    [
                        "Master SOW",
                        "SOW"
                    ]
                )

                if sow_col:

                    sow_dropdown_list = (
                        df_dropdown[sow_col]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .unique()
                        .tolist()
                    )

                else:

                    sow_dropdown_list = []

                if not sow_dropdown_list:

                    sow_dropdown_list = [
                        "Survey BSS",
                        "Instalasi BSS",
                        "Instalasi EVC",
                    ]

                # ------------------------------------------------------------------
                # INPUT
                # ------------------------------------------------------------------

                col1, col2, col3 = st.columns(3)

                with col1:

                    selected_sow_type = st.selectbox(

                        "Pilih Jenis SOW",

                        options=sow_dropdown_list,

                        key="sow_type_select",
                    )

                    is_survey = (
                        "survey"
                        in str(
                            selected_sow_type
                        ).lower()
                    )

                    auto_pekerjaan = (
                        "Survey Location"
                        if is_survey
                        else "Construction"
                    )

                    target_wo_col_idx = (
                        11
                        if is_survey
                        else 22
                    )

                    if (
                        df_query.shape[1]
                        > target_wo_col_idx
                    ):

                        raw_wos = (
                            df_query
                            .iloc[:, target_wo_col_idx]
                            .dropna()
                            .unique()
                        )

                        wo_list = [

                            str(wo).strip()

                            for wo in raw_wos

                            if str(wo).strip()
                            and str(wo).strip().lower()
                            != "nan"
                        ]

                    else:

                        wo_list = []

                    wo_label = (
                        "Pilih Nomor WO "
                        f"({'Kolom L - Survey' if is_survey else 'Kolom W - Cons'})"
                    )

                    selected_wo = st.selectbox(

                        wo_label,

                        options=(
                            wo_list
                            if wo_list
                            else [
                                "- Tidak ada WO -"
                            ]
                        ),

                        key=f"wo_select_{selected_sow_type}",
                    )

                    proyek_input = st.text_input(
                        "Proyek",
                        value="V-Green"
                    )

                    pekerjaan_input = st.text_input(

                        "Pekerjaan",

                        value=auto_pekerjaan,

                        key=(
                            f"pekerjaan_input_"
                            f"{selected_sow_type}"
                        ),
                    )

                # ------------------------------------------------------------------
                # MITRA
                # ------------------------------------------------------------------

                with col2:

                    mitra_col_name = (
                        df_dropdown.columns[0]
                        if not df_dropdown.empty
                        else None
                    )

                    if mitra_col_name:

                        mitra_list = (
                            df_dropdown[
                                mitra_col_name
                            ]
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .unique()
                            .tolist()
                        )

                    else:

                        mitra_list = []

                    selected_pic = st.selectbox(

                        "Penanggung Jawab (Mitra)",

                        options=(
                            mitra_list
                            if mitra_list
                            else ["Edy"]
                        ),
                    )

                    default_phone = DEFAULT_PHONE

                    if (
                        not df_dropdown.empty
                        and len(
                            df_dropdown.columns
                        ) > 1
                    ):

                        phone_col_name = (
                            df_dropdown.columns[1]
                        )

                        matched_row = (
                            df_dropdown[
                                df_dropdown[
                                    mitra_col_name
                                ]
                                .astype(str)
                                .str.strip()
                                ==
                                selected_pic
                            ]
                        )

                        if not matched_row.empty:

                            default_phone = str(
                                matched_row.iloc[0][
                                    phone_col_name
                                ]
                            ).strip()

                    pic_phone = st.text_input(

                        "No. Telepon Penanggung Jawab",

                        value=default_phone,
                    )

                # ------------------------------------------------------------------
                # PUBLISHER
                # ------------------------------------------------------------------

                with col3:

                    st.markdown(
                        "**Kontak Penerbit SPK (PT. CLX):**"
                    )

                    st.markdown(
                        "1. Wikantiyoso Suyono "
                        "(0878-8855-0300)"
                    )

                    st.markdown(
                        "2. Hartono "
                        "(0818-0690-9317)"
                    )

                # ------------------------------------------------------------------
                # FILTER SITE
                # ------------------------------------------------------------------

                if (
                    selected_wo
                    and selected_wo
                    != "- Tidak ada WO -"
                ):

                    filtered_df = (
                        df_query[
                            df_query.iloc[
                                :,
                                target_wo_col_idx
                            ]
                            .astype(str)
                            .str.strip()
                            ==
                            selected_wo
                        ]
                        .copy()
                    )

                    filtered_df[
                        "col_charge"
                    ] = (
                        filtered_df.iloc[:, 2]
                        if filtered_df.shape[1] > 2
                        else "-"
                    )

                    filtered_df[
                        "col_site"
                    ] = (
                        filtered_df.iloc[:, 5]
                        if filtered_df.shape[1] > 5
                        else "-"
                    )

                    filtered_df[
                        "col_gmaps"
                    ] = (
                        filtered_df.iloc[:, 7]
                        if filtered_df.shape[1] > 7
                        else "-"
                    )

                    filtered_df[
                        "col_province"
                    ] = (
                        filtered_df.iloc[:, 8]
                        if filtered_df.shape[1] > 8
                        else "-"
                    )

                    filtered_df[
                        "col_pic"
                    ] = (
                        filtered_df.iloc[:, 10]
                        if filtered_df.shape[1] > 10
                        else "-"
                    )

                    st.markdown("---")

                    st.markdown(
                        f"### 📍 Daftar Site untuk WO: "
                        f"`{selected_wo}`"
                    )

                    if "Pilih" not in filtered_df.columns:

                        filtered_df.insert(
                            0,
                            "Pilih",
                            True
                        )

                    edited_df = st.data_editor(

                        filtered_df,

                        use_container_width=True,

                        hide_index=True,

                        key=(
                            f"site_editor_"
                            f"{selected_wo}"
                        ),
                    )

                    st.markdown("---")

                    # ------------------------------------------------------------------
                    # GENERATE
                    # ------------------------------------------------------------------

                    if st.button(

                        "🚀 Generate SPK & "
                        "Save to Database Sheet",

                        type="primary",
                    ):

                        selected_sites = (

                            edited_df[
                                edited_df[
                                    "Pilih"
                                ] == True
                            ]

                            if "Pilih"
                            in edited_df.columns

                            else pd.DataFrame()
                        )

                        if selected_sites.empty:

                            st.error(
                                "❌ Silakan centang "
                                "minimal satu site "
                                "terlebih dahulu."
                            )

                        else:

                            with st.spinner(
                                "Memproses dokumen "
                                "PDF SPK & memperbarui "
                                "Database..."
                            ):

                                # ----------------------------------------------------------
                                # MATCH SOW
                                # ----------------------------------------------------------

                                matched_sow_df = (
                                    pd.DataFrame()
                                )

                                if (
                                    not df_master_sow.empty
                                ):

                                    kode_col = (
                                        df_master_sow
                                        .columns[0]
                                    )

                                    matched_sow_df = (
                                        df_master_sow[
                                            df_master_sow[
                                                kode_col
                                            ]
                                            .astype(str)
                                            .str.strip()
                                            .str.lower()
                                            ==
                                            str(
                                                selected_sow_type
                                            )
                                            .strip()
                                            .lower()
                                        ]
                                    )

                                # ----------------------------------------------------------
                                # LOCATION
                                # ----------------------------------------------------------

                                provinces = (
                                    selected_sites[
                                        "col_province"
                                    ]
                                    .dropna()
                                    .astype(str)
                                    .str.strip()
                                    .tolist()
                                )

                                unique_provinces = sorted(
                                    list(
                                        set(
                                            [
                                                p
                                                for p
                                                in provinces
                                                if p
                                                and p != "-"
                                            ]
                                        )
                                    )
                                )

                                auto_lokasi = (

                                    ", ".join(
                                        unique_provinces
                                    )

                                    if unique_provinces

                                    else "Jawa Tengah"
                                )

                                # ----------------------------------------------------------
                                # TARGET DB
                                # ----------------------------------------------------------

                                target_db_name = (
                                    "DB SPK Survey"
                                    if is_survey
                                    else "DB SPK Cons"
                                )

                                try:

                                    target_sheet = (
                                        sh.worksheet(
                                            target_db_name
                                        )
                                    )

                                except Exception:

                                    target_sheet = (
                                        sh.add_worksheet(
                                            title=target_db_name,
                                            rows=1000,
                                            cols=20,
                                        )
                                    )

                                # IMPORTANT:
                                # memastikan kolom Sign COO
                                ensure_database_columns(
                                    target_sheet
                                )

                                # ----------------------------------------------------------
                                # SEQUENCE
                                # ----------------------------------------------------------

                                existing_rows = (
                                    target_sheet
                                    .get_all_values()
                                )

                                existing_df = (
                                    dataframe_from_sheet_rows(
                                        existing_rows
                                    )
                                )

                                spk_col = find_column(
                                    existing_df,
                                    [
                                        "No. SPK",
                                        "No SPK",
                                        "SPK"
                                    ]
                                )

                                if (
                                    spk_col
                                    and not existing_df.empty
                                ):

                                    spk_ids = (
                                        existing_df[
                                            spk_col
                                        ]
                                        .astype(str)
                                        .str.strip()
                                    )

                                    spk_ids = [
                                        x
                                        for x in spk_ids
                                        if x
                                        and x.lower()
                                        != "nan"
                                    ]

                                    seq_number = (
                                        len(
                                            set(spk_ids)
                                        ) + 1
                                    )

                                else:

                                    seq_number = 1

                                no_spk = (
                                    generate_spk_number(
                                        selected_sow_type,
                                        seq_number
                                    )
                                )

                                # ----------------------------------------------------------
                                # METADATA
                                # ----------------------------------------------------------

                                spk_metadata = {

                                    "no_spk": no_spk,

                                    "proyek":
                                        proyek_input,

                                    "pekerjaan":
                                        pekerjaan_input,

                                    "lokasi":
                                        auto_lokasi,

                                    "pic_name":
                                        selected_pic,

                                    "pic_phone":
                                        pic_phone,

                                    "sow_type":
                                        selected_sow_type,
                                }

                                # ----------------------------------------------------------
                                # PDF DRAFT
                                # ----------------------------------------------------------

                                no_spk, pdf_bytes = (
                                    generate_spk_pdf_bytes(

                                        selected_wo,

                                        selected_sites,

                                        spk_metadata,

                                        matched_sow_df,

                                        approved=False,
                                    )
                                )

                                safe_filename = (
                                    f"SPK_"
                                    f"{no_spk.replace('/', '_')}"
                                    f".pdf"
                                )

                                current_date_str = (
                                    datetime.datetime
                                    .now()
                                    .strftime(
                                        "%d/%m/%Y"
                                    )
                                )

                                current_time_str = (
                                    datetime.datetime
                                    .now()
                                    .strftime(
                                        "%Y-%m-%d %H:%M:%S"
                                    )
                                )

                                new_rows = []

                                for _, row in (
                                    selected_sites
                                    .iterrows()
                                ):

                                    site_name_val = (
                                        row.get(
                                            "col_site",
                                            "-"
                                        )
                                    )

                                    charger_type_val = (
                                        row.get(
                                            "col_charge",
                                            "-"
                                        )
                                    )

                                    new_rows.append([

                                        current_date_str,

                                        no_spk,

                                        str(
                                            selected_wo
                                        ),

                                        pekerjaan_input,

                                        "PT CLX",

                                        site_name_val,

                                        charger_type_val,

                                        selected_sow_type,

                                        current_time_str,

                                        selected_pic,

                                        "",

                                        "",

                                        "",

                                        "Pending COO Approval",

                                        "",
                                    ])

                                if new_rows:

                                    # Ambil posisi header aktual
                                    headers_now = (
                                        normalize_headers(
                                            target_sheet
                                            .get_all_values()[0]
                                        )
                                    )

                                    target_sheet.append_rows(
                                        new_rows
                                    )

                                st.success(

                                    f"✅ SPK `{no_spk}` "
                                    "berhasil dibuat & "
                                    "dikirim ke COO "
                                    "Dashboard untuk "
                                    "Approval!"
                                )

                                st.download_button(

                                    label=(
                                        "📥 Download "
                                        "File PDF SPK "
                                        "(Draft)"
                                    ),

                                    data=pdf_bytes,

                                    file_name=safe_filename,

                                    mime="application/pdf",

                                    type="primary",
                                )

        except Exception as e:

            st.error(
                "❌ Terjadi kesalahan pada "
                f"Menu Generate SPK: {e}"
            )

    # ==========================================================================
    # MENU 2 - TAKE OVER
    # ==========================================================================

    elif main_menu == "🔄 Take Over Site":

        st.subheader(
            "🔄 Menu Take Over Site"
        )

        st.markdown(
            "Pilih **Sheet Target** "
            "(Survey / Cons) dan **Site Name** "
            "yang akan di-*Take Over*."
        )

        try:

            db_target_type = st.radio(

                "Pilih Kategori Database SPK",

                options=[
                    "DB SPK Cons",
                    "DB SPK Survey"
                ],

                horizontal=True,
            )

            target_sheet_to = (
                sh.worksheet(
                    db_target_type
                )
            )

            to_rows = (
                target_sheet_to
                .get_all_values()
            )

            if len(to_rows) <= 1:

                st.warning(
                    f"⚠️ Sheet `{db_target_type}` "
                    "belum memiliki data."
                )

            else:

                df_to = (
                    dataframe_from_sheet_rows(
                        to_rows
                    )
                )

                col_site_name = find_column(
                    df_to,
                    [
                        "Site Name",
                        "Site"
                    ],
                    "Site Name"
                )

                col_spk_num = find_column(
                    df_to,
                    [
                        "No. SPK",
                        "No SPK"
                    ],
                    "No. SPK"
                )

                col_mitra_name = find_column(
                    df_to,
                    [
                        "Mitra"
                    ],
                    "Mitra"
                )

                site_options = (
                    df_to[
                        col_site_name
                    ]
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                )

                site_options = [
                    s
                    for s in site_options
                    if s.strip()
                    not in ["", "-"]
                ]

                col_to1, col_to2 = (
                    st.columns(2)
                )

                with col_to1:

                    selected_site_to = (
                        st.selectbox(

                            "Pilih Site Name "
                            "yang akan di "
                            "Take Over",

                            options=(
                                site_options
                                if site_options
                                else [
                                    "- Tidak Ada Site -"
                                ]
                            ),
                        )
                    )

                if (
                    selected_site_to
                    and
                    selected_site_to
                    != "- Tidak Ada Site -"
                ):

                    matched_to_row = (
                        df_to[
                            df_to[
                                col_site_name
                            ]
                            .astype(str)
                            ==
                            str(
                                selected_site_to
                            )
                        ]
                    )

                    old_spk = (
                        matched_to_row.iloc[0][
                            col_spk_num
                        ]
                        if not matched_to_row.empty
                        else "-"
                    )

                    old_mitra = (
                        matched_to_row.iloc[0][
                            col_mitra_name
                        ]
                        if not matched_to_row.empty
                        else "-"
                    )

                    with col_to2:

                        st.info(
                            f"📌 **No SPK Lama:** "
                            f"{old_spk}"
                        )

                        st.info(
                            f"👤 **Mitra Lama:** "
                            f"{old_mitra}"
                        )

                    st.markdown(
                        "### 📝 Informasi "
                        "Take Over Baru"
                    )

                    col_to3, col_to4 = (
                        st.columns(2)
                    )

                    try:

                        dropdown_sheet = (
                            sh.worksheet(
                                "Master Dropdown"
                            )
                        )

                        dropdown_data = (
                            dropdown_sheet
                            .get_all_values()
                        )

                        df_dropdown = (
                            dataframe_from_sheet_rows(
                                dropdown_data
                            )
                        )

                        mitra_col = (
                            df_dropdown.columns[0]
                        )

                        new_mitra_list = (
                            df_dropdown[
                                mitra_col
                            ]
                            .dropna()
                            .astype(str)
                            .unique()
                            .tolist()
                        )

                    except Exception:

                        new_mitra_list = []

                    with col_to3:

                        new_mitra = (
                            st.selectbox(

                                "Pilih Mitra Baru "
                                "(Pengambil Alih)",

                                options=(
                                    new_mitra_list
                                    if new_mitra_list
                                    else [
                                        "Mitra Baru"
                                    ]
                                ),
                            )
                        )

                        sow_type_to = (
                            "Survey BSS"
                            if "Survey"
                            in db_target_type
                            else "Instalasi BSS"
                        )

                    with col_to4:

                        today_date_str = (
                            datetime.datetime
                            .now()
                            .strftime(
                                "%d/%m/%Y"
                            )
                        )

                        st.text_input(

                            "Tanggal Take Over "
                            "(Otomatis)",

                            value=today_date_str,

                            disabled=True,
                        )

                    if st.button(

                        "🔥 Proses & Simpan "
                        "Take Over Site",

                        type="primary",
                    ):

                        with st.spinner(
                            "Memproses Take Over..."
                        ):

                            spk_ids = []

                            for row in to_rows[1:]:

                                if (
                                    len(row) > 1
                                    and row[1].strip()
                                ):

                                    spk_ids.append(
                                        row[1].strip()
                                    )

                            seq_number = (
                                len(
                                    set(spk_ids)
                                ) + 1
                            )

                            new_spk_no = (
                                generate_spk_number(
                                    sow_type_to,
                                    seq_number
                                )
                            )

                            # Cari posisi kolom
                            headers = (
                                normalize_headers(
                                    to_rows[0]
                                )
                            )

                            date_to_col = (
                                headers.index(
                                    "Date SPK (Take Over)"
                                ) + 1
                                if
                                "Date SPK (Take Over)"
                                in headers
                                else 11
                            )

                            old_spk_col = (
                                headers.index(
                                    "No. SPK (Take Over)"
                                ) + 1
                                if
                                "No. SPK (Take Over)"
                                in headers
                                else 12
                            )

                            new_mitra_col = (
                                headers.index(
                                    "Mitra (Take Over)"
                                ) + 1
                                if
                                "Mitra (Take Over)"
                                in headers
                                else 13
                            )

                            for idx, row in enumerate(
                                to_rows[1:],
                                start=2
                            ):

                                if (
                                    len(row) > 5
                                    and str(
                                        row[5]
                                    )
                                    == str(
                                        selected_site_to
                                    )
                                ):

                                    target_sheet_to.update_cell(
                                        idx,
                                        date_to_col,
                                        today_date_str
                                    )

                                    target_sheet_to.update_cell(
                                        idx,
                                        old_spk_col,
                                        old_spk
                                    )

                                    target_sheet_to.update_cell(
                                        idx,
                                        new_mitra_col,
                                        new_mitra
                                    )

                                    break

                            st.success(

                                f"🎉 Take Over "
                                "Berhasil!\n\n"
                                f"• **Site:** "
                                f"{selected_site_to}\n"
                                f"• **No. SPK Baru:** "
                                f"`{new_spk_no}`\n"
                                f"• **SPK Lama:** "
                                f"`{old_spk}`\n"
                                f"• **Mitra Baru:** "
                                f"{new_mitra}"
                            )

        except Exception as e:

            st.error(
                "❌ Terjadi kesalahan pada "
                f"Menu Take Over: {e}"
            )

    # ==========================================================================
    # MENU 3 - COO APPROVAL
    # ==========================================================================

    elif main_menu == "🔒 COO Approval Dashboard":

        st.subheader(
            "🔒 COO Approval Dashboard "
            "(Otorisasi SPK)"
        )

        st.markdown(
            "Halaman khusus Manajemen/COO "
            "untuk menyetujui (**Approve**) "
            "atau menolak (**Reject**) "
            "pengajuan SPK."
        )

        pin_input = st.text_input(

            "Masukkan PIN Khusus COO:",

            type="password",

            placeholder="Masukkan PIN...",
        )

        if pin_input == COO_PIN_SECRET:

            st.success(
                "🔓 Akses Diterima! "
                "Selamat datang, COO."
            )

            db_approval_type = st.radio(

                "Pilih Kategori Database SPK",

                options=[
                    "DB SPK Cons",
                    "DB SPK Survey"
                ],

                horizontal=True,

                key="approval_db_choice",
            )

            try:

                sheet_appr = (
                    sh.worksheet(
                        db_approval_type
                    )
                )

                appr_rows = (
                    sheet_appr
                    .get_all_values()
                )

                if len(appr_rows) <= 1:

                    st.info(
                        f"ℹ️ Belum ada data pada "
                        f"sheet `{db_approval_type}`."
                    )

                else:

                    df_appr = (
                        dataframe_from_sheet_rows(
                            appr_rows
                        )
                    )

                    # Pastikan kolom
                    ensure_database_columns(
                        sheet_appr
                    )

                    # Reload setelah memastikan kolom
                    appr_rows = (
                        sheet_appr
                        .get_all_values()
                    )

                    df_appr = (
                        dataframe_from_sheet_rows(
                            appr_rows
                        )
                    )

                    status_col = find_column(
                        df_appr,
                        [
                            "Status Approval"
                        ]
                    )

                    spk_col = find_column(
                        df_appr,
                        [
                            "No. SPK",
                            "No SPK"
                        ]
                    )

                    if not status_col:

                        st.error(
                            "❌ Kolom "
                            "`Status Approval` "
                            "tidak ditemukan."
                        )

                    elif not spk_col:

                        st.error(
                            "❌ Kolom "
                            "`No. SPK` "
                            "tidak ditemukan."
                        )

                    else:

                        pending_df = (
                            df_appr[
                                df_appr[
                                    status_col
                                ]
                                .astype(str)
                                .str.strip()
                                .str.lower()
                                ==
                                "pending coo approval"
                            ]
                        )

                        st.markdown("---")

                        st.markdown(
                            "### 📋 Daftar SPK "
                            f"Menunggu Persetujuan "
                            f"({len(pending_df)} Pending)"
                        )

                        if pending_df.empty:

                            st.success(
                                "🎉 Semua pengajuan "
                                "SPK telah diproses! "
                                "Tidak ada antrean "
                                "pending."
                            )

                        else:

                            pending_spks = (
                                pending_df[
                                    spk_col
                                ]
                                .dropna()
                                .astype(str)
                                .str.strip()
                                .unique()
                                .tolist()
                            )

                            selected_appr_spk = (
                                st.selectbox(

                                    "Pilih SPK "
                                    "yang akan ditinjau:",

                                    options=pending_spks,
                                )
                            )

                            if selected_appr_spk:

                                spk_details = (
                                    pending_df[
                                        pending_df[
                                            spk_col
                                        ]
                                        .astype(str)
                                        .str.strip()
                                        ==
                                        selected_appr_spk
                                    ]
                                )

                                st.markdown(
                                    "#### 🔍 Detail "
                                    "Informasi SPK:"
                                )

                                # Hindari error
                                # ['Date SPK'] not in index
                                display_candidates = [
                                    "Date SPK",
                                    "No. SPK",
                                    "No. WO",
                                    "Pekerjaan",
                                    "Site Name",
                                    "Mitra",
                                ]

                                display_columns = []

                                for c in display_candidates:

                                    actual = find_column(
                                        spk_details,
                                        [c]
                                    )

                                    if actual:
                                        display_columns.append(
                                            actual
                                        )

                                if display_columns:

                                    st.dataframe(
                                        spk_details[
                                            display_columns
                                        ],
                                        use_container_width=True,
                                        hide_index=True,
                                    )

                                else:

                                    st.dataframe(
                                        spk_details,
                                        use_container_width=True,
                                        hide_index=True,
                                    )

                                col_btn1, col_btn2, _ = (
                                    st.columns(
                                        [1, 1, 2]
                                    )
                                )

                                # ------------------------------------------------------
                                # APPROVE
                                # ------------------------------------------------------

                                with col_btn1:

                                    if st.button(

                                        "✅ APPROVE SPK",

                                        type="primary",

                                        key=(
                                            "approve_"
                                            f"{selected_appr_spk}"
                                        ),
                                    ):

                                        with st.spinner(
                                            "Memproses "
                                            "Approval..."
                                        ):

                                            headers = (
                                                normalize_headers(
                                                    appr_rows[0]
                                                )
                                            )

                                            status_col_idx = (
                                                headers.index(
                                                    status_col
                                                ) + 1
                                            )

                                            sign_col = (
                                                find_column(
                                                    df_appr,
                                                    [
                                                        "Sign COO"
                                                    ]
                                                )
                                            )

                                            if sign_col:

                                                sign_col_idx = (
                                                    headers.index(
                                                        sign_col
                                                    ) + 1
                                                )

                                            else:

                                                sign_col_idx = (
                                                    len(
                                                        headers
                                                    ) + 1
                                                )

                                                sheet_appr.update_cell(
                                                    1,
                                                    sign_col_idx,
                                                    "Sign COO"
                                                )

                                            updated_count = 0

                                            for idx, row in enumerate(
                                                appr_rows[1:],
                                                start=2
                                            ):

                                                if (
                                                    len(row) > 1
                                                    and
                                                    str(
                                                        row[1]
                                                    ).strip()
                                                    ==
                                                    str(
                                                        selected_appr_spk
                                                    ).strip()
                                                ):

                                                    sheet_appr.update_cell(
                                                        idx,
                                                        status_col_idx,
                                                        "Approved by COO"
                                                    )

                                                    sheet_appr.update_cell(
                                                        idx,
                                                        sign_col_idx,
                                                        "Approved CFO.png"
                                                    )

                                                    updated_count += 1

                                            if updated_count:

                                                st.success(

                                                    f"✅ SPK "
                                                    f"`{selected_appr_spk}` "
                                                    "BERHASIL DI-APPROVE!"
                                                )

                                                st.info(
                                                    "✍️ Digital signature "
                                                    "COO telah ditandai "
                                                    "pada kolom `Sign COO`."
                                                )

                                                st.rerun()

                                            else:

                                                st.error(
                                                    "❌ Data SPK tidak "
                                                    "ditemukan."
                                                )

                                # ------------------------------------------------------
                                # REJECT
                                # ------------------------------------------------------

                                with col_btn2:

                                    if st.button(

                                        "❌ REJECT SPK",

                                        key=(
                                            "reject_"
                                            f"{selected_appr_spk}"
                                        ),
                                    ):

                                        with st.spinner(
                                            "Memproses "
                                            "Penolakan..."
                                        ):

                                            headers = (
                                                normalize_headers(
                                                    appr_rows[0]
                                                )
                                            )

                                            status_col_idx = (
                                                headers.index(
                                                    status_col
                                                ) + 1
                                            )

                                            sign_col = (
                                                find_column(
                                                    df_appr,
                                                    [
                                                        "Sign COO"
                                                    ]
                                                )
                                            )

                                            if sign_col:

                                                sign_col_idx = (
                                                    headers.index(
                                                        sign_col
                                                    ) + 1
                                                )

                                            else:

                                                sign_col_idx = (
                                                    len(
                                                        headers
                                                    ) + 1
                                                )

                                                sheet_appr.update_cell(
                                                    1,
                                                    sign_col_idx,
                                                    "Sign COO"
                                                )

                                            for idx, row in enumerate(
                                                appr_rows[1:],
                                                start=2
                                            ):

                                                if (
                                                    len(row) > 1
                                                    and
                                                    str(
                                                        row[1]
                                                    ).strip()
                                                    ==
                                                    str(
                                                        selected_appr_spk
                                                    ).strip()
                                                ):

                                                    sheet_appr.update_cell(
                                                        idx,
                                                        status_col_idx,
                                                        "Rejected by COO"
                                                    )

                                                    sheet_appr.update_cell(
                                                        idx,
                                                        sign_col_idx,
                                                        ""
                                                    )

                                            st.warning(
                                                f"❌ SPK "
                                                f"`{selected_appr_spk}` "
                                                "DITOLAK!"
                                            )

                                            st.rerun()

            except Exception as e:

                st.error(
                    "❌ Terjadi kesalahan pada "
                    "Dashboard Approval COO: "
                    f"{e}"
                )

        elif pin_input != "":

            st.error(
                "❌ PIN Salah! Akses ditolak."
            )

    # ==========================================================================
    # MENU 4 - DOWNLOAD SPK
    # ==========================================================================

    elif main_menu == "📥 Download SPK":

        st.subheader(
            "📥 Download SPK Approved"
        )

        st.markdown(
            "Menu ini hanya menampilkan SPK "
            "yang telah **Approved by COO**. "
            "PDF yang dihasilkan merupakan "
            "**Final SPK** dengan digital "
            "signature COO."
        )

        # ----------------------------------------------------------------------
        # Cek signature
        # ----------------------------------------------------------------------

        signature_path = get_signature_path()

        if signature_path:

            st.success(
                "✍️ Digital Signature COO ditemukan: "
                "`assets/Approved CFO.png`"
            )

        else:

            st.warning(
                "⚠️ File digital signature tidak ditemukan: "
                "`assets/Approved CFO.png`"
            )

        # ----------------------------------------------------------------------
        # DB TYPE
        # ----------------------------------------------------------------------

        db_download_type = st.radio(

            "Pilih Kategori Database SPK",

            options=[
                "DB SPK Cons",
                "DB SPK Survey"
            ],

            horizontal=True,

            key="download_db_choice",
        )

        try:

            sheet_download = (
                sh.worksheet(
                    db_download_type
                )
            )

            download_rows = (
                sheet_download
                .get_all_values()
            )

            if len(download_rows) <= 1:

                st.info(
                    f"ℹ️ Belum ada data pada "
                    f"sheet `{db_download_type}`."
                )

            else:

                df_download = (
                    dataframe_from_sheet_rows(
                        download_rows
                    )
                )

                spk_col = find_column(
                    df_download,
                    [
                        "No. SPK",
                        "No SPK",
                        "SPK"
                    ]
                )

                status_col = find_column(
                    df_download,
                    [
                        "Status Approval",
                        "Approval Status"
                    ]
                )

                if not spk_col:

                    st.error(
                        "❌ Kolom `No. SPK` "
                        "tidak ditemukan."
                    )

                elif not status_col:

                    st.error(
                        "❌ Kolom `Status Approval` "
                        "tidak ditemukan."
                    )

                else:

                    approved_df = (
                        df_download[
                            df_download[
                                status_col
                            ]
                            .astype(str)
                            .str.strip()
                            .str.lower()
                            .str.contains(
                                "approved"
                            )
                        ]
                    )

                    approved_spks = (
                        approved_df[
                            spk_col
                        ]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .unique()
                        .tolist()
                    )

                    st.markdown(
                        f"### 📋 SPK Approved "
                        f"({len(approved_spks)})"
                    )

                    if not approved_spks:

                        st.info(
                            "ℹ️ Belum ada SPK "
                            "yang Approved oleh COO."
                        )

                    else:

                        selected_download_spk = (
                            st.selectbox(

                                "Pilih SPK yang "
                                "akan didownload",

                                options=approved_spks,

                                key="download_spk_select",
                            )
                        )

                        selected_download_rows = (
                            approved_df[
                                approved_df[
                                    spk_col
                                ]
                                .astype(str)
                                .str.strip()
                                ==
                                selected_download_spk
                            ]
                        )

                        # --------------------------------------------------------------
                        # SUMMARY
                        # --------------------------------------------------------------

                        st.markdown(
                            "#### 📄 Informasi SPK"
                        )

                        summary_cols = []

                        for candidate in [
                            "Date SPK",
                            "No. SPK",
                            "No. WO",
                            "Pekerjaan",
                            "EPC",
                            "Site Name",
                            "Mitra",
                            "Status Approval",
                            "Sign COO",
                        ]:

                            actual = find_column(
                                selected_download_rows,
                                [candidate]
                            )

                            if actual:
                                summary_cols.append(
                                    actual
                                )

                        if summary_cols:

                            st.dataframe(
                                selected_download_rows[
                                    summary_cols
                                ],
                                use_container_width=True,
                                hide_index=True,
                            )

                        # --------------------------------------------------------------
                        # GENERATE FINAL PDF
                        # --------------------------------------------------------------

                        if st.button(

                            "📄 Generate & Download "
                            "Approved SPK",

                            type="primary",

                            key=(
                                "generate_approved_"
                                f"{selected_download_spk}"
                            ),
                        ):

                            try:

                                with st.spinner(
                                    "Membuat Final PDF "
                                    "dengan digital "
                                    "signature COO..."
                                ):

                                    no_spk, approved_pdf = (
                                        build_approved_pdf(

                                            sh,

                                            db_download_type,

                                            selected_download_spk,
                                        )
                                    )

                                    safe_filename = (
                                        "SPK_"
                                        f"{no_spk.replace('/', '_')}"
                                        "_APPROVED.pdf"
                                    )

                                    st.success(
                                        "✅ Final SPK berhasil "
                                        "dibuat dengan digital "
                                        "signature COO."
                                    )

                                    st.download_button(

                                        label=(
                                            "📥 DOWNLOAD "
                                            "FINAL SPK PDF"
                                        ),

                                        data=approved_pdf,

                                        file_name=safe_filename,

                                        mime="application/pdf",

                                        type="primary",

                                        key=(
                                            "download_final_"
                                            f"{selected_download_spk}"
                                        ),
                                    )

                            except Exception as e:

                                st.error(
                                    "❌ Gagal membuat "
                                    f"Approved PDF: {e}"
                                )

        except Exception as e:

            st.error(
                "❌ Terjadi kesalahan pada "
                "Menu Download SPK: "
                f"{e}"
            )


# ==============================================================================
# RUN
# ==============================================================================

if __name__ == "__main__":
    show_spk_page()
