import datetime
import io
import os
from typing import List

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

from core.database import get_google_sheet_connection


# ==============================================================================
# CONFIGURATION
# ==============================================================================

COO_PIN_SECRET = "1234"

SHEET_QUERY = "Query"
SHEET_DROPDOWN = "Master Dropdown"
SHEET_SOW = "Master SOW"

SHEET_SURVEY = "DB SPK Survey"
SHEET_CONS = "DB SPK Cons"

STATUS_PENDING = "Pending COO Approval"
STATUS_APPROVED = "Approved by COO"
STATUS_REJECTED = "Rejected by COO"

CACHE_TTL_SHEET = 120


# ==============================================================================
# MAINTENANCE SOW
# ==============================================================================

MAINTENANCE_SOW_EVCS = "Maintenance Service EVCS"
MAINTENANCE_SOW_BSS = "Maintenance Service BSS"

MAINTENANCE_SOW_LIST = [
    MAINTENANCE_SOW_EVCS,
    MAINTENANCE_SOW_BSS,
]


# ==============================================================================
# STANDARD HEADERS
# ==============================================================================

DEFAULT_SPK_HEADERS = [
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
]


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

    div[data-testid="stAlert"] {
        border-radius: 8px !important;
    }
</style>
"""


# ==============================================================================
# CONNECTION
# ==============================================================================

@st.cache_resource(show_spinner=False)
def get_cached_connection():
    """
    Satu koneksi Google Sheets untuk session/process.
    """
    return get_google_sheet_connection()


# ==============================================================================
# GOOGLE SHEETS READ CACHE
# ==============================================================================

@st.cache_data(
    ttl=CACHE_TTL_SHEET,
    show_spinner=False,
)
def load_sheet_values_cached(sheet_name: str) -> List[List[str]]:
    """
    Membaca satu worksheet dan melakukan cache.

    Tidak menerima object worksheet sebagai argument.
    Hanya menerima nama sheet.
    """

    sh = get_cached_connection()

    if not sh:
        return []

    worksheet = sh.worksheet(sheet_name)

    return worksheet.get_all_values()


def clear_sheet_cache(sheet_name: str = None):
    """
    Clear cache pembacaan sheet.
    """

    try:
        load_sheet_values_cached.clear()
    except Exception:
        pass


def get_worksheet(sheet_name: str):
    """
    Mendapatkan worksheet tanpa melakukan read tambahan.
    """

    sh = get_cached_connection()

    if not sh:
        return None

    return sh.worksheet(sheet_name)


# ==============================================================================
# GENERIC HELPERS
# ==============================================================================

def normalize_header(value):
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\n", " ")
    text = text.replace("\r", " ")
    text = text.strip()
    text = " ".join(text.split())

    return text


def normalize_dataframe_headers(df):
    if df is None or df.empty:
        return df

    df = df.copy()

    df.columns = [
        normalize_header(col)
        for col in df.columns
    ]

    return df


def find_column(
    df,
    candidates,
    fallback=None,
):
    if df is None or df.empty:
        return fallback

    normalized_columns = {
        normalize_header(col).lower(): col
        for col in df.columns
    }

    # Exact match
    for candidate in candidates:

        candidate_norm = (
            normalize_header(candidate)
            .lower()
        )

        if candidate_norm in normalized_columns:
            return normalized_columns[candidate_norm]

    # Partial match
    for candidate in candidates:

        candidate_norm = (
            normalize_header(candidate)
            .lower()
        )

        for normalized_col, original_col in (
            normalized_columns.items()
        ):

            if (
                candidate_norm in normalized_col
                or normalized_col in candidate_norm
            ):
                return original_col

    return fallback


def safe_str(value, default=""):
    if value is None:
        return default

    try:
        if pd.isna(value):
            return default
    except Exception:
        pass

    text = str(value).strip()

    if text.lower() == "nan":
        return default

    return text


def dataframe_from_sheet_rows(rows):
    """
    Convert get_all_values() result menjadi DataFrame
    secara aman.
    """

    if not rows or len(rows) <= 1:
        return pd.DataFrame()

    headers = [
        normalize_header(h)
        for h in rows[0]
    ]

    max_cols = max(
        len(headers),
        max(
            [
                len(row)
                for row in rows[1:]
            ],
            default=0,
        ),
    )

    while len(headers) < max_cols:
        headers.append(
            f"Column_{len(headers) + 1}"
        )

    normalized_rows = []

    for row in rows[1:]:

        row_copy = list(row)

        while len(row_copy) < max_cols:
            row_copy.append("")

        normalized_rows.append(
            row_copy[:max_cols]
        )

    df = pd.DataFrame(
        normalized_rows,
        columns=headers,
    )

    return normalize_dataframe_headers(df)


def ensure_dataframe_columns(
    df,
    required_columns,
):
    """
    Pastikan DataFrame memiliki kolom tertentu.
    """

    df = df.copy()

    for col in required_columns:

        if col not in df.columns:
            df[col] = ""

    return df


# ==============================================================================
# MAINTENANCE HELPERS
# ==============================================================================

def is_maintenance_sow(sow_type):
    """
    True jika SOW adalah Maintenance Service EVCS/BSS.
    """

    sow = safe_str(sow_type).strip().lower()

    return sow in {
        MAINTENANCE_SOW_EVCS.lower(),
        MAINTENANCE_SOW_BSS.lower(),
    }


def is_maintenance_evcs(sow_type):
    return (
        safe_str(sow_type).strip().lower()
        == MAINTENANCE_SOW_EVCS.lower()
    )


def is_maintenance_bss(sow_type):
    return (
        safe_str(sow_type).strip().lower()
        == MAINTENANCE_SOW_BSS.lower()
    )


def get_sow_category(sow_type):
    """
    Mengembalikan kategori:
    - MAINTENANCE_EVCS
    - MAINTENANCE_BSS
    - SURVEY
    - CONS
    """

    sow = safe_str(sow_type).lower()

    if "maintenance service evcs" in sow:
        return "MAINTENANCE_EVCS"

    if "maintenance service bss" in sow:
        return "MAINTENANCE_BSS"

    if "survey" in sow:
        return "SURVEY"

    return "CONS"


def get_sow_category_label(sow_type):
    category = get_sow_category(sow_type)

    if category == "MAINTENANCE_EVCS":
        return "Maintenance Service EVCS"

    if category == "MAINTENANCE_BSS":
        return "Maintenance Service BSS"

    if category == "SURVEY":
        return "Survey"

    return "Construction"


def get_wo_column_index(sow_type):
    """
    Survey:
        Query kolom L = index 11

    Construction:
        Query kolom W = index 22

    Maintenance:
        Tidak menggunakan WO.
    """

    category = get_sow_category(sow_type)

    if category == "SURVEY":
        return 11

    if category == "CONS":
        return 22

    return None


# ==============================================================================
# BATCH WRITE HELPERS
# ==============================================================================

def batch_update_cells(
    worksheet,
    updates,
):
    """
    updates:
        [
            {
                "range": "N2",
                "values": [["Approved by COO"]]
            }
        ]

    Satu HTTP request batch untuk banyak cell.
    """

    if not updates:
        return

    worksheet.batch_update(
        updates,
        value_input_option="USER_ENTERED",
    )


def column_letter(column_number: int) -> str:
    """
    1 -> A
    2 -> B
    ...
    26 -> Z
    27 -> AA
    """

    result = ""

    while column_number > 0:

        column_number, remainder = divmod(
            column_number - 1,
            26,
        )

        result = (
            chr(65 + remainder)
            + result
        )

    return result


# ==============================================================================
# PIC
# ==============================================================================

def resolve_pic_name(
    selected_pic,
    manual_pic_name="",
):
    selected_pic = safe_str(
        selected_pic
    )

    manual_pic_name = safe_str(
        manual_pic_name
    )

    if selected_pic.upper() == "IN HOUSE":
        return manual_pic_name

    return selected_pic


# ==============================================================================
# GENERATE NOMOR SPK
# ==============================================================================

def generate_spk_number(
    sow_type="GENERAL",
    sequence_num=1,
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

    month_roman = roman_months[
        now.month - 1
    ]

    sow_lower = safe_str(
        sow_type
    ).lower()

    if "survey" in sow_lower:
        sow_code = "SURVEY"

    elif is_maintenance_evcs(sow_type):
        sow_code = "CONS"

    elif is_maintenance_bss(sow_type):
        sow_code = "CONS"

    else:
        sow_code = "CONS"

    seq_str = f"{sequence_num:04d}"

    return (
        f"{seq_str}/CLX/SPK/"
        f"{sow_code}/{month_roman}/{now.year}"
    )


# ==============================================================================
# SPK PDF
# ==============================================================================

def generate_spk_pdf_bytes(
    selected_wo,
    selected_sites,
    spk_metadata,
    matched_sow_df,
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
        textColor=colors.HexColor("#000000"),
        spaceAfter=2,
    )

    subtitle_style = ParagraphStyle(
        "DocSubtitle",
        parent=styles["Normal"],
        fontSize=9,
        leading=11,
        alignment=1,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#000000"),
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

    elements = []

    # --------------------------------------------------------------------------
    # KOP
    # --------------------------------------------------------------------------

    logo_path = "assets/CLX.png"

    if os.path.exists(logo_path):

        company_logo = Image(
            logo_path,
            width=1.5 * inch,
            height=1.0 * inch,
        )

    else:

        company_logo = Paragraph(
            "<b>PT. Connectivity Leads eXcellence</b>",
            header_left,
        )

    header_data = [
        [
            company_logo,
            Paragraph(
                "<b>PT. Connectivity Leads eXcellence</b><br/>"
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
            5.1 * inch,
        ],
    )

    t_header.setStyle(
        TableStyle(
            [
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
            ]
        )
    )

    elements.append(t_header)
    elements.append(Spacer(1, 8))

    # --------------------------------------------------------------------------
    # TITLE
    # --------------------------------------------------------------------------

    no_spk = spk_metadata.get(
        "no_spk",
        "0001/CLX/SPK/CONS/VII/2026",
    )

    date_str = datetime.datetime.now().strftime(
        "%d %B %Y"
    )

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

    # --------------------------------------------------------------------------
    # METADATA
    # --------------------------------------------------------------------------

    pic_name = safe_str(
        spk_metadata.get(
            "pic_name",
            "-",
        ),
        "-",
    )

    pic_phone = safe_str(
        spk_metadata.get(
            "pic_phone",
            "-",
        ),
        "-",
    )

    display_wo = safe_str(
        selected_wo,
        "-",
    )

    meta_table_data = [
        [
            Paragraph("Proyek:", meta_label),
            Paragraph(
                safe_str(
                    spk_metadata.get(
                        "proyek",
                        "-",
                    ),
                    "-",
                ),
                meta_val,
            ),
            Paragraph("No. WO:", meta_label),
            Paragraph(
                display_wo,
                meta_val,
            ),
        ],
        [
            Paragraph("Pekerjaan:", meta_label),
            Paragraph(
                safe_str(
                    spk_metadata.get(
                        "pekerjaan",
                        "-",
                    ),
                    "-",
                ),
                meta_val,
            ),
            Paragraph(
                "Penanggung Jawab:",
                meta_label,
            ),
            Paragraph(
                f"{pic_name} ({pic_phone})",
                meta_val,
            ),
        ],
        [
            Paragraph("Lokasi:", meta_label),
            Paragraph(
                safe_str(
                    spk_metadata.get(
                        "lokasi",
                        "-",
                    ),
                    "-",
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
            Paragraph(
                "Detail Site:",
                meta_label,
            ),
            Paragraph(
                (
                    "Tidak menggunakan Site Selection"
                    if is_maintenance_sow(
                        spk_metadata.get(
                            "sow_type",
                            "",
                        )
                    )
                    else "Terlampir pada tabel di bawah"
                ),
                meta_val,
            ),
            Paragraph("SOW:", meta_label),
            Paragraph(
                safe_str(
                    spk_metadata.get(
                        "sow_type",
                        "-",
                    ),
                    "-",
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
        TableStyle(
            [
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
            ]
        )
    )

    elements.append(t_meta)
    elements.append(Spacer(1, 10))

    # --------------------------------------------------------------------------
    # SOW
    # --------------------------------------------------------------------------

    sow_headers = [
        Paragraph("<b>No.</b>", cell_head),
        Paragraph(
            "<b>Uraian Pekerjaan (SoW)</b>",
            cell_head,
        ),
        Paragraph(
            "<b>Target Penyelesaian</b>",
            cell_head,
        ),
    ]

    sow_rows = [sow_headers]

    if (
        matched_sow_df is not None
        and not matched_sow_df.empty
    ):

        for idx, (_, row) in enumerate(
            matched_sow_df.iterrows(),
            1,
        ):

            deskripsi = (
                safe_str(
                    row.iloc[2],
                    "-",
                )
                if row.shape[0] > 2
                else "-"
            )

            target = (
                safe_str(
                    row.iloc[3],
                    "1 Hari",
                )
                if row.shape[0] > 3
                else "1 Hari"
            )

            sow_rows.append(
                [
                    Paragraph(
                        str(idx),
                        cell_body,
                    ),
                    Paragraph(
                        deskripsi,
                        cell_body,
                    ),
                    Paragraph(
                        target,
                        cell_body,
                    ),
                ]
            )

    else:

        sow_rows.append(
            [
                Paragraph("1", cell_body),
                Paragraph(
                    "SOW Pekerjaan Standar",
                    cell_body,
                ),
                Paragraph(
                    "1 Hari",
                    cell_body,
                ),
            ]
        )

    t_sow = Table(
        sow_rows,
        colWidths=[
            0.4 * inch,
            5.0 * inch,
            1.9 * inch,
        ],
    )

    t_sow.setStyle(
        TableStyle(
            [
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
            ]
        )
    )

    elements.append(t_sow)
    elements.append(Spacer(1, 10))

    # --------------------------------------------------------------------------
    # SITE
    # --------------------------------------------------------------------------

    if (
        not is_maintenance_sow(
            spk_metadata.get(
                "sow_type",
                "",
            )
        )
        and selected_sites is not None
    ):

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
            Paragraph("<b>No</b>", cell_head),
            Paragraph(
                "<b>Site Name</b>",
                cell_head,
            ),
            Paragraph(
                "<b>Charging Type</b>",
                cell_head,
            ),
            Paragraph(
                "<b>WO Number</b>",
                cell_head,
            ),
            Paragraph(
                "<b>Province</b>",
                cell_head,
            ),
            Paragraph(
                "<b>PIC + Contact</b>",
                cell_head,
            ),
            Paragraph(
                "<b>Gmaps</b>",
                cell_head,
            ),
        ]

        site_rows = [site_headers]

        for idx, (_, row) in enumerate(
            selected_sites.iterrows(),
            1,
        ):

            site_rows.append(
                [
                    Paragraph(
                        str(idx),
                        cell_body,
                    ),
                    Paragraph(
                        safe_str(
                            row.get(
                                "col_site",
                                "-",
                            ),
                            "-",
                        ),
                        cell_body,
                    ),
                    Paragraph(
                        safe_str(
                            row.get(
                                "col_charge",
                                "-",
                            ),
                            "-",
                        ),
                        cell_body,
                    ),
                    Paragraph(
                        safe_str(
                            selected_wo,
                            "-",
                        ),
                        cell_body,
                    ),
                    Paragraph(
                        safe_str(
                            row.get(
                                "col_province",
                                "-",
                            ),
                            "-",
                        ),
                        cell_body,
                    ),
                    Paragraph(
                        safe_str(
                            row.get(
                                "col_pic",
                                "-",
                            ),
                            "-",
                        ),
                        cell_body,
                    ),
                    Paragraph(
                        safe_str(
                            row.get(
                                "col_gmaps",
                                "-",
                            ),
                            "-",
                        ),
                        cell_body,
                    ),
                ]
            )

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
            TableStyle(
                [
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
                ]
            )
        )

        elements.append(t_site)
        elements.append(Spacer(1, 15))

    # --------------------------------------------------------------------------
    # SIGNATURE
    # --------------------------------------------------------------------------

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

    sign_data = [
        [
            Paragraph(
                "<b>PELAKSANA PEKERJAAN</b>",
                sign_style_bold,
            ),
            Paragraph(
                "<b>PEMBERI PERINTAH KERJA</b><br/>"
                "<b>PT. Connectivity Leads eXcellence</b>",
                sign_style_bold,
            ),
        ],
        [
            Spacer(1, 35),
            Spacer(1, 35),
        ],
        [
            Paragraph(
                f"<b>{pic_name}</b><br/>"
                f"Contact: {pic_phone}",
                sign_style_norm,
            ),
            Paragraph(
                "<b>Wikantiyoso Suyono</b><br/>"
                "Chief Operating Officer",
                sign_style_norm,
            ),
        ],
    ]

    t_sign = Table(
        sign_data,
        colWidths=[
            3.6 * inch,
            3.7 * inch,
        ],
    )

    t_sign.setStyle(
        TableStyle(
            [
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
            ]
        )
    )

    elements.append(
        KeepTogether([t_sign])
    )

    doc.build(elements)

    buffer.seek(0)

    return (
        no_spk,
        buffer.getvalue(),
    )


# ==============================================================================
# CREATE / ENSURE SPK SHEET
# ==============================================================================

def ensure_spk_sheet(
    sh,
    sheet_name,
):
    """
    Memastikan sheet SPK tersedia.

    Hanya membuat sheet jika benar-benar belum ada.
    """

    try:
        return sh.worksheet(sheet_name)

    except Exception:

        target_sheet = sh.add_worksheet(
            title=sheet_name,
            rows=1000,
            cols=20,
        )

        target_sheet.update(
            "A1:N1",
            [DEFAULT_SPK_HEADERS],
            value_input_option="USER_ENTERED",
        )

        clear_sheet_cache()

        return target_sheet


# ==============================================================================
# ENSURE STATUS COLUMN
# ==============================================================================

def ensure_status_column(
    sheet,
    df,
):
    """
    Pastikan Status Approval tersedia.
    """

    df = normalize_dataframe_headers(
        df.copy()
    )

    status_col = find_column(
        df,
        [
            "Status Approval",
            "Approval Status",
            "Status",
        ],
        fallback=None,
    )

    if status_col:
        return df, status_col

    status_col = "Status Approval"

    try:

        headers = [
            normalize_header(h)
            for h in sheet.row_values(1)
        ]

        if status_col not in headers:

            headers.append(status_col)

            header_end = column_letter(
                len(headers)
            )

            sheet.update(
                f"A1:{header_end}1",
                [headers],
                value_input_option="USER_ENTERED",
            )

            if len(df) > 0:

                status_col_index = len(headers)

                status_letter = column_letter(
                    status_col_index
                )

                status_values = [
                    [STATUS_PENDING]
                    for _ in range(len(df))
                ]

                sheet.update(
                    f"{status_letter}2:"
                    f"{status_letter}{len(df) + 1}",
                    status_values,
                    value_input_option="USER_ENTERED",
                )

            clear_sheet_cache()

    except Exception as e:

        st.warning(
            "⚠️ Kolom Status Approval belum "
            f"dapat ditambahkan ke Google Sheet: {e}"
        )

    df[status_col] = STATUS_PENDING

    return df, status_col


# ==============================================================================
# SPK SEQUENCE
# ==============================================================================

def get_next_spk_sequence_from_rows(rows):
    """
    Menghitung sequence berdasarkan No. SPK
    dari data yang SUDAH dibaca.

    Tidak melakukan request tambahan.
    """

    if not rows or len(rows) <= 1:
        return 1

    spk_ids = set()

    for row in rows[1:]:

        if len(row) > 1:

            value = safe_str(
                row[1]
            )

            if value:
                spk_ids.add(value)

    return len(spk_ids) + 1


# ==============================================================================
# LOAD MASTER DROPDOWN
# ==============================================================================

def load_master_dropdown():

    rows = load_sheet_values_cached(
        SHEET_DROPDOWN
    )

    if not rows or len(rows) <= 1:
        return pd.DataFrame()

    return dataframe_from_sheet_rows(
        rows
    )


# ==============================================================================
# LOAD MASTER SOW
# ==============================================================================

def load_master_sow():

    rows = load_sheet_values_cached(
        SHEET_SOW
    )

    if not rows or len(rows) <= 1:
        return pd.DataFrame()

    return dataframe_from_sheet_rows(
        rows
    )


# ==============================================================================
# MAIN PAGE
# ==============================================================================

def show_spk_page():

    st.markdown(
        PASTEL_ORANGE_CSS,
        unsafe_allow_html=True,
    )

    sh = get_cached_connection()

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
            "centang site yang diinginkan, lalu klik **Generate SPK**."
        )

        try:

            # ------------------------------------------------------------------
            # QUERY
            # ------------------------------------------------------------------

            query_rows = load_sheet_values_cached(
                SHEET_QUERY
            )

            # ------------------------------------------------------------------
            # MASTER DROPDOWN
            # ------------------------------------------------------------------

            df_dropdown = load_master_dropdown()

            # ------------------------------------------------------------------
            # MASTER SOW
            # ------------------------------------------------------------------

            df_master_sow = load_master_sow()

            # ------------------------------------------------------------------
            # VALIDASI QUERY
            # ------------------------------------------------------------------

            if (
                not query_rows
                or len(query_rows) <= 1
            ):

                st.warning(
                    "⚠️ Belum ada data pada sheet 'Query'."
                )

            else:

                df_query = dataframe_from_sheet_rows(
                    query_rows
                )

                # --------------------------------------------------------------
                # MASTER SOW DROPDOWN
                # --------------------------------------------------------------

                sow_dropdown_list = []

                master_sow_col = find_column(
                    df_dropdown,
                    ["Master SOW"],
                    fallback=None,
                )

                if (
                    master_sow_col
                    and not df_dropdown.empty
                ):

                    sow_dropdown_list = (
                        df_dropdown[
                            master_sow_col
                        ]
                        .dropna()
                        .unique()
                        .tolist()
                    )

                sow_dropdown_list = [
                    safe_str(s)
                    for s in sow_dropdown_list
                    if safe_str(s)
                ]

                if not sow_dropdown_list:

                    sow_dropdown_list = [
                        "Survey BSS",
                        "Instalasi BSS",
                        "Instalasi EVC",
                        MAINTENANCE_SOW_EVCS,
                        MAINTENANCE_SOW_BSS,
                    ]

                # Pastikan Maintenance selalu tersedia
                for maintenance_sow in MAINTENANCE_SOW_LIST:

                    if maintenance_sow not in sow_dropdown_list:
                        sow_dropdown_list.append(
                            maintenance_sow
                        )

                col1, col2, col3 = st.columns(3)

                # --------------------------------------------------------------
                # COLUMN 1
                # --------------------------------------------------------------

                with col1:

                    selected_sow_type = st.selectbox(
                        "Pilih Jenis SOW",
                        options=sow_dropdown_list,
                        key="sow_type_select",
                    )

                    sow_category = get_sow_category(
                        selected_sow_type
                    )

                    is_survey = (
                        sow_category == "SURVEY"
                    )

                    is_maintenance = (
                        sow_category
                        in [
                            "MAINTENANCE_EVCS",
                            "MAINTENANCE_BSS",
                        ]
                    )

                    # ----------------------------------------------------------
                    # PEKERJAAN
                    # ----------------------------------------------------------

                    if is_maintenance:

                        auto_pekerjaan = (
                            selected_sow_type
                        )

                    elif is_survey:

                        auto_pekerjaan = (
                            "Survey Location"
                        )

                    else:

                        auto_pekerjaan = (
                            "Construction"
                        )

                    pekerjaan_input = st.text_input(
                        "Pekerjaan",
                        value=auto_pekerjaan,
                        key=(
                            f"pekerjaan_input_"
                            f"{selected_sow_type}"
                        ),
                    )

                    # ----------------------------------------------------------
                    # WO
                    # ----------------------------------------------------------

                    target_wo_col_idx = (
                        get_wo_column_index(
                            selected_sow_type
                        )
                    )

                    # ==========================================================
                    # MAINTENANCE
                    # ==========================================================

                    if is_maintenance:

                        st.info(
                            "🔧 **Maintenance Service** dipilih. "
                            "Nomor WO bersifat opsional dan dapat dikosongkan."
                        )

                        selected_wo = ""

                        st.text_input(
                            "Pilih Nomor WO",
                            value="",
                            disabled=True,
                            placeholder="Tidak diperlukan untuk Maintenance",
                            key=(
                                f"maintenance_wo_display_"
                                f"{selected_sow_type}"
                            ),
                        )

                    # ==========================================================
                    # SURVEY / CONS
                    # ==========================================================

                    else:

                        if (
                            target_wo_col_idx is not None
                            and df_query.shape[1]
                            > target_wo_col_idx
                        ):

                            raw_wos = (
                                pd.Series(
                                    df_query.iloc[
                                        :,
                                        target_wo_col_idx,
                                    ].values.ravel()
                                )
                                .dropna()
                                .unique()
                            )

                            wo_list = [
                                safe_str(wo)
                                for wo in raw_wos
                                if safe_str(wo)
                                and safe_str(wo).lower()
                                != "nan"
                            ]

                        else:

                            wo_list = []

                        wo_label = (
                            f"Pilih Nomor WO "
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
                            key=(
                                f"wo_select_"
                                f"{selected_sow_type}"
                            ),
                        )

                    # ----------------------------------------------------------
                    # PROYEK
                    # ----------------------------------------------------------

                    proyek_input = st.text_input(
                        "Proyek",
                        value="V-Green",
                    )

                # --------------------------------------------------------------
                # COLUMN 2
                # --------------------------------------------------------------

                with col2:

                    if (
                        not df_dropdown.empty
                        and len(
                            df_dropdown.columns
                        ) > 0
                    ):

                        mitra_col_name = (
                            df_dropdown.columns[0]
                        )

                        mitra_list = (
                            df_dropdown[
                                mitra_col_name
                            ]
                            .dropna()
                            .unique()
                            .tolist()
                        )

                    else:

                        mitra_list = []

                    mitra_list = [
                        safe_str(m)
                        for m in mitra_list
                        if safe_str(m)
                    ]

                    mitra_list_clean = []

                    for m in mitra_list:

                        if (
                            safe_str(m).upper()
                            != "IN HOUSE"
                        ):

                            mitra_list_clean.append(
                                m
                            )

                    if "IN HOUSE" not in [
                        safe_str(x).upper()
                        for x in mitra_list_clean
                    ]:

                        mitra_list_clean.append(
                            "IN HOUSE"
                        )

                    selected_pic = st.selectbox(
                        "Penanggung Jawab (Mitra)",
                        options=(
                            mitra_list_clean
                            if mitra_list_clean
                            else [
                                "IN HOUSE"
                            ]
                        ),
                        key=(
                            f"selected_pic_"
                            f"{selected_sow_type}"
                        ),
                    )

                    manual_pic_name = ""

                    if (
                        safe_str(
                            selected_pic
                        ).upper()
                        == "IN HOUSE"
                    ):

                        st.info(
                            "🏢 **IN HOUSE** dipilih. "
                            "Silakan isi nama Penanggung Jawab secara manual."
                        )

                        manual_pic_name = st.text_input(
                            "Nama Penanggung Jawab (IN HOUSE)",
                            value="",
                            placeholder=(
                                "Masukkan nama Penanggung Jawab..."
                            ),
                            key=(
                                f"manual_pic_name_"
                                f"{selected_sow_type}"
                            ),
                        )

                    final_pic_name = resolve_pic_name(
                        selected_pic,
                        manual_pic_name,
                    )

                    default_phone = (
                        "0851-8259-6296"
                    )

                    if (
                        safe_str(
                            selected_pic
                        ).upper()
                        != "IN HOUSE"
                        and not df_dropdown.empty
                        and len(
                            df_dropdown.columns
                        ) > 1
                        and selected_pic
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
                                safe_str(
                                    selected_pic
                                )
                            ]
                        )

                        if not matched_row.empty:

                            default_phone = safe_str(
                                matched_row.iloc[0][
                                    phone_col_name
                                ],
                                default_phone,
                            )

                    pic_phone = st.text_input(
                        "No. Telepon Penanggung Jawab",
                        value=default_phone,
                        key=(
                            f"pic_phone_"
                            f"{selected_sow_type}"
                        ),
                    )

                # --------------------------------------------------------------
                # COLUMN 3
                # --------------------------------------------------------------

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

                    # ==========================================================
                    # MAINTENANCE INFORMATION
                    # ==========================================================

                    if is_maintenance:

                        st.info(
                            f"🔧 **{selected_sow_type}**\n\n"
                            "• No. WO: **Tidak wajib**\n"
                            "• Site Selection: **Tidak digunakan**\n"
                            "• SPK akan dibuat sebagai Maintenance Service."
                        )

                # ----------------------------------------------------------------
                # FILTER SITE
                # ----------------------------------------------------------------

                filtered_df = pd.DataFrame()
                edited_df = pd.DataFrame()
                selected_sites = pd.DataFrame()

                # ==============================================================
                # NORMAL SURVEY / CONS
                # ==============================================================

                if (
                    not is_maintenance
                    and selected_wo
                    and selected_wo
                    != "- Tidak ada WO -"
                ):

                    filtered_df = (
                        df_query[
                            df_query.iloc[
                                :,
                                target_wo_col_idx,
                            ]
                            .astype(str)
                            .str.strip()
                            == selected_wo
                        ]
                        .copy()
                    )

                    filtered_df["col_charge"] = (
                        filtered_df.iloc[:, 2]
                        if filtered_df.shape[1] > 2
                        else "-"
                    )

                    filtered_df["col_site"] = (
                        filtered_df.iloc[:, 5]
                        if filtered_df.shape[1] > 5
                        else "-"
                    )

                    filtered_df["col_gmaps"] = (
                        filtered_df.iloc[:, 7]
                        if filtered_df.shape[1] > 7
                        else "-"
                    )

                    filtered_df["col_province"] = (
                        filtered_df.iloc[:, 8]
                        if filtered_df.shape[1] > 8
                        else "-"
                    )

                    filtered_df["col_pic"] = (
                        filtered_df.iloc[:, 10]
                        if filtered_df.shape[1] > 10
                        else "-"
                    )

                    # ----------------------------------------------------------
                    # SOW INFORMATION
                    # ----------------------------------------------------------

                    filtered_df["col_sow"] = (
                        selected_sow_type
                    )

                    st.markdown("---")

                    st.markdown(
                        f"### 📍 Daftar Site untuk WO: "
                        f"`{selected_wo}` (Sheet Query)"
                    )

                    # ----------------------------------------------------------
                    # DISPLAY COLUMN SESUAI SOW
                    # ----------------------------------------------------------

                    display_columns = [
                        "Pilih",
                        "col_site",
                        "col_charge",
                        "col_province",
                        "col_pic",
                        "col_gmaps",
                        "col_sow",
                    ]

                    display_column_names = {
                        "Pilih": "Pilih",
                        "col_site": "Site Name",
                        "col_charge": "Charging Type",
                        "col_province": "Province",
                        "col_pic": "PIC + Contact",
                        "col_gmaps": "Gmaps",
                        "col_sow": "SOW",
                    }

                    site_editor_df = pd.DataFrame()

                    for source_col in display_columns:

                        if source_col == "Pilih":

                            site_editor_df[
                                "Pilih"
                            ] = True

                        elif source_col in filtered_df.columns:

                            site_editor_df[
                                display_column_names[
                                    source_col
                                ]
                            ] = filtered_df[
                                source_col
                            ].values

                    edited_display_df = st.data_editor(
                        site_editor_df,
                        use_container_width=True,
                        hide_index=True,
                        key=(
                            f"site_editor_"
                            f"{selected_wo}_"
                            f"{selected_sow_type}"
                        ),
                        column_config={
                            "Pilih": st.column_config.CheckboxColumn(
                                "Pilih",
                                default=True,
                            ),
                            "Site Name": st.column_config.TextColumn(
                                "Site Name",
                                disabled=True,
                            ),
                            "Charging Type": st.column_config.TextColumn(
                                "Charging Type",
                                disabled=True,
                            ),
                            "Province": st.column_config.TextColumn(
                                "Province",
                                disabled=True,
                            ),
                            "PIC + Contact": st.column_config.TextColumn(
                                "PIC + Contact",
                                disabled=True,
                            ),
                            "Gmaps": st.column_config.TextColumn(
                                "Gmaps",
                                disabled=True,
                            ),
                            "SOW": st.column_config.TextColumn(
                                "SOW",
                                disabled=True,
                            ),
                        },
                    )

                    # ----------------------------------------------------------
                    # MAP BACK TO ORIGINAL DATAFRAME
                    # ----------------------------------------------------------

                    selected_mask = (
                        edited_display_df[
                            "Pilih"
                        ]
                        == True
                    )

                    selected_row_indices = (
                        edited_display_df.index[
                            selected_mask
                        ]
                        .tolist()
                    )

                    if selected_row_indices:

                        selected_sites = (
                            filtered_df.iloc[
                                selected_row_indices
                            ]
                            .copy()
                        )

                    else:

                        selected_sites = (
                            pd.DataFrame()
                        )

                # ==============================================================
                # MAINTENANCE
                # ==============================================================

                elif is_maintenance:

                    # Maintenance tidak menggunakan site selection.
                    selected_sites = pd.DataFrame()

                    st.markdown("---")

                    st.success(
                        f"🔧 **{selected_sow_type}** dipilih.\n\n"
                        "Nomor WO dapat dikosongkan dan "
                        "tidak ada pemilihan site."
                    )

                # ----------------------------------------------------------------
                # GENERATE BUTTON
                # ----------------------------------------------------------------

                can_generate = False

                if is_maintenance:

                    can_generate = True

                elif (
                    selected_wo
                    and selected_wo
                    != "- Tidak ada WO -"
                    and not selected_sites.empty
                ):

                    can_generate = True

                if can_generate:

                    st.markdown("---")

                    # ----------------------------------------------------------
                    # GENERATE BUTTON
                    # ----------------------------------------------------------

                    if st.button(
                        "🚀 Generate SPK & Save to Database Sheet",
                        type="primary",
                    ):

                        # ======================================================
                        # VALIDATE PIC
                        # ======================================================

                        if (
                            safe_str(
                                selected_pic
                            ).upper()
                            == "IN HOUSE"
                        ):

                            if not safe_str(
                                manual_pic_name
                            ):

                                st.error(
                                    "❌ Karena Penanggung Jawab "
                                    "dipilih **IN HOUSE**, "
                                    "Nama Penanggung Jawab "
                                    "wajib diisi terlebih dahulu."
                                )

                                st.stop()

                            final_pic_name = safe_str(
                                manual_pic_name
                            )

                        else:

                            final_pic_name = safe_str(
                                selected_pic
                            )

                        # ======================================================
                        # VALIDATE PHONE
                        # ======================================================

                        if not safe_str(pic_phone):

                            st.error(
                                "❌ No. Telepon Penanggung Jawab "
                                "wajib diisi."
                            )

                            st.stop()

                        # ======================================================
                        # VALIDATE NORMAL SITE
                        # ======================================================

                        if (
                            not is_maintenance
                            and selected_sites.empty
                        ):

                            st.error(
                                "❌ Silakan centang minimal "
                                "satu site terlebih dahulu."
                            )

                            st.stop()

                        # ======================================================
                        # MAINTENANCE
                        # ======================================================

                        if is_maintenance:

                            # --------------------------------------------------
                            # WO HARUS KOSONG
                            # --------------------------------------------------

                            selected_wo_for_save = ""

                            # --------------------------------------------------
                            # LOCATION
                            # --------------------------------------------------

                            auto_lokasi = "-"

                            # --------------------------------------------------
                            # TARGET DATABASE
                            # --------------------------------------------------

                            target_db_name = SHEET_CONS

                        # ======================================================
                        # NORMAL SURVEY / CONS
                        # ======================================================

                        else:

                            selected_wo_for_save = (
                                safe_str(
                                    selected_wo
                                )
                            )

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
                                {
                                    p
                                    for p in provinces
                                    if p and p != "-"
                                }
                            )

                            auto_lokasi = (
                                ", ".join(
                                    unique_provinces
                                )
                                if unique_provinces
                                else "Jawa Tengah"
                            )

                            target_db_name = (
                                SHEET_SURVEY
                                if is_survey
                                else SHEET_CONS
                            )

                        with st.spinner(
                            "Memproses dokumen PDF SPK "
                            "& memperbarui Database..."
                        ):

                            # --------------------------------------------------
                            # MATCH SOW
                            # --------------------------------------------------

                            matched_sow_df = (
                                pd.DataFrame()
                            )

                            if (
                                not df_master_sow.empty
                            ):

                                kode_col = (
                                    df_master_sow.columns[
                                        0
                                    ]
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
                                        safe_str(
                                            selected_sow_type
                                        ).lower()
                                    ]
                                )

                            # --------------------------------------------------
                            # TARGET SHEET
                            # --------------------------------------------------

                            target_sheet = (
                                ensure_spk_sheet(
                                    sh,
                                    target_db_name,
                                )
                            )

                            # --------------------------------------------------
                            # READ EXISTING DATA
                            # --------------------------------------------------

                            existing_rows = (
                                load_sheet_values_cached(
                                    target_db_name
                                )
                            )

                            if (
                                not existing_rows
                                or len(existing_rows) <= 1
                            ):

                                existing_rows = [
                                    DEFAULT_SPK_HEADERS
                                ]

                                target_sheet.update(
                                    "A1:N1",
                                    [
                                        DEFAULT_SPK_HEADERS
                                    ],
                                    value_input_option=(
                                        "USER_ENTERED"
                                    ),
                                )

                                clear_sheet_cache(
                                    target_db_name
                                )

                            # --------------------------------------------------
                            # SEQUENCE
                            # --------------------------------------------------

                            seq_number = (
                                get_next_spk_sequence_from_rows(
                                    existing_rows
                                )
                            )

                            no_spk = (
                                generate_spk_number(
                                    selected_sow_type,
                                    sequence_num=seq_number,
                                )
                            )

                            # --------------------------------------------------
                            # METADATA
                            # --------------------------------------------------

                            spk_metadata = {
                                "no_spk": no_spk,
                                "proyek": proyek_input,
                                "pekerjaan": pekerjaan_input,
                                "lokasi": auto_lokasi,
                                "pic_name": final_pic_name,
                                "pic_phone": pic_phone,
                                "sow_type": selected_sow_type,
                            }

                            # --------------------------------------------------
                            # PDF
                            # --------------------------------------------------

                            (
                                no_spk,
                                pdf_bytes,
                            ) = (
                                generate_spk_pdf_bytes(
                                    selected_wo_for_save,
                                    selected_sites,
                                    spk_metadata,
                                    matched_sow_df,
                                )
                            )

                            safe_filename = (
                                f"SPK_"
                                f"{no_spk.replace('/', '_')}"
                                f".pdf"
                            )

                            now = datetime.datetime.now()

                            current_date_str = (
                                now.strftime(
                                    "%d/%m/%Y"
                                )
                            )

                            current_time_str = (
                                now.strftime(
                                    "%Y-%m-%d %H:%M:%S"
                                )
                            )

                            new_rows = []

                            # ==================================================
                            # MAINTENANCE ROW
                            # ==================================================

                            if is_maintenance:

                                new_rows.append(
                                    [
                                        current_date_str,
                                        no_spk,
                                        "",
                                        pekerjaan_input,
                                        "PT CLX",
                                        "",
                                        "",
                                        selected_sow_type,
                                        current_time_str,
                                        final_pic_name,
                                        "",
                                        "",
                                        "",
                                        STATUS_PENDING,
                                    ]
                                )

                            # ==================================================
                            # NORMAL SITE ROWS
                            # ==================================================

                            else:

                                for _, row in (
                                    selected_sites.iterrows()
                                ):

                                    site_name_val = safe_str(
                                        row.get(
                                            "col_site",
                                            "-",
                                        ),
                                        "-",
                                    )

                                    charger_type_val = safe_str(
                                        row.get(
                                            "col_charge",
                                            "-",
                                        ),
                                        "-",
                                    )

                                    new_rows.append(
                                        [
                                            current_date_str,
                                            no_spk,
                                            selected_wo_for_save,
                                            pekerjaan_input,
                                            "PT CLX",
                                            site_name_val,
                                            charger_type_val,
                                            selected_sow_type,
                                            current_time_str,
                                            final_pic_name,
                                            "",
                                            "",
                                            "",
                                            STATUS_PENDING,
                                        ]
                                    )

                            # --------------------------------------------------
                            # APPEND ROWS
                            # --------------------------------------------------

                            if new_rows:

                                target_sheet.append_rows(
                                    new_rows,
                                    value_input_option=(
                                        "USER_ENTERED"
                                    ),
                                )

                                clear_sheet_cache(
                                    target_db_name
                                )

                            # --------------------------------------------------
                            # SUCCESS
                            # --------------------------------------------------

                            if is_maintenance:

                                st.success(
                                    f"✅ SPK `{no_spk}` "
                                    f"untuk **{selected_sow_type}** "
                                    "berhasil dibuat & dikirim "
                                    "ke COO Dashboard untuk Approval!"
                                )

                                st.info(
                                    "🔧 **Maintenance Information**\n\n"
                                    "• No. WO: **kosong**\n"
                                    f"• SOW: **{selected_sow_type}**\n"
                                    f"• Pekerjaan: **{pekerjaan_input}**\n"
                                    f"• PIC: **{final_pic_name}**"
                                )

                            else:

                                st.success(
                                    f"✅ SPK `{no_spk}` "
                                    "berhasil dibuat & dikirim "
                                    "ke COO Dashboard untuk Approval!"
                                )

                            if (
                                safe_str(
                                    selected_pic
                                ).upper()
                                == "IN HOUSE"
                            ):

                                st.info(
                                    f"🏢 **IN HOUSE**\n\n"
                                    f"Nama Penanggung Jawab: "
                                    f"**{final_pic_name}**\n\n"
                                    f"Nilai yang disimpan ke "
                                    f"kolom **J (Mitra)**: "
                                    f"**{final_pic_name}**"
                                )

                            st.download_button(
                                label=(
                                    "📥 Download File PDF SPK (Draft)"
                                ),
                                data=pdf_bytes,
                                file_name=safe_filename,
                                mime="application/pdf",
                                type="primary",
                            )

        except Exception as e:

            st.error(
                "❌ Terjadi kesalahan pada Menu "
                f"Generate SPK: {e}"
            )

    # ==========================================================================
    # MENU 2 - TAKE OVER
    # ==========================================================================

    elif main_menu == "🔄 Take Over Site":

        st.subheader(
            "🔄 Menu Take Over Site"
        )

        st.markdown(
            "Pilih **Sheet Target** (Survey / Cons) "
            "dan **Site Name** yang akan di-*Take Over*. "
            "Sistem akan otomatis membuat No. SPK Baru "
            "dan mencatat riwayat *Take Over* pada kolom K:M."
        )

        try:

            db_target_type = st.radio(
                "Pilih Kategori Database SPK",
                options=[
                    SHEET_CONS,
                    SHEET_SURVEY,
                ],
                horizontal=True,
                key="takeover_db_choice",
            )

            target_sheet_to = sh.worksheet(
                db_target_type
            )

            # ------------------------------------------------------------------
            # SATU READ + CACHE
            # ------------------------------------------------------------------

            to_rows = load_sheet_values_cached(
                db_target_type
            )

            if len(to_rows) <= 1:

                st.warning(
                    f"⚠️ Sheet `{db_target_type}` "
                    "belum memiliki data."
                )

            else:

                df_to = dataframe_from_sheet_rows(
                    to_rows
                )

                # --------------------------------------------------------------
                # FIND COLUMNS
                # --------------------------------------------------------------

                col_site_name = find_column(
                    df_to,
                    [
                        "Site Name",
                        "Site",
                        "Nama Site",
                    ],
                    fallback="Site Name",
                )

                col_spk_num = find_column(
                    df_to,
                    [
                        "No. SPK",
                        "No SPK",
                        "SPK",
                    ],
                    fallback="No. SPK",
                )

                col_mitra_name = find_column(
                    df_to,
                    [
                        "Mitra",
                        "Mitra Name",
                    ],
                    fallback="Mitra",
                )

                if (
                    col_site_name not in df_to.columns
                ):

                    st.error(
                        "❌ Kolom Site Name tidak ditemukan."
                    )

                    st.stop()

                site_options = (
                    df_to[
                        col_site_name
                    ]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .unique()
                    .tolist()
                )

                site_options = [
                    s
                    for s in site_options
                    if s not in [
                        "",
                        "-",
                    ]
                ]

                col_to1, col_to2 = (
                    st.columns(2)
                )

                with col_to1:

                    selected_site_to = st.selectbox(
                        "Pilih Site Name yang akan di Take Over",
                        options=(
                            site_options
                            if site_options
                            else [
                                "- Tidak Ada Site -"
                            ]
                        ),
                        key="takeover_site_select",
                    )

                if (
                    selected_site_to
                    and selected_site_to
                    != "- Tidak Ada Site -"
                ):

                    matched_to_row = (
                        df_to[
                            df_to[
                                col_site_name
                            ]
                            .astype(str)
                            .str.strip()
                            ==
                            safe_str(
                                selected_site_to
                            )
                        ]
                    )

                    old_spk = (
                        safe_str(
                            matched_to_row.iloc[0][
                                col_spk_num
                            ],
                            "-",
                        )
                        if (
                            not matched_to_row.empty
                            and col_spk_num
                            in matched_to_row.columns
                        )
                        else "-"
                    )

                    old_mitra = (
                        safe_str(
                            matched_to_row.iloc[0][
                                col_mitra_name
                            ],
                            "-",
                        )
                        if (
                            not matched_to_row.empty
                            and col_mitra_name
                            in matched_to_row.columns
                        )
                        else "-"
                    )

                    with col_to2:

                        st.info(
                            f"📌 **No SPK Lama "
                            f"(di Take Over):** {old_spk}"
                        )

                        st.info(
                            f"👤 **Mitra Lama:** "
                            f"{old_mitra}"
                        )

                    st.markdown(
                        "### 📝 Informasi Take Over Baru"
                    )

                    col_to3, col_to4 = (
                        st.columns(2)
                    )

                    # ----------------------------------------------------------
                    # MASTER MITRA
                    # ----------------------------------------------------------

                    df_dropdown_to = (
                        load_master_dropdown()
                    )

                    new_mitra_list = []

                    if (
                        not df_dropdown_to.empty
                        and len(
                            df_dropdown_to.columns
                        ) > 0
                    ):

                        mitra_col = (
                            df_dropdown_to.columns[0]
                        )

                        new_mitra_list = (
                            df_dropdown_to[
                                mitra_col
                            ]
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .unique()
                            .tolist()
                        )

                        new_mitra_list = [
                            x
                            for x in new_mitra_list
                            if x
                        ]

                    with col_to3:

                        new_mitra = st.selectbox(
                            "Pilih Mitra Baru (Pengambil Alih)",
                            options=(
                                new_mitra_list
                                if new_mitra_list
                                else [
                                    "Mitra Baru"
                                ]
                            ),
                            key="takeover_mitra_select",
                        )

                        sow_type_to = (
                            "Survey BSS"
                            if "Survey"
                            in db_target_type
                            else "Instalasi BSS"
                        )

                    with col_to4:

                        today_date_str = (
                            datetime.datetime.now()
                            .strftime(
                                "%d/%m/%Y"
                            )
                        )

                        st.text_input(
                            "Tanggal Take Over (Otomatis)",
                            value=today_date_str,
                            disabled=True,
                            key="takeover_date_display",
                        )

                    # ----------------------------------------------------------
                    # PROCESS
                    # ----------------------------------------------------------

                    if st.button(
                        "🔥 Proses & Simpan Take Over Site",
                        type="primary",
                        key="process_takeover",
                    ):

                        with st.spinner(
                            "Memproses Take Over "
                            "& memperbarui sheet..."
                        ):

                            # --------------------------------------------------
                            # SEQUENCE
                            # --------------------------------------------------

                            seq_number = (
                                get_next_spk_sequence_from_rows(
                                    to_rows
                                )
                            )

                            new_spk_no = (
                                generate_spk_number(
                                    sow_type_to,
                                    sequence_num=seq_number,
                                )
                            )

                            # --------------------------------------------------
                            # FIND TARGET ROW
                            # --------------------------------------------------

                            site_col_index = (
                                list(
                                    df_to.columns
                                ).index(
                                    col_site_name
                                )
                                + 1
                            )

                            target_row_number = None

                            for row_index, row in enumerate(
                                to_rows[1:],
                                start=2,
                            ):

                                if (
                                    len(row)
                                    >= site_col_index
                                ):

                                    current_site = safe_str(
                                        row[
                                            site_col_index
                                            - 1
                                        ]
                                    )

                                    if (
                                        current_site
                                        ==
                                        safe_str(
                                            selected_site_to
                                        )
                                    ):

                                        target_row_number = (
                                            row_index
                                        )

                                        break

                            if (
                                target_row_number
                                is None
                            ):

                                st.error(
                                    "❌ Site tidak ditemukan "
                                    "pada Google Sheet."
                                )

                                st.stop()

                            # --------------------------------------------------
                            # BATCH UPDATE K:M
                            # --------------------------------------------------

                            takeover_updates = [
                                {
                                    "range": (
                                        f"K{target_row_number}:M"
                                        f"{target_row_number}"
                                    ),
                                    "values": [
                                        [
                                            today_date_str,
                                            old_spk,
                                            safe_str(
                                                new_mitra
                                            ),
                                        ]
                                    ],
                                }
                            ]

                            batch_update_cells(
                                target_sheet_to,
                                takeover_updates,
                            )

                            clear_sheet_cache(
                                db_target_type
                            )

                            st.success(
                                f"🎉 Take Over Berhasil!\n\n"
                                f"• **Site:** "
                                f"{selected_site_to}\n"
                                f"• **No. SPK Baru:** "
                                f"`{new_spk_no}`\n"
                                f"• **Nomor SPK Lama "
                                f"Ter-record:** "
                                f"`{old_spk}`\n"
                                f"• **Mitra Baru:** "
                                f"`{new_mitra}`"
                            )

        except Exception as e:

            st.error(
                "❌ Terjadi kesalahan pada Menu "
                f"Take Over: {e}"
            )

    # ==========================================================================
    # MENU 3 - COO APPROVAL DASHBOARD
    # ==========================================================================

    elif main_menu == "🔒 COO Approval Dashboard":

        st.subheader(
            "🔒 COO Approval Dashboard "
            "(Otorisasi SPK)"
        )

        st.markdown(
            "Halaman khusus Manajemen/COO untuk "
            "menyetujui (**Approve**) atau "
            "menolak (**Reject**) pengajuan SPK."
        )

        # ----------------------------------------------------------------------
        # LOGIN
        # ----------------------------------------------------------------------

        pin_input = st.text_input(
            "Masukkan PIN Khusus COO:",
            type="password",
            placeholder="Masukkan PIN...",
            key="coo_pin_input",
        )

        if pin_input == COO_PIN_SECRET:

            st.success(
                "🔓 Akses Diterima! "
                "Selamat datang, COO."
            )

            db_approval_type = st.radio(
                "Pilih Kategori Database SPK",
                options=[
                    SHEET_CONS,
                    SHEET_SURVEY,
                ],
                horizontal=True,
                key="approval_db_choice",
            )

            try:

                # ------------------------------------------------------------------
                # LOAD SHEET
                # ------------------------------------------------------------------

                sheet_appr = sh.worksheet(
                    db_approval_type
                )

                appr_rows = load_sheet_values_cached(
                    db_approval_type
                )

                if len(appr_rows) <= 1:

                    st.info(
                        f"ℹ️ Belum ada data pada sheet "
                        f"`{db_approval_type}`."
                    )

                else:

                    # ==============================================================
                    # DATAFRAME
                    # ==============================================================

                    df_appr = dataframe_from_sheet_rows(
                        appr_rows
                    )

                    # ==============================================================
                    # FIND COLUMNS
                    # ==============================================================

                    col_date_spk = find_column(
                        df_appr,
                        [
                            "Date SPK",
                            "Date",
                            "Tanggal SPK",
                            "Tanggal",
                        ],
                        fallback=None,
                    )

                    col_no_spk = find_column(
                        df_appr,
                        [
                            "No. SPK",
                            "No SPK",
                            "SPK Number",
                            "Nomor SPK",
                        ],
                        fallback=None,
                    )

                    col_no_wo = find_column(
                        df_appr,
                        [
                            "No. WO",
                            "No WO",
                            "WO Number",
                            "Nomor WO",
                        ],
                        fallback=None,
                    )

                    col_pekerjaan = find_column(
                        df_appr,
                        [
                            "Pekerjaan",
                            "Job",
                            "Work",
                        ],
                        fallback=None,
                    )

                    col_site = find_column(
                        df_appr,
                        [
                            "Site Name",
                            "Site",
                            "Nama Site",
                        ],
                        fallback=None,
                    )

                    col_mitra = find_column(
                        df_appr,
                        [
                            "Mitra",
                            "Mitra Name",
                            "Partner",
                        ],
                        fallback=None,
                    )

                    # ==============================================================
                    # STATUS
                    # ==============================================================

                    df_appr, col_status = (
                        ensure_status_column(
                            sheet_appr,
                            df_appr,
                        )
                    )

                    if col_status not in df_appr.columns:

                        df_appr[col_status] = (
                            STATUS_PENDING
                        )

                    # ==============================================================
                    # VALIDASI NO SPK
                    # ==============================================================

                    if not col_no_spk:

                        st.error(
                            "❌ Kolom **No. SPK** tidak ditemukan "
                            f"pada `{db_approval_type}`."
                        )

                        st.info(
                            "Header yang terbaca oleh sistem:"
                        )

                        st.code(
                            " | ".join(
                                df_appr.columns.tolist()
                            )
                        )

                        st.stop()

                    # ==============================================================
                    # CLEAN STATUS
                    # ==============================================================

                    df_appr[col_status] = (
                        df_appr[
                            col_status
                        ]
                        .fillna("")
                        .astype(str)
                        .str.strip()
                    )

                    # ==============================================================
                    # PENDING
                    # ==============================================================

                    pending_df = (
                        df_appr[
                            df_appr[
                                col_status
                            ]
                            .str.lower()
                            ==
                            STATUS_PENDING.lower()
                        ]
                        .copy()
                    )

                    st.markdown("---")

                    st.markdown(
                        f"### 📋 Daftar SPK Menunggu "
                        f"Persetujuan "
                        f"({len(pending_df)} Pending)"
                    )

                    # ==============================================================
                    # EMPTY
                    # ==============================================================

                    if pending_df.empty:

                        st.success(
                            "🎉 Semua pengajuan SPK "
                            "telah diproses! "
                            "Tidak ada antrean pending."
                        )

                    else:

                        # ==========================================================
                        # UNIQUE SPK
                        # ==========================================================

                        pending_spks = (
                            pending_df[
                                col_no_spk
                            ]
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .unique()
                            .tolist()
                        )

                        pending_spks = [
                            x
                            for x in pending_spks
                            if x
                        ]

                        col_app1, col_app2 = (
                            st.columns(
                                [2, 1]
                            )
                        )

                        with col_app1:

                            selected_appr_spk = st.selectbox(
                                "Pilih SPK yang akan ditinjau:",
                                options=pending_spks,
                                key=(
                                    f"approval_spk_select_"
                                    f"{db_approval_type}"
                                ),
                            )

                        # ==========================================================
                        # DETAIL
                        # ==========================================================

                        if selected_appr_spk:

                            spk_details = (
                                pending_df[
                                    pending_df[
                                        col_no_spk
                                    ]
                                    .astype(str)
                                    .str.strip()
                                    ==
                                    safe_str(
                                        selected_appr_spk
                                    )
                                ]
                            )

                            st.markdown(
                                "#### 🔍 Detail Informasi SPK:"
                            )

                            display_columns = []

                            for col in [
                                col_date_spk,
                                col_no_spk,
                                col_no_wo,
                                col_pekerjaan,
                                col_site,
                                col_mitra,
                            ]:

                                if (
                                    col
                                    and col
                                    in spk_details.columns
                                    and col
                                    not in display_columns
                                ):

                                    display_columns.append(
                                        col
                                    )

                            if not col_date_spk:

                                st.warning(
                                    "⚠️ Kolom `Date SPK` "
                                    "tidak ditemukan pada "
                                    f"`{db_approval_type}`. "
                                    "Data tetap dapat diproses."
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

                            # ======================================================
                            # BUTTONS
                            # ======================================================

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
                                        f"approve_"
                                        f"{db_approval_type}_"
                                        f"{selected_appr_spk}"
                                    ),
                                ):

                                    with st.spinner(
                                        "Memproses Approval..."
                                    ):

                                        # --------------------------------------------------
                                        # COLUMN INDEX
                                        # --------------------------------------------------

                                        status_col_index = (
                                            list(
                                                df_appr.columns
                                            ).index(
                                                col_status
                                            )
                                            + 1
                                        )

                                        status_letter = (
                                            column_letter(
                                                status_col_index
                                            )
                                        )

                                        # --------------------------------------------------
                                        # CARI ROW
                                        # --------------------------------------------------

                                        update_ranges = []

                                        updated_count = 0

                                        for df_index in (
                                            spk_details.index
                                        ):

                                            excel_row = (
                                                int(
                                                    df_index
                                                )
                                                + 2
                                            )

                                            update_ranges.append(
                                                {
                                                    "range": (
                                                        f"{status_letter}"
                                                        f"{excel_row}"
                                                    ),
                                                    "values": [
                                                        [
                                                            STATUS_APPROVED
                                                        ]
                                                    ],
                                                }
                                            )

                                            updated_count += 1

                                        # --------------------------------------------------
                                        # ONE BATCH REQUEST
                                        # --------------------------------------------------

                                        if update_ranges:

                                            batch_update_cells(
                                                sheet_appr,
                                                update_ranges,
                                            )

                                            clear_sheet_cache(
                                                db_approval_type
                                            )

                                        if updated_count > 0:

                                            st.success(
                                                f"✅ SPK "
                                                f"`{selected_appr_spk}` "
                                                "BERHASIL DI-APPROVE!"
                                            )

                                            st.rerun()

                                        else:

                                            st.error(
                                                "❌ SPK tidak ditemukan "
                                                "pada Google Sheet."
                                            )

                            # ------------------------------------------------------
                            # REJECT
                            # ------------------------------------------------------

                            with col_btn2:

                                if st.button(
                                    "❌ REJECT SPK",
                                    key=(
                                        f"reject_"
                                        f"{db_approval_type}_"
                                        f"{selected_appr_spk}"
                                    ),
                                ):

                                    with st.spinner(
                                        "Memproses Penolakan..."
                                    ):

                                        status_col_index = (
                                            list(
                                                df_appr.columns
                                            ).index(
                                                col_status
                                            )
                                            + 1
                                        )

                                        status_letter = (
                                            column_letter(
                                                status_col_index
                                            )
                                        )

                                        # --------------------------------------------------
                                        # BATCH UPDATE
                                        # --------------------------------------------------

                                        update_ranges = []

                                        updated_count = 0

                                        for df_index in (
                                            spk_details.index
                                        ):

                                            excel_row = (
                                                int(
                                                    df_index
                                                )
                                                + 2
                                            )

                                            update_ranges.append(
                                                {
                                                    "range": (
                                                        f"{status_letter}"
                                                        f"{excel_row}"
                                                    ),
                                                    "values": [
                                                        [
                                                            STATUS_REJECTED
                                                        ]
                                                    ],
                                                }
                                            )

                                            updated_count += 1

                                        if update_ranges:

                                            batch_update_cells(
                                                sheet_appr,
                                                update_ranges,
                                            )

                                            clear_sheet_cache(
                                                db_approval_type
                                            )

                                        if updated_count > 0:

                                            st.warning(
                                                f"❌ SPK "
                                                f"`{selected_appr_spk}` "
                                                "DITOLAK!"
                                            )

                                            st.rerun()

                                        else:

                                            st.error(
                                                "❌ SPK tidak ditemukan "
                                                "pada Google Sheet."
                                            )

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


# ==============================================================================
# RUN APPLICATION
# ==============================================================================

if __name__ == "__main__":
    show_spk_page()
