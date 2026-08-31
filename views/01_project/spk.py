```python
# ==============================================================================
# SPK.PY
# ERP - SURAT PERINTAH KERJA
#
# VERSION:
# SPK.PY V2.2 FINAL - ANTI 429
#
# MAJOR IMPROVEMENTS:
# 1. Cached Google Sheets READ
# 2. No worksheet.cell() inside loops
# 3. No worksheet.update_cell() inside loops
# 4. Batch update for COO APPROVE
# 5. Batch update for COO REJECT
# 6. Batch update for TAKE OVER
# 7. Batch update for Status Approval initialization
# 8. Cached Query / Master Dropdown / Master SOW
# 9. Cache invalidation after WRITE
# 10. DataFrame used as the source for row matching
# 11. Reduced Google Sheets API traffic
# 12. Existing business logic preserved
# ==============================================================================

import datetime
import io
import os
import string

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

# TTL cache.
# Tujuannya supaya perubahan widget / rerun Streamlit tidak selalu
# menyebabkan READ baru ke Google Sheets.
SHEET_CACHE_TTL = 120


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
# HELPER - GOOGLE SHEETS CACHE
# ==============================================================================

@st.cache_data(
    ttl=SHEET_CACHE_TTL,
    show_spinner=False,
)
def get_cached_sheet_values(sheet_name):
    """
    Membaca seluruh isi worksheet dan menyimpannya di Streamlit cache.

    PENTING:
    Fungsi ini adalah salah satu titik utama pencegahan 429.

    Selama TTL masih aktif:
        get_all_values() TIDAK dipanggil ulang.

    Parameter hanya sheet_name sehingga cache tidak bergantung
    pada object worksheet / connection.
    """

    sh = get_google_sheet_connection()

    if not sh:
        return []

    worksheet = sh.worksheet(sheet_name)

    return worksheet.get_all_values()


def clear_spk_sheet_cache():
    """
    Clear cache READ milik halaman SPK.

    Dipanggil setelah WRITE agar data baru dapat terbaca
    tanpa harus menunggu TTL.
    """

    try:
        get_cached_sheet_values.clear()
    except Exception:
        pass


# ==============================================================================
# HELPER - SAFE STRING / HEADER
# ==============================================================================

def normalize_header(value):
    """
    Membersihkan nama header.
    """

    if value is None:
        return ""

    text = str(value).replace("\n", " ")
    text = text.replace("\r", " ")
    text = text.strip()

    text = " ".join(text.split())

    return text


def normalize_dataframe_headers(df):
    """
    Normalisasi seluruh nama kolom DataFrame.
    """

    if df is None or df.empty:
        return df

    df = df.copy()

    df.columns = [
        normalize_header(col)
        for col in df.columns
    ]

    return df


def find_column(df, candidates, fallback=None):
    """
    Mencari nama kolom secara fleksibel.

    Exact match terlebih dahulu,
    kemudian partial match.
    """

    if df is None or df.empty:
        return fallback

    normalized_columns = {
        normalize_header(col).lower(): col
        for col in df.columns
    }

    # Exact match
    for candidate in candidates:

        candidate_norm = (
            normalize_header(candidate).lower()
        )

        if candidate_norm in normalized_columns:
            return normalized_columns[candidate_norm]

    # Partial match
    for candidate in candidates:

        candidate_norm = (
            normalize_header(candidate).lower()
        )

        for normalized_col, original_col in normalized_columns.items():

            if (
                candidate_norm in normalized_col
                or normalized_col in candidate_norm
            ):
                return original_col

    return fallback


# ==============================================================================
# HELPER - DATAFRAME FROM SHEET VALUES
# ==============================================================================

def dataframe_from_sheet_values(sheet_values):
    """
    Mengubah hasil get_all_values() menjadi DataFrame secara aman.
    """

    if not sheet_values:
        return pd.DataFrame()

    if len(sheet_values) <= 1:
        return pd.DataFrame()

    headers = [
        normalize_header(h)
        for h in sheet_values[0]
    ]

    max_cols = max(
        len(headers),
        max(
            [
                len(row)
                for row in sheet_values[1:]
            ],
            default=0,
        ),
    )

    while len(headers) < max_cols:
        headers.append(
            f"Column_{len(headers) + 1}"
        )

    normalized_data = []

    for row in sheet_values[1:]:

        row_copy = list(row)

        while len(row_copy) < max_cols:
            row_copy.append("")

        normalized_data.append(
            row_copy[:max_cols]
        )

    df = pd.DataFrame(
        normalized_data,
        columns=headers,
    )

    return normalize_dataframe_headers(df)


# ==============================================================================
# HELPER - A1 COLUMN
# ==============================================================================

def column_number_to_letter(column_number):
    """
    Convert nomor kolom 1-based menjadi huruf Excel/A1.

    Contoh:
        1  -> A
        2  -> B
        26 -> Z
        27 -> AA
        28 -> AB
    """

    if column_number < 1:
        raise ValueError(
            "column_number harus >= 1"
        )

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


def make_a1_cell(row_number, column_number):
    """
    Membuat alamat A1.

    Contoh:
        row=2, col=14 -> N2
    """

    return (
        f"{column_number_to_letter(column_number)}"
        f"{row_number}"
    )


# ==============================================================================
# HELPER - BATCH UPDATE
# ==============================================================================

def batch_update_cells(
    worksheet,
    updates,
):
    """
    Melakukan banyak update cell dalam SATU API request.

    updates:
        [
            {
                "row": 2,
                "col": 14,
                "value": "Approved by COO"
            },
            ...
        ]

    Tujuan utama:
        menggantikan banyak update_cell()
        dengan satu batch request.
    """

    if not updates:
        return 0

    batch_data = []

    for item in updates:

        row_number = int(
            item["row"]
        )

        column_number = int(
            item["col"]
        )

        value = item.get(
            "value",
            "",
        )

        a1_address = make_a1_cell(
            row_number,
            column_number,
        )

        batch_data.append(
            {
                "range": a1_address,
                "values": [
                    [value]
                ],
            }
        )

    # Satu WRITE request ke Google Sheets.
    worksheet.batch_update(
        batch_data
    )

    return len(batch_data)


# ==============================================================================
# HELPER - BATCH APPEND / UPDATE HEADER
# ==============================================================================

def ensure_sheet_header(
    worksheet,
    headers,
):
    """
    Memastikan header sheet tersedia.

    Hanya melakukan WRITE jika diperlukan.
    """

    headers = [
        normalize_header(h)
        for h in headers
    ]

    try:
        current_values = (
            worksheet.get_all_values()
        )
    except Exception:
        current_values = []

    if not current_values:

        worksheet.update(
            "1:1",
            [headers],
        )

        return headers

    current_headers = [
        normalize_header(h)
        for h in current_values[0]
    ]

    if current_headers != headers:

        worksheet.update(
            "1:1",
            [headers],
        )

        return headers

    return current_headers


# ==============================================================================
# HELPER - STATUS APPROVAL
# ==============================================================================

def ensure_status_column(
    sheet,
    df,
    raw_rows=None,
):
    """
    Memastikan kolom Status Approval tersedia.

    IMPORTANT:
    Tidak lagi menggunakan update_cell() satu per satu.

    Jika kolom belum ada:
        1x update header
        1x batch_update seluruh status lama
    """

    df = normalize_dataframe_headers(df)

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

        # --------------------------------------------------------------
        # Ambil header dari DataFrame yang SUDAH ada di memory.
        # Tidak melakukan row_values() lagi.
        # --------------------------------------------------------------

        headers = list(
            df.columns
        )

        # Jika DataFrame kosong tetapi raw_rows tersedia,
        # gunakan header dari raw_rows.
        if (
            not headers
            and raw_rows
            and len(raw_rows) > 0
        ):

            headers = [
                normalize_header(h)
                for h in raw_rows[0]
            ]

        # Tambahkan header baru.
        headers.append(
            status_col
        )

        # --------------------------------------------------------------
        # WRITE HEADER
        # --------------------------------------------------------------

        sheet.update(
            "1:1",
            [headers],
        )

        # --------------------------------------------------------------
        # BATCH STATUS DATA
        # --------------------------------------------------------------

        if raw_rows:

            data_row_count = max(
                0,
                len(raw_rows) - 1,
            )

        else:

            data_row_count = len(df)

        if data_row_count > 0:

            status_column_index = len(
                headers
            )

            status_updates = []

            for excel_row_idx in range(
                2,
                data_row_count + 2,
            ):

                status_updates.append(
                    {
                        "row": excel_row_idx,
                        "col": status_column_index,
                        "value": "Pending COO Approval",
                    }
                )

            batch_update_cells(
                sheet,
                status_updates,
            )

        df[status_col] = (
            "Pending COO Approval"
        )

    except Exception as e:

        st.warning(
            f"⚠️ Kolom `{status_col}` belum dapat "
            f"ditambahkan ke Google Sheet: {e}"
        )

        df[status_col] = (
            "Pending COO Approval"
        )

    return df, status_col


# ==============================================================================
# HELPER - PIC
# ==============================================================================

def resolve_pic_name(
    selected_pic,
    manual_pic_name="",
):
    """
    Menentukan nama Penanggung Jawab.
    """

    selected_pic = (
        str(selected_pic).strip()
        if selected_pic is not None
        else ""
    )

    manual_pic_name = (
        str(manual_pic_name).strip()
        if manual_pic_name is not None
        else ""
    )

    if selected_pic.upper() == "IN HOUSE":
        return manual_pic_name

    return selected_pic


# ==============================================================================
# HELPER - GENERATE NOMOR SPK
# ==============================================================================

def generate_spk_number(
    sow_type="GENERAL",
    sequence_num=1,
):
    """
    Format:
    0001/CLX/SPK/CONS/VII/2026
    """

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

    sow_code = (
        "SURVEY"
        if "survey"
        in str(sow_type).lower()
        else "CONS"
    )

    seq_str = (
        f"{sequence_num:04d}"
    )

    return (
        f"{seq_str}/CLX/SPK/"
        f"{sow_code}/{month_roman}/{now.year}"
    )


# ==============================================================================
# HELPER - DEFAULT SPK HEADERS
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
# GENERATE PDF SPK
# ==============================================================================

def generate_spk_pdf_bytes(
    selected_wo,
    selected_sites,
    spk_metadata,
    matched_sow_df,
):
    """
    Membuat file PDF SPK dalam bentuk bytes.
    """

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

    # --------------------------------------------------------------------------
    # STYLE
    # --------------------------------------------------------------------------

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
    # 1. KOP SURAT
    # --------------------------------------------------------------------------

    logo_path = "assets/logo.png"

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
    # 2. JUDUL
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
    # 3. METADATA
    # --------------------------------------------------------------------------

    pic_name = (
        spk_metadata.get(
            "pic_name",
            "-",
        )
        or "-"
    )

    pic_phone = (
        spk_metadata.get(
            "pic_phone",
            "-",
        )
        or "-"
    )

    meta_table_data = [
        [
            Paragraph(
                "Proyek:",
                meta_label,
            ),
            Paragraph(
                str(
                    spk_metadata.get(
                        "proyek",
                        "-",
                    )
                ),
                meta_val,
            ),
            Paragraph(
                "No. WO:",
                meta_label,
            ),
            Paragraph(
                str(selected_wo),
                meta_val,
            ),
        ],
        [
            Paragraph(
                "Pekerjaan:",
                meta_label,
            ),
            Paragraph(
                str(
                    spk_metadata.get(
                        "pekerjaan",
                        "-",
                    )
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
            Paragraph(
                "Lokasi:",
                meta_label,
            ),
            Paragraph(
                str(
                    spk_metadata.get(
                        "lokasi",
                        "-",
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
            Paragraph(
                "Detail Site:",
                meta_label,
            ),
            Paragraph(
                "Terlampir pada tabel di bawah",
                meta_val,
            ),
            Paragraph(
                "SOW:",
                meta_label,
            ),
            Paragraph(
                str(
                    spk_metadata.get(
                        "sow_type",
                        "-",
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
    # 4. SOW
    # --------------------------------------------------------------------------

    sow_headers = [
        Paragraph(
            "<b>No.</b>",
            cell_head,
        ),
        Paragraph(
            "<b>Uraian Pekerjaan (SoW)</b>",
            cell_head,
        ),
        Paragraph(
            "<b>Target Penyelesaian</b>",
            cell_head,
        ),
    ]

    sow_rows = [
        sow_headers
    ]

    if (
        matched_sow_df is not None
        and not matched_sow_df.empty
    ):

        for idx, (_, r) in enumerate(
            matched_sow_df.iterrows(),
            1,
        ):

            deskripsi = (
                str(r.iloc[2])
                if r.shape[0] > 2
                else "-"
            )

            target = (
                str(r.iloc[3])
                if r.shape[0] > 3
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
                Paragraph(
                    "1",
                    cell_body,
                ),
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
    # 5. DETAIL SITE
    # --------------------------------------------------------------------------

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
            cell_head,
        ),
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

    site_rows = [
        site_headers
    ]

    if (
        selected_sites is not None
        and not selected_sites.empty
    ):

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
                        str(
                            row.get(
                                "col_site",
                                "-",
                            )
                        ),
                        cell_body,
                    ),
                    Paragraph(
                        str(
                            row.get(
                                "col_charge",
                                "-",
                            )
                        ),
                        cell_body,
                    ),
                    Paragraph(
                        str(selected_wo),
                        cell_body,
                    ),
                    Paragraph(
                        str(
                            row.get(
                                "col_province",
                                "-",
                            )
                        ),
                        cell_body,
                    ),
                    Paragraph(
                        str(
                            row.get(
                                "col_pic",
                                "-",
                            )
                        ),
                        cell_body,
                    ),
                    Paragraph(
                        str(
                            row.get(
                                "col_gmaps",
                                "-",
                            )
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
    # 6. SIGNATURE
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
        KeepTogether(
            [t_sign]
        )
    )

    doc.build(elements)

    buffer.seek(0)

    return (
        no_spk,
        buffer.getvalue(),
    )


# ==============================================================================
# MAIN PAGE
# ==============================================================================

def show_spk_page():

    st.markdown(
        PASTEL_ORANGE_CSS,
        unsafe_allow_html=True,
    )

    # ==========================================================================
    # CONNECTION
    # ==========================================================================

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

            # ==================================================================
            # QUERY - CACHED
            # ==================================================================

            query_rows = (
                get_cached_sheet_values(
                    "Query"
                )
            )

            # ==================================================================
            # MASTER DROPDOWN - CACHED
            # ==================================================================

            try:

                dropdown_data = (
                    get_cached_sheet_values(
                        "Master Dropdown"
                    )
                )

                df_dropdown = (
                    dataframe_from_sheet_values(
                        dropdown_data
                    )
                )

                if not df_dropdown.empty:

                    df_dropdown = (
                        df_dropdown.loc[
                            :,
                            ~df_dropdown.columns.duplicated(),
                        ]
                        .copy()
                    )

                    df_dropdown = (
                        normalize_dataframe_headers(
                            df_dropdown
                        )
                    )

            except Exception:

                df_dropdown = pd.DataFrame()

            # ==================================================================
            # MASTER SOW - CACHED
            # ==================================================================

            try:

                sow_data = (
                    get_cached_sheet_values(
                        "Master SOW"
                    )
                )

                df_master_sow = (
                    dataframe_from_sheet_values(
                        sow_data
                    )
                )

                if not df_master_sow.empty:

                    df_master_sow = (
                        df_master_sow.loc[
                            :,
                            ~df_master_sow.columns.duplicated(),
                        ]
                        .copy()
                    )

                    df_master_sow = (
                        normalize_dataframe_headers(
                            df_master_sow
                        )
                    )

            except Exception:

                df_master_sow = pd.DataFrame()

            # ==================================================================
            # QUERY DATAFRAME
            # ==================================================================

            if (
                not query_rows
                or len(query_rows) <= 1
            ):

                st.warning(
                    "⚠️ Belum ada data pada sheet 'Query'."
                )

            else:

                df_query = (
                    dataframe_from_sheet_values(
                        query_rows
                    )
                )

                # ==============================================================
                # SOW DROPDOWN
                # ==============================================================

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
                    str(s).strip()
                    for s in sow_dropdown_list
                    if str(s).strip() != ""
                ]

                if not sow_dropdown_list:

                    sow_dropdown_list = [
                        "Survey BSS",
                        "Instalasi BSS",
                        "Instalasi EVC",
                    ]

                col1, col2, col3 = (
                    st.columns(3)
                )

                # ==============================================================
                # COLUMN 1
                # ==============================================================

                with col1:

                    selected_sow_type = (
                        st.selectbox(
                            "Pilih Jenis SOW",
                            options=sow_dropdown_list,
                            key="sow_type_select",
                        )
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
                            pd.Series(
                                df_query.iloc[
                                    :,
                                    target_wo_col_idx,
                                ]
                                .values.ravel()
                            )
                            .dropna()
                            .unique()
                        )

                        wo_list = [
                            str(wo).strip()
                            for wo in raw_wos
                            if (
                                str(wo).strip()
                                != ""
                                and str(wo).strip()
                                .lower()
                                != "nan"
                            )
                        ]

                    else:

                        wo_list = []

                    wo_label = (
                        f"Pilih Nomor WO "
                        f"({'Kolom L - Survey' if is_survey else 'Kolom W - Cons'})"
                    )

                    selected_wo = (
                        st.selectbox(
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
                    )

                    proyek_input = (
                        st.text_input(
                            "Proyek",
                            value="V-Green",
                        )
                    )

                    pekerjaan_input = (
                        st.text_input(
                            "Pekerjaan",
                            value=auto_pekerjaan,
                            key=(
                                f"pekerjaan_input_"
                                f"{selected_sow_type}"
                            ),
                        )
                    )

                # ==============================================================
                # COLUMN 2
                # ==============================================================

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
                        str(m).strip()
                        for m in mitra_list
                        if str(m).strip() != ""
                    ]

                    mitra_list_clean = []

                    for m in mitra_list:

                        if (
                            str(m).strip().upper()
                            != "IN HOUSE"
                        ):

                            mitra_list_clean.append(
                                m
                            )

                    mitra_list_clean.append(
                        "IN HOUSE"
                    )

                    selected_pic = (
                        st.selectbox(
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
                    )

                    manual_pic_name = ""

                    if (
                        str(
                            selected_pic
                        ).strip().upper()
                        == "IN HOUSE"
                    ):

                        st.info(
                            "🏢 **IN HOUSE** dipilih. "
                            "Silakan isi nama Penanggung Jawab secara manual."
                        )

                        manual_pic_name = (
                            st.text_input(
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
                        )

                    final_pic_name = resolve_pic_name(
                        selected_pic=selected_pic,
                        manual_pic_name=manual_pic_name,
                    )

                    default_phone = (
                        "0851-8259-6296"
                    )

                    if (
                        str(
                            selected_pic
                        ).strip().upper()
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
                                str(
                                    selected_pic
                                ).strip()
                            ]
                        )

                        if not matched_row.empty:

                            default_phone = str(
                                matched_row.iloc[0][
                                    phone_col_name
                                ]
                            ).strip()

                    pic_phone = (
                        st.text_input(
                            "No. Telepon Penanggung Jawab",
                            value=default_phone,
                            key=(
                                f"pic_phone_"
                                f"{selected_sow_type}"
                            ),
                        )
                    )

                # ==============================================================
                # COLUMN 3
                # ==============================================================

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

                # ==============================================================
                # FILTER SITE
                # ==============================================================

                if (
                    selected_wo
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

                    filtered_df[
                        "col_charge"
                    ] = (
                        filtered_df.iloc[:, 2]
                        if filtered_df.shape[1]
                        > 2
                        else "-"
                    )

                    filtered_df[
                        "col_site"
                    ] = (
                        filtered_df.iloc[:, 5]
                        if filtered_df.shape[1]
                        > 5
                        else "-"
                    )

                    filtered_df[
                        "col_gmaps"
                    ] = (
                        filtered_df.iloc[:, 7]
                        if filtered_df.shape[1]
                        > 7
                        else "-"
                    )

                    filtered_df[
                        "col_province"
                    ] = (
                        filtered_df.iloc[:, 8]
                        if filtered_df.shape[1]
                        > 8
                        else "-"
                    )

                    filtered_df[
                        "col_pic"
                    ] = (
                        filtered_df.iloc[:, 10]
                        if filtered_df.shape[1]
                        > 10
                        else "-"
                    )

                    st.markdown("---")

                    st.markdown(
                        f"### 📍 Daftar Site untuk WO: "
                        f"`{selected_wo}` "
                        f"(Sheet Query)"
                    )

                    if (
                        "Pilih"
                        not in filtered_df.columns
                    ):

                        filtered_df.insert(
                            0,
                            "Pilih",
                            True,
                        )

                    edited_df = (
                        st.data_editor(
                            filtered_df,
                            use_container_width=True,
                            hide_index=True,
                            key=(
                                f"site_editor_"
                                f"{selected_wo}"
                            ),
                        )
                    )

                    st.markdown("---")

                    # ==========================================================
                    # GENERATE SPK
                    # ==========================================================

                    if st.button(
                        "🚀 Generate SPK & Save to Database Sheet",
                        type="primary",
                    ):

                        # ------------------------------------------------------
                        # VALIDASI PIC
                        # ------------------------------------------------------

                        if (
                            str(
                                selected_pic
                            ).strip().upper()
                            == "IN HOUSE"
                        ):

                            if (
                                not manual_pic_name
                                or not str(
                                    manual_pic_name
                                ).strip()
                            ):

                                st.error(
                                    "❌ Karena Penanggung Jawab "
                                    "dipilih **IN HOUSE**, "
                                    "Nama Penanggung Jawab "
                                    "wajib diisi terlebih dahulu."
                                )

                                st.stop()

                            final_pic_name = (
                                str(
                                    manual_pic_name
                                ).strip()
                            )

                        else:

                            final_pic_name = (
                                str(
                                    selected_pic
                                ).strip()
                            )

                        # ------------------------------------------------------
                        # VALIDASI PHONE
                        # ------------------------------------------------------

                        if (
                            not pic_phone
                            or not str(
                                pic_phone
                            ).strip()
                        ):

                            st.error(
                                "❌ No. Telepon Penanggung Jawab "
                                "wajib diisi."
                            )

                            st.stop()

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
                                "❌ Silakan centang minimal "
                                "satu site terlebih dahulu."
                            )

                        else:

                            with st.spinner(
                                "Memproses dokumen PDF SPK "
                                "& memperbarui Database..."
                            ):

                                # ==================================================
                                # MATCH MASTER SOW
                                # ==================================================

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
                                            str(
                                                selected_sow_type
                                            )
                                            .strip()
                                            .lower()
                                        ]
                                    )

                                # ==================================================
                                # LOKASI
                                # ==================================================

                                provinces = (
                                    selected_sites[
                                        "col_province"
                                    ]
                                    .dropna()
                                    .astype(str)
                                    .str.strip()
                                    .tolist()
                                )

                                unique_provinces = (
                                    sorted(
                                        list(
                                            set(
                                                [
                                                    p
                                                    for p in provinces
                                                    if (
                                                        p
                                                        != ""
                                                        and p
                                                        != "-"
                                                    )
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

                                # ==================================================
                                # TARGET DB
                                # ==================================================

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

                                    target_sheet.update(
                                        "1:1",
                                        [
                                            DEFAULT_SPK_HEADERS
                                        ],
                                    )

                                # ==================================================
                                # READ EXISTING DATA - CACHED
                                #
                                # Hanya satu read untuk sequence.
                                # ==================================================

                                existing_rows = (
                                    get_cached_sheet_values(
                                        target_db_name
                                    )
                                )

                                if not existing_rows:

                                    target_sheet.update(
                                        "1:1",
                                        [
                                            DEFAULT_SPK_HEADERS
                                        ],
                                    )

                                    existing_rows = [
                                        DEFAULT_SPK_HEADERS
                                    ]

                                # ==================================================
                                # GENERATE SEQUENCE
                                # ==================================================

                                spk_ids = []

                                if (
                                    existing_rows
                                    and len(existing_rows) > 1
                                ):

                                    for row in (
                                        existing_rows[1:]
                                    ):

                                        if (
                                            len(row) > 1
                                            and str(
                                                row[1]
                                            ).strip()
                                        ):

                                            spk_ids.append(
                                                str(
                                                    row[1]
                                                ).strip()
                                            )

                                unique_spk_count = len(
                                    set(spk_ids)
                                )

                                seq_number = (
                                    unique_spk_count
                                    + 1
                                )

                                no_spk = (
                                    generate_spk_number(
                                        selected_sow_type,
                                        sequence_num=seq_number,
                                    )
                                )

                                # ==================================================
                                # METADATA
                                # ==================================================

                                spk_metadata = {
                                    "no_spk": no_spk,
                                    "proyek": proyek_input,
                                    "pekerjaan": pekerjaan_input,
                                    "lokasi": auto_lokasi,
                                    "pic_name": final_pic_name,
                                    "pic_phone": pic_phone,
                                    "sow_type": selected_sow_type,
                                }

                                # ==================================================
                                # PDF
                                # ==================================================

                                (
                                    no_spk,
                                    pdf_bytes,
                                ) = (
                                    generate_spk_pdf_bytes(
                                        selected_wo,
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

                                current_date_str = (
                                    datetime.datetime.now()
                                    .strftime(
                                        "%d/%m/%Y"
                                    )
                                )

                                current_time_str = (
                                    datetime.datetime.now()
                                    .strftime(
                                        "%Y-%m-%d %H:%M:%S"
                                    )
                                )

                                new_rows = []

                                # ==================================================
                                # BUILD ROWS
                                # ==================================================

                                for _, row in (
                                    selected_sites.iterrows()
                                ):

                                    site_name_val = (
                                        row.get(
                                            "col_site",
                                            "-",
                                        )
                                    )

                                    charger_type_val = (
                                        row.get(
                                            "col_charge",
                                            "-",
                                        )
                                    )

                                    mitra_value_for_db = (
                                        final_pic_name
                                    )

                                    new_rows.append(
                                        [
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
                                            mitra_value_for_db,
                                            "",
                                            "",
                                            "",
                                            "Pending COO Approval",
                                        ]
                                    )

                                # ==================================================
                                # SAVE
                                # ==================================================

                                if new_rows:

                                    target_sheet.append_rows(
                                        new_rows
                                    )

                                    # Setelah WRITE:
                                    # cache sheet harus dihapus.
                                    clear_spk_sheet_cache()

                                st.success(
                                    f"✅ SPK `{no_spk}` "
                                    "berhasil dibuat & dikirim "
                                    "ke COO Dashboard untuk Approval!"
                                )

                                if (
                                    str(
                                        selected_pic
                                    ).strip().upper()
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
    # MENU 2 - TAKE OVER SITE
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
                    "DB SPK Cons",
                    "DB SPK Survey",
                ],
                horizontal=True,
                key="takeover_db_choice",
            )

            # ==================================================================
            # TAKE OVER DATABASE - CACHED
            # ==================================================================

            to_rows = (
                get_cached_sheet_values(
                    db_target_type
                )
            )

            try:

                target_sheet_to = (
                    sh.worksheet(
                        db_target_type
                    )
                )

            except Exception:

                st.error(
                    f"❌ Sheet `{db_target_type}` tidak ditemukan."
                )

                st.stop()

            if len(to_rows) <= 1:

                st.warning(
                    f"⚠️ Sheet `{db_target_type}` "
                    "belum memiliki data."
                )

            else:

                df_to = (
                    dataframe_from_sheet_values(
                        to_rows
                    )
                )

                # ==============================================================
                # FIND COLUMNS
                # ==============================================================

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

                site_options = (
                    df_to[
                        col_site_name
                    ]
                    .dropna()
                    .unique()
                    .tolist()
                )

                site_options = [
                    s
                    for s in site_options
                    if str(s).strip()
                    not in [
                        "",
                        "-",
                    ]
                ]

                col_to1, col_to2 = (
                    st.columns(2)
                )

                with col_to1:

                    selected_site_to = (
                        st.selectbox(
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
                            == str(
                                selected_site_to
                            )
                        ]
                    )

                    old_spk = (
                        matched_to_row.iloc[0][
                            col_spk_num
                        ]
                        if (
                            not matched_to_row.empty
                            and col_spk_num
                            in matched_to_row.columns
                        )
                        else "-"
                    )

                    old_mitra = (
                        matched_to_row.iloc[0][
                            col_mitra_name
                        ]
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

                    # ==========================================================
                    # MASTER MITRA - CACHED
                    # ==========================================================

                    try:

                        dropdown_data = (
                            get_cached_sheet_values(
                                "Master Dropdown"
                            )
                        )

                        df_dropdown = (
                            dataframe_from_sheet_values(
                                dropdown_data
                            )
                        )

                        if not df_dropdown.empty:

                            df_dropdown = (
                                normalize_dataframe_headers(
                                    df_dropdown
                                )
                            )

                        if (
                            not df_dropdown.empty
                            and len(
                                df_dropdown.columns
                            ) > 0
                        ):

                            mitra_col = (
                                df_dropdown.columns[0]
                            )

                            new_mitra_list = (
                                df_dropdown[
                                    mitra_col
                                ]
                                .dropna()
                                .unique()
                                .tolist()
                            )

                        else:

                            new_mitra_list = []

                    except Exception:

                        new_mitra_list = []

                    with col_to3:

                        new_mitra = (
                            st.selectbox(
                                "Pilih Mitra Baru (Pengambil Alih)",
                                options=(
                                    new_mitra_list
                                    if new_mitra_list
                                    else [
                                        "Mitra Baru"
                                    ]
                                ),
                                key="takeover_new_mitra",
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
                            datetime.datetime.now()
                            .strftime(
                                "%d/%m/%Y"
                            )
                        )

                        st.text_input(
                            "Tanggal Take Over (Otomatis)",
                            value=today_date_str,
                            disabled=True,
                            key="takeover_date",
                        )

                    # ==========================================================
                    # TAKE OVER BUTTON
                    # ==========================================================

                    if st.button(
                        "🔥 Proses & Simpan Take Over Site",
                        type="primary",
                        key="takeover_process_button",
                    ):

                        with st.spinner(
                            "Memproses Take Over "
                            "& memperbarui sheet..."
                        ):

                            # ==================================================
                            # GENERATE SEQUENCE
                            #
                            # Menggunakan data yang SUDAH di-read.
                            # Tidak melakukan read ulang.
                            # ==================================================

                            spk_ids = []

                            if len(to_rows) > 1:

                                for row in (
                                    to_rows[1:]
                                ):

                                    if (
                                        len(row) > 1
                                        and str(
                                            row[1]
                                        ).strip()
                                    ):

                                        spk_ids.append(
                                            str(
                                                row[1]
                                            ).strip()
                                        )

                            seq_number = (
                                len(
                                    set(
                                        spk_ids
                                    )
                                )
                                + 1
                            )

                            new_spk_no = (
                                generate_spk_number(
                                    sow_type_to,
                                    sequence_num=seq_number,
                                )
                            )

                            # ==================================================
                            # TENTUKAN ROW YANG AKAN DIUPDATE
                            #
                            # DataFrame index 0 = Excel row 2.
                            # ==================================================

                            matching_indexes = (
                                df_to.index[
                                    df_to[
                                        col_site_name
                                    ]
                                    .astype(str)
                                    .str.strip()
                                    ==
                                    str(
                                        selected_site_to
                                    ).strip()
                                ]
                                .tolist()
                            )

                            if not matching_indexes:

                                st.error(
                                    "❌ Site tidak ditemukan "
                                    "pada data database."
                                )

                                st.stop()

                            # Saat ini logic lama hanya mengubah
                            # baris pertama yang cocok.
                            target_df_index = (
                                matching_indexes[0]
                            )

                            excel_row_idx = (
                                target_df_index
                                + 2
                            )

                            # ==================================================
                            # K = Date Take Over
                            # L = No SPK Lama
                            # M = Mitra Baru
                            #
                            # SATU BATCH WRITE.
                            # ==================================================

                            takeover_updates = [
                                {
                                    "row": excel_row_idx,
                                    "col": 11,
                                    "value": today_date_str,
                                },
                                {
                                    "row": excel_row_idx,
                                    "col": 12,
                                    "value": old_spk,
                                },
                                {
                                    "row": excel_row_idx,
                                    "col": 13,
                                    "value": new_mitra,
                                },
                            ]

                            batch_update_cells(
                                target_sheet_to,
                                takeover_updates,
                            )

                            clear_spk_sheet_cache()

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
        # LOGIN COO
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
                    "DB SPK Cons",
                    "DB SPK Survey",
                ],
                horizontal=True,
                key="approval_db_choice",
            )

            try:

                # ==============================================================
                # LOAD SHEET OBJECT
                # ==============================================================

                try:

                    sheet_appr = sh.worksheet(
                        db_approval_type
                    )

                except Exception:

                    st.error(
                        f"❌ Sheet `{db_approval_type}` "
                        "tidak ditemukan."
                    )

                    st.stop()

                # ==============================================================
                # SINGLE CACHED READ
                #
                # INI SANGAT PENTING:
                #
                # Sebelum:
                #     get_all_values()
                #     cell()
                #     cell()
                #     cell()
                #     ...
                #
                # Sekarang:
                #     1x cached get_all_values()
                #
                # Semua pencarian row dilakukan dari DataFrame.
                # ==============================================================

                appr_rows = (
                    get_cached_sheet_values(
                        db_approval_type
                    )
                )

                if len(appr_rows) <= 1:

                    st.info(
                        f"ℹ️ Belum ada data pada sheet "
                        f"`{db_approval_type}`."
                    )

                else:

                    # ==========================================================
                    # CREATE DATAFRAME
                    # ==========================================================

                    df_appr = (
                        dataframe_from_sheet_values(
                            appr_rows
                        )
                    )

                    # ==========================================================
                    # FIND COLUMNS
                    # ==========================================================

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

                    # ==========================================================
                    # PASTIKAN STATUS APPROVAL
                    # ==========================================================

                    df_appr, col_status = (
                        ensure_status_column(
                            sheet_appr,
                            df_appr,
                            raw_rows=appr_rows,
                        )
                    )

                    # Jika ensure_status_column menulis ke sheet,
                    # cache lama harus dibersihkan.
                    if col_status not in (
                        dataframe_from_sheet_values(
                            appr_rows
                        ).columns
                        if appr_rows
                        else []
                    ):

                        clear_spk_sheet_cache()

                    # ==========================================================
                    # VALIDASI NO SPK
                    # ==========================================================

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

                    # ==========================================================
                    # CLEAN STATUS
                    # ==========================================================

                    df_appr[col_status] = (
                        df_appr[col_status]
                        .fillna("")
                        .astype(str)
                        .str.strip()
                    )

                    # ==========================================================
                    # FILTER PENDING
                    # ==============================================================

                    pending_df = (
                        df_appr[
                            df_appr[
                                col_status
                            ]
                            .str.lower()
                            ==
                            "pending coo approval"
                        ]
                        .copy()
                    )

                    st.markdown("---")

                    st.markdown(
                        f"### 📋 Daftar SPK Menunggu "
                        f"Persetujuan "
                        f"({len(pending_df)} Pending)"
                    )

                    # ==========================================================
                    # EMPTY
                    # ==========================================================

                    if pending_df.empty:

                        st.balloons()

                        st.success(
                            "🎉 Semua pengajuan SPK "
                            "telah diproses! "
                            "Tidak ada antrean pending."
                        )

                    else:

                        # ======================================================
                        # UNIQUE SPK
                        # ======================================================

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
                            if x != ""
                        ]

                        col_app1, col_app2 = (
                            st.columns(
                                [2, 1]
                            )
                        )

                        with col_app1:

                            selected_appr_spk = (
                                st.selectbox(
                                    "Pilih SPK yang akan ditinjau:",
                                    options=pending_spks,
                                    key=(
                                        f"approval_spk_select_"
                                        f"{db_approval_type}"
                                    ),
                                )
                            )

                        # ======================================================
                        # DETAIL
                        # ======================================================

                        if selected_appr_spk:

                            spk_details = (
                                pending_df[
                                    pending_df[
                                        col_no_spk
                                    ]
                                    .astype(str)
                                    .str.strip()
                                    ==
                                    str(
                                        selected_appr_spk
                                    ).strip()
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

                            # ==================================================
                            # BUTTON
                            # ==================================================

                            col_btn1, col_btn2, _ = (
                                st.columns(
                                    [1, 1, 2]
                                )
                            )

                            # ==================================================
                            # APPROVE
                            # ==================================================

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

                                        # --------------------------------------
                                        # Tentukan kolom
                                        # --------------------------------------

                                        status_col_index = (
                                            list(
                                                df_appr.columns
                                            ).index(
                                                col_status
                                            )
                                            + 1
                                        )

                                        spk_col_index = (
                                            list(
                                                df_appr.columns
                                            ).index(
                                                col_no_spk
                                            )
                                            + 1
                                        )

                                        # --------------------------------------
                                        # CARI BARIS DARI DATAFRAME
                                        #
                                        # TIDAK ADA:
                                        #     sheet_appr.cell()
                                        #
                                        # --------------------------------------

                                        matching_indices = (
                                            df_appr.index[
                                                df_appr[
                                                    col_no_spk
                                                ]
                                                .astype(str)
                                                .str.strip()
                                                ==
                                                str(
                                                    selected_appr_spk
                                                ).strip()
                                            ]
                                            .tolist()
                                        )

                                        if not matching_indices:

                                            st.error(
                                                "❌ SPK tidak ditemukan "
                                                "pada data yang sudah dibaca."
                                            )

                                        else:

                                            approval_updates = []

                                            for df_index in (
                                                matching_indices
                                            ):

                                                # DataFrame row 0
                                                # = Excel row 2.
                                                excel_row_idx = (
                                                    int(
                                                        df_index
                                                    )
                                                    + 2
                                                )

                                                approval_updates.append(
                                                    {
                                                        "row": excel_row_idx,
                                                        "col": status_col_index,
                                                        "value": "Approved by COO",
                                                    }
                                                )

                                            # ----------------------------------
                                            # SATU BATCH WRITE
                                            # ----------------------------------

                                            updated_count = (
                                                batch_update_cells(
                                                    sheet_appr,
                                                    approval_updates,
                                                )
                                            )

                                            # ----------------------------------
                                            # INVALIDATE CACHE
                                            # ----------------------------------

                                            clear_spk_sheet_cache()

                                            if updated_count > 0:

                                                st.success(
                                                    f"✅ SPK "
                                                    f"`{selected_appr_spk}` "
                                                    "BERHASIL DI-APPROVE!"
                                                )

                                                st.rerun()

                                            else:

                                                st.error(
                                                    "❌ Tidak ada baris "
                                                    "yang berhasil di-update."
                                                )

                            # ==================================================
                            # REJECT
                            # ==================================================

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

                                        # --------------------------------------
                                        # Tentukan kolom
                                        # --------------------------------------

                                        status_col_index = (
                                            list(
                                                df_appr.columns
                                            ).index(
                                                col_status
                                            )
                                            + 1
                                        )

                                        # --------------------------------------
                                        # CARI ROW DARI DATAFRAME
                                        #
                                        # TIDAK ADA:
                                        #     sheet_appr.cell()
                                        # --------------------------------------

                                        matching_indices = (
                                            df_appr.index[
                                                df_appr[
                                                    col_no_spk
                                                ]
                                                .astype(str)
                                                .str.strip()
                                                ==
                                                str(
                                                    selected_appr_spk
                                                ).strip()
                                            ]
                                            .tolist()
                                        )

                                        if not matching_indices:

                                            st.error(
                                                "❌ SPK tidak ditemukan "
                                                "pada data yang sudah dibaca."
                                            )

                                        else:

                                            reject_updates = []

                                            for df_index in (
                                                matching_indices
                                            ):

                                                excel_row_idx = (
                                                    int(
                                                        df_index
                                                    )
                                                    + 2
                                                )

                                                reject_updates.append(
                                                    {
                                                        "row": excel_row_idx,
                                                        "col": status_col_index,
                                                        "value": "Rejected by COO",
                                                    }
                                                )

                                            # ----------------------------------
                                            # SATU BATCH WRITE
                                            # ----------------------------------

                                            updated_count = (
                                                batch_update_cells(
                                                    sheet_appr,
                                                    reject_updates,
                                                )
                                            )

                                            # ----------------------------------
                                            # INVALIDATE CACHE
                                            # ----------------------------------

                                            clear_spk_sheet_cache()

                                            if updated_count > 0:

                                                st.warning(
                                                    f"❌ SPK "
                                                    f"`{selected_appr_spk}` "
                                                    "DITOLAK!"
                                                )

                                                st.rerun()

                                            else:

                                                st.error(
                                                    "❌ Tidak ada baris "
                                                    "yang berhasil di-update."
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
```
