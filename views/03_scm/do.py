import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os


# ==============================================================================
# REPORTLAB
# ==============================================================================
try:
    from reportlab.lib.pagesizes import A5, portrait
    from reportlab.lib import colors
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
        Image as RLImage
    )
    from reportlab.lib.styles import (
        getSampleStyleSheet,
        ParagraphStyle
    )

    REPORTLAB_AVAILABLE = True

except ImportError:
    REPORTLAB_AVAILABLE = False


# ==============================================================================
# IMPORT CORE DATABASE
# ==============================================================================
try:

    from core.database import (
        generate_do_number,
        save_do_to_db_material_out,
        get_all_do_numbers,
        get_do_by_number,
        update_do_in_db_material_out,
        get_used_sites_from_db_material_out,
        get_query_sheet_data,
        get_sheet_values
    )

except ImportError as e:

    st.error(
        "❌ Modul `core.database` tidak dapat di-load.\n\n"
        f"Detail error: {e}"
    )

    st.stop()


# ==============================================================================
# CONSTANT
# ==============================================================================

MAX_SITE_SELECTION = 15

# ------------------------------------------------------------------------------
# DEFAULT CHARGING TYPE
#
# Tetap dipertahankan sebagai fallback agar logic lama tidak rusak.
# Charging Type yang ada di sheet "Standart Charging Type" akan ditambahkan
# secara otomatis ke dropdown.
# ------------------------------------------------------------------------------
DEFAULT_CHARGING_TYPES = [
    "6S1P",
    "12S1P",
    "DC20",
    "DC30",
    "DC60",
    "DC120"
]

DEFAULT_EXPEDITIONS = [
    "BCE",
    "Lalamove",
    "Self Pick Up",
    "JNE",
    "TIKI"
]

# ------------------------------------------------------------------------------
# NAMA SHEET
# ------------------------------------------------------------------------------
MASTER_ITEM_SHEET = "Master Item"

STANDARD_CHARGING_TYPE_SHEET = "Standart Charging Type"


# ==============================================================================
# HELPER SESSION STATE
# ==============================================================================

def initialize_session_state():

    defaults = {
        "relocation_history": [],
        "current_do": None,
        "edit_do_data": None,
        "do_number_draft": None,
        "do_number_generated_at": None,
        "do_success_notification": None
    }

    for key, value in defaults.items():

        if key not in st.session_state:

            st.session_state[key] = value


# ==============================================================================
# SAFE VALUE
# ==============================================================================

def safe_text(value, default=""):

    if value is None:
        return default

    try:

        if pd.isna(value):
            return default

    except Exception:
        pass

    value = str(value).strip()

    if not value:
        return default

    if value.lower() in [
        "nan",
        "none",
        "null"
    ]:
        return default

    return value


# ==============================================================================
# UOM HELPER
# ==============================================================================

def safe_uom(value, default="Pcs"):
    """
    Menjaga nilai UoM asli.

    Prioritas:
    1. Nilai UoM yang diberikan.
    2. Tidak memaksa menjadi Pcs.
    3. Default hanya jika benar-benar kosong.
    """

    if value is None:
        return default

    if isinstance(value, float) and pd.isna(value):
        return default

    value = str(value).strip()

    if not value:
        return default

    if value.lower() in [
        "nan",
        "none",
        "null"
    ]:
        return default

    return value


def get_uom_from_row(row, default="Pcs"):
    """
    Mengambil UoM tanpa memaksa menjadi Pcs.

    Mendukung:
    - UoM
    - uom
    - UOM
    """

    if row is None:
        return default

    candidates = [
        "UoM",
        "uom",
        "UOM"
    ]

    for key in candidates:

        if key in row:

            value = row.get(key)

            if value is not None:

                if not (
                    isinstance(value, float)
                    and pd.isna(value)
                ):

                    value = str(value).strip()

                    if value and value.lower() not in [
                        "nan",
                        "none",
                        "null"
                    ]:

                        return value

    return default


# ==============================================================================
# MASTER ITEM - UOM
# ==============================================================================

@st.cache_data(
    ttl=3600,
    show_spinner=False
)
def load_master_item_uom():

    """
    Membaca Master Item.

    Struktur yang digunakan:
        Kolom A = Material Code
        Kolom B = Material Name
        Kolom C = Specification
        Kolom D = UoM

    Return:
        {
            "AC0001": "Pcs",
            "MM0003": "Unit",
            ...
        }
    """

    try:

        all_values = get_sheet_values(
            MASTER_ITEM_SHEET
        )

        if not all_values or len(all_values) < 2:

            return {}

        uom_map = {}

        headers = [
            str(x).strip()
            for x in all_values[0]
        ]

        def find_header(candidates):

            lower_map = {
                str(header).strip().lower(): index
                for index, header in enumerate(headers)
            }

            for candidate in candidates:

                key = (
                    str(candidate)
                    .strip()
                    .lower()
                )

                if key in lower_map:

                    return lower_map[key]

            return None

        code_col = find_header([
            "Material Code",
            "MaterialCode",
            "Code",
            "Item Code",
            "Kode Material",
            "Kode"
        ])

        uom_col = find_header([
            "UoM",
            "UOM",
            "uom",
            "Unit",
            "Unit of Measure"
        ])

        if code_col is None:
            code_col = 0

        if uom_col is None:
            uom_col = 3

        for row in all_values[1:]:

            if len(row) <= code_col:
                continue

            if len(row) <= uom_col:
                continue

            material_code = safe_text(
                row[code_col],
                default=""
            )

            material_uom = safe_uom(
                row[uom_col],
                default=""
            )

            if not material_code:
                continue

            uom_map[
                material_code.upper()
            ] = material_uom

        return uom_map

    except Exception as e:

        st.warning(
            "⚠️ Gagal membaca UoM dari "
            f"sheet '{MASTER_ITEM_SHEET}': {e}"
        )

        return {}


def get_master_uom(material_code, default="Pcs"):

    code = safe_text(
        material_code,
        default=""
    )

    if not code:
        return default

    uom_map = load_master_item_uom()

    uom = uom_map.get(
        code.upper(),
        ""
    )

    if uom:

        return safe_uom(
            uom,
            default=default
        )

    return default


# ==============================================================================
# SUCCESS NOTIFICATION
# ==============================================================================

def show_pending_do_notification():

    notification = st.session_state.get(
        "do_success_notification"
    )

    if not notification:
        return

    no_do = notification.get(
        "no_do",
        ""
    )

    site_count = notification.get(
        "site_count",
        0
    )

    material_count = notification.get(
        "material_count",
        0
    )

    st.toast(
        f"✅ SUKSES! Delivery Order {no_do} berhasil dibuat dan disimpan.",
        icon="🎉"
    )

    st.success(
        f"""
        ### ✅ SUKSES — Delivery Order Berhasil Dibuat!

        **Nomor DO:** `{no_do}`

        **Status:** Berhasil dibuat dan disimpan ke **DB Material Out**

        **Site Allocated:** `{site_count} Site`

        **Total Material Row:** `{material_count} Row`

        Silakan buka tab **🖨️ Preview & PDF Cetak (A5)** untuk
        melihat atau mencetak Delivery Order.
        """,
        icon="✅"
    )

    st.session_state.do_success_notification = None


# ==============================================================================
# DO NUMBER
# ==============================================================================

def get_current_do_number():

    if st.session_state.get("do_number_draft"):

        return st.session_state["do_number_draft"]

    try:

        number = generate_do_number(
            is_reloc=False
        )

        st.session_state["do_number_draft"] = number

        st.session_state["do_number_generated_at"] = (
            datetime.now()
        )

        return number

    except Exception as e:

        st.error(
            "❌ Gagal membuat Nomor DO otomatis.\n\n"
            f"Detail: {e}"
        )

        return ""


def reset_do_number():

    st.session_state["do_number_draft"] = None

    st.session_state["do_number_generated_at"] = None


# ==============================================================================
# QUERY DATA
# ==============================================================================

@st.cache_data(
    ttl=900,
    show_spinner=False
)
def fetch_raw_query_data_cached():

    try:

        df = get_query_sheet_data()

        if isinstance(df, pd.DataFrame):

            return df.copy()

        return pd.DataFrame()

    except Exception as e:

        st.warning(
            f"⚠️ Gagal membaca data Query: {e}"
        )

        return pd.DataFrame()


def fetch_raw_query_data():

    return fetch_raw_query_data_cached()


# ==============================================================================
# STANDARD CHARGING TYPE - READ SHEET
# ==============================================================================

@st.cache_data(
    ttl=300,
    show_spinner=False
)
def load_standard_charging_type_sheet():

    """
    Membaca Sheet:

        Standart Charging Type

    Struktur:

        A = Material Code
        B = Material Name
        C = 6S1P
        D = 12S1P
        E = DC20
        F = DC30
        G = DC60
        dst...

    Fungsi dibuat dinamis sehingga apabila nanti user menambahkan
    charging type baru pada kolom H, I, J, dst., kolom tersebut
    otomatis dapat digunakan.

    Return:
        DataFrame dengan header asli dari Google Sheet.
    """

    try:

        all_values = get_sheet_values(
            STANDARD_CHARGING_TYPE_SHEET
        )

        if not all_values:

            return pd.DataFrame()

        if len(all_values) < 1:

            return pd.DataFrame()

        headers = [
            safe_text(
                x,
                default=""
            )
            for x in all_values[0]
        ]

        # ----------------------------------------------------------------------
        # Hapus kolom kosong di bagian belakang.
        # ----------------------------------------------------------------------

        while headers and not headers[-1]:

            headers.pop()

        if len(headers) < 2:

            return pd.DataFrame()

        cleaned_rows = []

        for row in all_values[1:]:

            normalized_row = list(row)

            if len(normalized_row) < len(headers):

                normalized_row.extend(
                    [""] * (
                        len(headers)
                        - len(normalized_row)
                    )
                )

            elif len(normalized_row) > len(headers):

                normalized_row = normalized_row[
                    :len(headers)
                ]

            cleaned_rows.append(
                normalized_row
            )

        df = pd.DataFrame(
            cleaned_rows,
            columns=headers
        )

        return df

    except Exception as e:

        st.warning(
            "⚠️ Gagal membaca sheet "
            f"'{STANDARD_CHARGING_TYPE_SHEET}': {e}"
        )

        return pd.DataFrame()


# ==============================================================================
# DETECT STANDARD CHARGING TYPE COLUMNS
# ==============================================================================

def get_standard_charging_type_columns():

    df = load_standard_charging_type_sheet()

    if df.empty:

        return []

    columns = list(df.columns)

    if len(columns) <= 2:

        return []

    result = []

    # --------------------------------------------------------------------------
    # Sesuai struktur:
    # A = Material Code
    # B = Material Name
    # C dst = Charging Type
    # --------------------------------------------------------------------------

    for col in columns[2:]:

        charging_type = safe_text(
            col,
            default=""
        )

        if charging_type:

            result.append(
                charging_type
            )

    return result


# ==============================================================================
# MASTER DROPDOWN
# ==============================================================================

@st.cache_data(
    ttl=300,
    show_spinner=False
)
def load_master_dropdown():

    # --------------------------------------------------------------------------
    # Charging Type:
    #
    # Default lama tetap dipertahankan.
    # Charging Type baru yang ditambahkan pada header
    # "Standart Charging Type" otomatis ikut masuk.
    # --------------------------------------------------------------------------

    charging_types = DEFAULT_CHARGING_TYPES.copy()

    sheet_charging_types = (
        get_standard_charging_type_columns()
    )

    for charging_type in sheet_charging_types:

        if not any(
            str(existing).strip().lower()
            ==
            str(charging_type).strip().lower()
            for existing in charging_types
        ):

            charging_types.append(
                charging_type
            )

    expeditions = DEFAULT_EXPEDITIONS.copy()

    return charging_types, expeditions


# ==============================================================================
# REFRESH MASTER DATA
# ==============================================================================

def refresh_master_data():

    """
    Membersihkan cache master yang berkaitan dengan DO.

    Digunakan agar perubahan pada:
        - Standart Charging Type
        - Master Item
    dapat langsung dibaca kembali tanpa menunggu TTL.
    """

    try:
        load_standard_charging_type_sheet.clear()
    except Exception:
        pass

    try:
        load_standard_charging_materials.clear()
    except Exception:
        pass

    try:
        load_master_dropdown.clear()
    except Exception:
        pass

    try:
        load_master_item_uom.clear()
    except Exception:
        pass


# ==============================================================================
# COLUMN DETECTION
# ==============================================================================

def detect_query_columns(df):

    if df is None or df.empty:

        return {
            "epc": None,
            "charging": None,
            "status": None,
            "site": None
        }

    columns = list(df.columns)

    def find_column(candidates):

        for candidate in candidates:

            if candidate in columns:

                return candidate

        lower_map = {
            str(c).strip().lower(): c
            for c in columns
        }

        for candidate in candidates:

            key = (
                str(candidate)
                .strip()
                .lower()
            )

            if key in lower_map:

                return lower_map[key]

        return None

    return {

        "epc": find_column([
            "EPC Name",
            "EPC",
            "epc",
            "EPC name"
        ]),

        "charging": find_column([
            "Charging Type",
            "charging",
            "Charging"
        ]),

        "status": find_column([
            "Project Status",
            "status",
            "Status"
        ]),

        "site": find_column([
            "Project / Location Name",
            "Site Name",
            "Site",
            "Location",
            "site"
        ])

    }


# ==============================================================================
# EPC LIST
# ==============================================================================

@st.cache_data(
    ttl=900,
    show_spinner=False
)
def load_epc_list():

    df_query = fetch_raw_query_data_cached()

    if df_query.empty:

        return []

    columns = detect_query_columns(
        df_query
    )

    col_epc = columns["epc"]

    if not col_epc:

        return []

    values = (
        df_query[col_epc]
        .dropna()
        .astype(str)
        .str.strip()
    )

    values = [
        x
        for x in values.unique().tolist()
        if x
        and x.lower() not in [
            "nan",
            "none"
        ]
    ]

    return sorted(values)


# ==============================================================================
# USED SITE CACHE
# ==============================================================================

@st.cache_data(
    ttl=300,
    show_spinner=False
)
def get_used_sites_cached():

    try:

        sites = (
            get_used_sites_from_db_material_out()
        )

        if not sites:

            return []

        cleaned = []

        for site in sites:

            value = str(site).strip()

            if (
                value
                and value.lower() not in [
                    "nan",
                    "none"
                ]
            ):

                cleaned.append(value)

        return sorted(set(cleaned))

    except Exception as e:

        st.warning(
            f"⚠️ Gagal mengambil site yang sudah digunakan: {e}"
        )

        return []


def get_used_sites():

    return get_used_sites_cached()


# ==============================================================================
# FILTER SITE
# ==============================================================================

def load_filtered_sites(
    epc,
    charging_type
):

    df_query = fetch_raw_query_data_cached()

    if df_query.empty:

        return []

    columns = detect_query_columns(
        df_query
    )

    col_epc = columns["epc"]
    col_charging = columns["charging"]
    col_status = columns["status"]
    col_site = columns["site"]

    if not all([
        col_epc,
        col_charging,
        col_status,
        col_site
    ]):

        st.warning(
            "⚠️ Struktur kolom Sheet Query tidak sesuai. "
            "Pastikan terdapat EPC Name, Charging Type, "
            "Project Status, dan Project / Location Name."
        )

        return []

    target_epc = (
        str(epc).strip().lower()
        if epc
        else ""
    )

    target_charging = (
        str(charging_type).strip().lower()
        if charging_type
        else ""
    )

    used_sites = set(
        x.strip()
        for x in get_used_sites()
        if str(x).strip()
    )

    working_df = df_query.copy()

    working_df["_epc_clean"] = (
        working_df[col_epc]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    working_df["_charging_clean"] = (
        working_df[col_charging]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    working_df["_status_clean"] = (
        working_df[col_status]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    working_df["_site_clean"] = (
        working_df[col_site]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    mask = (

        (
            working_df["_epc_clean"]
            ==
            target_epc
        )

        &

        (
            working_df["_charging_clean"]
            ==
            target_charging
        )

        &

        (
            ~working_df["_status_clean"]
            .str.contains(
                "drop|cancel",
                regex=True,
                na=False
            )
        )

        &

        (
            working_df["_site_clean"]
            != ""
        )

    )

    filtered = working_df.loc[
        mask
    ]

    sites = []

    for site in filtered[
        "_site_clean"
    ].tolist():

        if site not in used_sites:

            if site not in sites:

                sites.append(site)

    return sites


# ==============================================================================
# AVAILABLE RELOCATION SITE
# ==============================================================================

@st.cache_data(
    ttl=900,
    show_spinner=False
)
def load_available_relocation_sites():

    df_query = fetch_raw_query_data_cached()

    if df_query.empty:

        return []

    columns = detect_query_columns(
        df_query
    )

    col_status = columns["status"]
    col_site = columns["site"]

    if not col_status or not col_site:

        return []

    working_df = df_query.copy()

    working_df["_status_clean"] = (
        working_df[col_status]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    working_df["_site_clean"] = (
        working_df[col_site]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    filtered = working_df[
        (
            ~working_df["_status_clean"]
            .str.contains(
                "drop|cancel",
                regex=True,
                na=False
            )
        )
        &
        (
            working_df["_site_clean"]
            != ""
        )
    ]

    sites = []

    for site in filtered[
        "_site_clean"
    ].tolist():

        if site not in sites:

            sites.append(site)

    return sites


# ==============================================================================
# SAFE NUMERIC QTY
# ==============================================================================

def safe_qty(value, default=0):

    if value is None:

        return default

    if isinstance(value, str):

        value = value.strip()

        if value == "":

            return default

        value = value.replace(",", "")

    try:

        number = float(value)

        if number.is_integer():

            return int(number)

        return number

    except (ValueError, TypeError):

        return default


# ==============================================================================
# QTY BREAKDOWN PER SITE
# ==============================================================================

def calculate_qty_per_site(
    total_qty,
    site_count
):
    """
    Menghitung Qty material yang dialokasikan untuk setiap site.

    Formula:

        Qty Per Site =
            Total Qty / Jumlah Site

    Contoh:

        Total Qty = 15
        Site      = 3

        Qty Per Site = 15 / 3 = 5

    Fungsi ini sengaja dipisahkan agar:
    - logic kalkulasi mudah dibaca
    - Qty yang masuk DB konsisten
    - PDF otomatis membaca Qty hasil breakdown
    - logic lain seperti relokasi tidak perlu berubah
    """

    total_qty = safe_qty(
        total_qty,
        default=0
    )

    site_count = safe_qty(
        site_count,
        default=0
    )

    if site_count <= 0:

        return total_qty

    result = (
        float(total_qty)
        /
        float(site_count)
    )

    if result.is_integer():

        return int(result)

    return result


# ==============================================================================
# STANDARD MATERIAL
# ==============================================================================
@st.cache_data(
    ttl=300,
    show_spinner=False
)
def load_standard_charging_materials(
    charging_type
):

    """
    Membaca standard material langsung dari:

        Standart Charging Type

    Struktur sheet:

        A = Material Code
        B = Material Name
        C = 6S1P
        D = 12S1P
        E = DC20
        F = DC30
        G = DC60
        dst...

    Contoh:

        AC0001 | Clamp Conduit | 5 | 5 |   |   |   |

    Jika charging_type = 6S1P:

        AC0001 -> Qty 5

    Jika charging_type = DC20:

        AC0001 -> tidak masuk

    UoM tetap mengambil dari:
        Master Item
    """

    if not charging_type:

        return []

    df_standard = (
        load_standard_charging_type_sheet()
    )

    if df_standard.empty:

        return []

    columns = list(
        df_standard.columns
    )

    if len(columns) < 2:

        return []

    # --------------------------------------------------------------------------
    # Cari Material Code dan Material Name.
    #
    # Sesuai struktur user:
    # A = Material Code
    # B = Material Name
    # --------------------------------------------------------------------------

    def find_column(
        candidates,
        fallback=None
    ):

        lower_map = {
            str(col).strip().lower(): col
            for col in columns
        }

        for candidate in candidates:

            key = (
                str(candidate)
                .strip()
                .lower()
            )

            if key in lower_map:

                return lower_map[key]

        return fallback

    code_col = find_column(
        [
            "Material Code",
            "MaterialCode",
            "Code",
            "Item Code",
            "Kode Material",
            "Kode"
        ],
        fallback=columns[0]
    )

    name_col = find_column(
        [
            "Material Name",
            "MaterialName",
            "Name",
            "Item Name",
            "Nama Material",
            "Nama"
        ],
        fallback=columns[1]
    )

    # --------------------------------------------------------------------------
    # Cari kolom Charging Type.
    # --------------------------------------------------------------------------

    target_charging = safe_text(
        charging_type,
        default=""
    )

    target_column = None

    for col in columns[2:]:

        if (
            str(col).strip().lower()
            ==
            target_charging.lower()
        ):

            target_column = col

            break

    # --------------------------------------------------------------------------
    # Kalau charging type tidak ditemukan di sheet,
    # return kosong.
    # --------------------------------------------------------------------------

    if target_column is None:

        return []

    result = []

    # --------------------------------------------------------------------------
    # LOOP MATERIAL
    # --------------------------------------------------------------------------

    for _, row in df_standard.iterrows():

        material_code = safe_text(
            row.get(
                code_col,
                ""
            ),
            default=""
        )

        material_name = safe_text(
            row.get(
                name_col,
                ""
            ),
            default=""
        )

        if not material_code:

            continue

        # ----------------------------------------------------------------------
        # Ambil quantity dari kolom charging type.
        # ----------------------------------------------------------------------

        raw_qty = row.get(
            target_column,
            ""
        )

        std_qty = safe_qty(
            raw_qty,
            default=0
        )

        # ----------------------------------------------------------------------
        # Hanya material dengan quantity > 0 yang ditampilkan.
        # ----------------------------------------------------------------------

        if std_qty <= 0:

            continue

        # ----------------------------------------------------------------------
        # UoM TETAP dari Master Item.
        # ----------------------------------------------------------------------

        master_uom = get_master_uom(
            material_code,
            default="Pcs"
        )

        result.append({

            "code":
                material_code,

            "name":
                material_name,

            "std_qty":
                std_qty,

            "uom":
                master_uom

        })

    return result


# ==============================================================================
# ENSURE RELOCATION COLUMNS
# ==============================================================================

def ensure_relocation_columns(df):

    relocation_columns = [

        "Date Reloc.",
        "No. DO Reloc.",
        "Qty Reloc.",
        "Site Reloc.",
        "Mitra Reloc.",
        "Remarks Reloc."

    ]

    result = df.copy()

    for col in relocation_columns:

        if col not in result.columns:

            result[col] = ""

    return result


# ==============================================================================
# PDF GENERATOR
# ==============================================================================

def generate_do_a5_pdf(data):

    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("ReportLab belum terpasang.")

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=portrait(A5),
        rightMargin=15,
        leftMargin=15,
        topMargin=15,
        bottomMargin=15
    )

    elements = []
    styles = getSampleStyleSheet()

    # Font styles dibuat lebih compact (leading diperkecil)
    title_style = ParagraphStyle(
        "T", fontName="Helvetica-Bold", fontSize=10, textColor=colors.HexColor("#1a365d")
    )
    subtitle_style = ParagraphStyle(
        "ST", fontName="Helvetica", fontSize=6, textColor=colors.HexColor("#4a5568"), leading=7
    )
    body_style = ParagraphStyle(
        "B", fontName="Helvetica", fontSize=6.5, leading=7.5, textColor=colors.HexColor("#2d3748")
    )
    body_bold = ParagraphStyle(
        "BB", fontName="Helvetica-Bold", fontSize=6.5, leading=7.5, textColor=colors.HexColor("#1a365d")
    )
    header_table_style = ParagraphStyle(
        "HT", fontName="Helvetica-Bold", fontSize=6.5, leading=7.5, textColor=colors.white, alignment=1
    )
    center_style = ParagraphStyle(
        "C", fontName="Helvetica", fontSize=6.5, leading=7.5, textColor=colors.HexColor("#2d3748"), alignment=1
    )

    # --------------------------------------------------------------------------
    # LOGO & COMPANY INFO
    # --------------------------------------------------------------------------
    logo_path = "assets/logo.png"
    if os.path.exists(logo_path):
        logo_img = RLImage(logo_path, width=90, height=25)
    else:
        logo_img = Paragraph("<b>PT. CLX</b>", title_style)

    company_info = [
        Paragraph("<b>PT. Connectivity Leads excellence</b>", title_style),
        Paragraph("Jl. M Ali 2 No. 19 RT 007 RW 004 Tanah Baru, Beji, Kota Depok, Jawa barat 16426", subtitle_style),
        Paragraph("E: clx.central@gmail.com | T: +62 821-4858-1879", subtitle_style)
    ]

    head_table = Table([[logo_img, company_info]], colWidths=[95, 290])
    head_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 0), (-1, -1), 1, colors.HexColor("#1a365d")),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4)
    ]))
    elements.append(head_table)
    elements.append(Spacer(1, 5))

    # --------------------------------------------------------------------------
    # TO BOX
    # --------------------------------------------------------------------------
    to_box = [
        [Paragraph("<b>To</b>", body_bold), ""],
        [Paragraph("Name:", body_style), Paragraph(str(data.get("to", "")), body_bold)],
        [Paragraph("Phone No.:", body_style), Paragraph(str(data.get("contact", "")), body_style)],
        [Paragraph("Address:", body_style), Paragraph(str(data.get("address", "")), body_style)]
    ]

    to_table = Table(to_box, colWidths=[45, 140])
    to_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e0")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#edf2f7")),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)
    ]))

    # --------------------------------------------------------------------------
    # META BOX
    # --------------------------------------------------------------------------
    meta_box = [
        [Paragraph("<b>DELIVERY ORDER</b>", ParagraphStyle("DO", fontName="Helvetica-Bold", fontSize=8, alignment=1, textColor=colors.HexColor("#1a365d"))), ""],
        [Paragraph("No. DO:", body_bold), Paragraph(str(data.get("no_do", "")), body_bold)],
        [Paragraph("Date:", body_style), Paragraph(str(data.get("date", "")), body_style)],
        [Paragraph("EPC:", body_style), Paragraph(str(data.get("epc", "")), body_style)],
        [Paragraph("Charging Type:", body_style), Paragraph(str(data.get("charging_type", "-")), body_style)],
        [Paragraph("Expedition:", body_style), Paragraph(str(data.get("expedition", "-")), body_style)]
    ]

    meta_table = Table(meta_box, colWidths=[65, 135])
    meta_table.setStyle(TableStyle([
        ("SPAN", (0, 0), (1, 0)),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e0")),
        ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#e2e8f0")),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)
    ]))

    top_info_table = Table([[to_table, meta_table]], colWidths=[185, 200])
    top_info_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    elements.append(top_info_table)
    elements.append(Spacer(1, 5))

    # --------------------------------------------------------------------------
    # MATERIAL TABLE
    # --------------------------------------------------------------------------
    mat_headers = [
        Paragraph("No", header_table_style),
        Paragraph("Material Code", header_table_style),
        Paragraph("Material Name", header_table_style),
        Paragraph("Qty", header_table_style),
        Paragraph("UoM", header_table_style),
        Paragraph("Site Allocation", header_table_style),
        Paragraph("Remarks", header_table_style)
    ]
    mat_rows = [mat_headers]

    materials = data.get("materials", [])
    for idx, item in enumerate(materials, start=1):
        code = item.get("Material Code", item.get("code", ""))
        name = item.get("Material Name", item.get("name", ""))
        uom = get_uom_from_row(item, default="")
        if not uom:
            uom = get_master_uom(code, default="Pcs")
            
        site = item.get("Site Alocation")
        if site is None:
            site = item.get("Site Allocation", "")
        if site is None:
            site = ""

        qty = safe_qty(item.get("Qty", 0), default=0)
        remarks = item.get("Remarks", "")
        if remarks is None:
            remarks = ""

        mat_rows.append([
            Paragraph(str(idx), center_style),
            Paragraph(str(code), body_style),
            Paragraph(str(name), body_style),
            Paragraph(str(qty), center_style),
            Paragraph(str(uom), center_style),
            Paragraph(str(site), body_style),
            Paragraph(str(remarks), body_style)
        ])

    # TOTAL SITE ROW - logic tidak diubah
    site_values = []
    for material in materials:
        site = material.get("Site Alocation")
        if site is None:
            site = material.get("Site Allocation", "")
        if site is None:
            site = ""
        site = str(site).strip()
        if site:
            site_values.append(site)
            
    site_allocated_count = data.get("site_count", len(set(site_values)))
    
    mat_rows.append([
        Paragraph("<b>TOTAL SITE</b>", ParagraphStyle("R", fontName="Helvetica-Bold", fontSize=6.5, leading=7.5, alignment=2)),
        "", "", "", "",
        Paragraph(f"<b>{site_allocated_count} Site Allocated</b>", ParagraphStyle("L", fontName="Helvetica-Bold", fontSize=6.5, leading=7.5)),
        ""
    ])

    # Proporsi lebar kolom yang baru (Total = 385)
    # No(15) | Code(45) | Name(115) | Qty(20) | UoM(28) | Site(100) | Remarks(62)
    materials_table = Table(
        mat_rows,
        colWidths=[15, 45, 115, 20, 28, 100, 62]
    )
    materials_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a365d")),
        ("GRID", (0, 0), (-1, -2), 0.5, colors.HexColor("#cbd5e0")),
        ("SPAN", (0, -1), (4, -1)),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#edf2f7")),
        ("BOX", (0, -1), (-1, -1), 0.5, colors.HexColor("#1a365d")),
        # Padding di-press seminimal mungkin agar isi compact & tinggi baris mengecil
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("VALIGN", (0, 0), (-1, -1), "TOP"), # Top align agar teks panjang turunnya rapi
    ]))
    elements.append(materials_table)
    elements.append(Spacer(1, 8))

    # --------------------------------------------------------------------------
    # SIGNATURE
    # --------------------------------------------------------------------------
    sign_title = ParagraphStyle("SIGN", fontName="Helvetica-Bold", fontSize=6.5, alignment=1)
    sign_data = [
        [
            Paragraph("Prepared By,", sign_title),
            Paragraph("Approved By,", sign_title),
            Paragraph("Received By,", sign_title)
        ],
        ["", "", ""],
        [
            "( ____________________ )",
            "( ____________________ )",
            "( ____________________ )"
        ]
    ]

    sign_table = Table(sign_data, colWidths=[128, 129, 128])
    sign_table.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 20),
        ("TOPPADDING", (0, 0), (-1, -1), 1)
    ]))
    elements.append(sign_table)

    doc.build(elements)
    buffer.seek(0)
    
    return buffer.getvalue()
    )

    # --------------------------------------------------------------------------
    # MATERIAL TABLE
    #
    # Qty di sini langsung membaca Qty yang sudah tersimpan.
    #
    # Karena pada saat CREATE DO Qty sudah dibreakdown:
    #
    #   Total Qty = Standard Qty x Site Count
    #   Qty / Site = Total Qty / Site Count
    #
    # Maka PDF tidak melakukan pembagian ulang.
    # --------------------------------------------------------------------------

    mat_headers = [

        Paragraph("No", header_table_style),

        Paragraph(
            "Material Code",
            header_table_style
        ),

        Paragraph(
            "Material Name",
            header_table_style
        ),

        Paragraph("Qty", header_table_style),

        Paragraph("UoM", header_table_style),

        Paragraph(
            "Site Allocation",
            header_table_style
        ),

        Paragraph(
            "Remarks",
            header_table_style
        )

    ]

    mat_rows = [mat_headers]

    materials = data.get(
        "materials",
        []
    )

    for idx, item in enumerate(
        materials,
        start=1
    ):

        code = item.get(
            "Material Code",
            item.get("code", "")
        )

        name = item.get(
            "Material Name",
            item.get("name", "")
        )

        uom = get_uom_from_row(
            item,
            default=""
        )

        if not uom:

            uom = get_master_uom(
                code,
                default="Pcs"
            )

        site = item.get(
            "Site Alocation"
        )

        if site is None:

            site = item.get(
                "Site Allocation",
                ""
            )

        if site is None:

            site = ""

        # ----------------------------------------------------------------------
        # QTY SUDAH DALAM BENTUK BREAKDOWN PER SITE
        # ----------------------------------------------------------------------

        qty = safe_qty(
            item.get(
                "Qty",
                0
            ),
            default=0
        )

        remarks = item.get(
            "Remarks",
            ""
        )

        if remarks is None:

            remarks = ""

        mat_rows.append([

            Paragraph(
                str(idx),
                ParagraphStyle(
                    "C",
                    alignment=1,
                    fontSize=6
                )
            ),

            Paragraph(
                str(code),
                body_style
            ),

            Paragraph(
                str(name),
                body_style
            ),

            Paragraph(
                str(qty),
                ParagraphStyle(
                    "C2",
                    alignment=1,
                    fontSize=6
                )
            ),

            Paragraph(
                str(uom),
                ParagraphStyle(
                    "C3",
                    alignment=1,
                    fontSize=6
                )
            ),

            Paragraph(
                str(site),
                body_style
            ),

            Paragraph(
                str(remarks),
                body_style
            )

        ])

    # --------------------------------------------------------------------------
    # TOTAL SITE
    # --------------------------------------------------------------------------

    site_values = []

    for material in materials:

        site = material.get(
            "Site Alocation"
        )

        if site is None:

            site = material.get(
                "Site Allocation",
                ""
            )

        if site is None:

            site = ""

        site = str(site).strip()

        if site:

            site_values.append(site)

    site_allocated_count = data.get(
        "site_count",
        len(set(site_values))
    )

    mat_rows.append([

        Paragraph(
            "<b>TOTAL SITE</b>",
            ParagraphStyle(
                "R",
                fontName="Helvetica-Bold",
                fontSize=6.5,
                alignment=2
            )
        ),

        "",
        "",
        "",
        "",

        Paragraph(
            f"<b>{site_allocated_count} Site Allocated</b>",
            ParagraphStyle(
                "L",
                fontName="Helvetica-Bold",
                fontSize=6.5
            )
        ),

        ""

    ])

    materials_table = Table(
        mat_rows,
        colWidths=[
            18,
            50,
            95,
            25,
            22,
            120,
            60
        ]
    )

    materials_table.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor("#1a365d")
            ),
            (
                "GRID",
                (0, 0),
                (-1, -2),
                0.5,
                colors.HexColor("#cbd5e0")
            ),
            (
                "SPAN",
                (0, -1),
                (4, -1)
            ),
            (
                "BACKGROUND",
                (0, -1),
                (-1, -1),
                colors.HexColor("#edf2f7")
            ),
            (
                "BOX",
                (0, -1),
                (-1, -1),
                0.5,
                colors.HexColor("#1a365d")
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
            )
        ])
    )

    elements.append(materials_table)

    elements.append(
        Spacer(1, 8)
    )

    # --------------------------------------------------------------------------
    # SIGNATURE
    # --------------------------------------------------------------------------

    sign_title = ParagraphStyle(
        "SIGN",
        fontName="Helvetica-Bold",
        fontSize=6.5,
        alignment=1
    )

    sign_data = [

        [
            Paragraph(
                "Prepared By,",
                sign_title
            ),

            Paragraph(
                "Approved By,",
                sign_title
            ),

            Paragraph(
                "Received By,",
                sign_title
            )
        ],

        [
            "",
            "",
            ""
        ],

        [
            "( ____________________ )",
            "( ____________________ )",
            "( ____________________ )"
        ]

    ]

    sign_table = Table(
        sign_data,
        colWidths=[
            130,
            130,
            130
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
                "BOTTOMPADDING",
                (0, 0),
                (-1, 0),
                20
            ),
            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                1
            )
        ])
    )

    elements.append(sign_table)

    doc.build(elements)

    buffer.seek(0)

    return buffer.getvalue()


# ==============================================================================
# SAVE VALIDATION
# ==============================================================================

def save_do_and_verify(rows_data, no_do):

    if not rows_data:

        return False, "Tidak ada data material yang akan disimpan."

    target_do = safe_text(
        no_do,
        default=""
    )

    if not target_do:

        return False, "Nomor DO kosong."

    try:

        save_result = (
            save_do_to_db_material_out(
                rows_data
            )
        )

        if not save_result:

            return False, (
                "Fungsi database mengembalikan status gagal "
                "saat menyimpan DB Material Out."
            )

        verify_values = get_sheet_values(
            "DB Material Out"
        )

        if not verify_values or len(verify_values) < 2:

            return False, (
                "Database tidak mengembalikan data setelah proses save."
            )

        found = False

        for row in verify_values[1:]:

            if len(row) > 1:

                existing_do = safe_text(
                    row[1],
                    default=""
                )

                if existing_do == target_do:

                    found = True

                    break

        if not found:

            return False, (
                f"DO {target_do} belum ditemukan pada "
                "sheet 'DB Material Out' setelah proses penyimpanan."
            )

        return True, ""

    except Exception as e:

        return False, (
            "Terjadi error saat menyimpan/verifikasi DO: "
            f"{type(e).__name__}: {e}"
        )


# ==============================================================================
# CSS
# ==============================================================================

def apply_page_style():

    st.markdown(
        """
        <style>

        div[data-baseweb="input"],
        div[data-baseweb="input"] > div,
        div[data-baseweb="textarea"],
        div[data-baseweb="textarea"] > textarea,
        div[data-baseweb="select"] > div,
        div[data-testid="stTextInput"] input,
        div[data-testid="stTextArea"] textarea,
        div[data-testid="stDateInput"] input {

            background-color: #FFFFFF !important;
            color: #000000 !important;
            -webkit-text-fill-color: #000000 !important;
            border: 1px solid #94A3B8 !important;
            border-radius: 6px !important;
            font-weight: 600 !important;
            opacity: 1 !important;

        }

        div[data-baseweb="popover"],
        div[data-baseweb="menu"] {

            background-color: #FFFFFF !important;
            color: #000000 !important;

        }

        li[role="option"] {

            color: #000000 !important;
            background-color: #FFFFFF !important;

        }

        li[role="option"]:hover {

            background-color: #E2E8F0 !important;

        }

        ::placeholder,
        textarea::placeholder,
        input::placeholder {

            color: #64748B !important;
            -webkit-text-fill-color: #64748B !important;
            opacity: 1 !important;

        }

        span[data-baseweb="tag"] {

            background-color: #3B82F6 !important;
            color: #FFFFFF !important;

        }

        div[data-testid="stDataFrame"],
        div[data-testid="stDataEditor"],
        .glideDataEditor {

            background-color: #FFFFFF !important;
            color: #000000 !important;
            border: 1px solid #CBD5E1 !important;

        }

        label,
        p,
        h1,
        h2,
        h3,
        h4,
        .stMarkdown {

            color: #0F172A !important;

        }

        div[data-testid="stToast"] {

            font-size: 16px !important;
            font-weight: 700 !important;

        }

        </style>
        """,
        unsafe_allow_html=True
    )


# ==============================================================================
# MAIN RENDER
# ==============================================================================

def render():

    initialize_session_state()

    apply_page_style()

    show_pending_do_notification()

    # ==========================================================================
    # HEADER
    # ==========================================================================

    st.title(
        "🚚 Delivery Order (DO) Generator"
    )

    st.caption(
        "Divisi Supply Chain Management (SCM) - "
        "Create, Print, Search & Relocation DO"
    )

    # ==========================================================================
    # MASTER DROPDOWN
    # ==========================================================================

    charging_list, exp_list = (
        load_master_dropdown()
    )

    epc_list = load_epc_list()

    # ==========================================================================
    # TABS
    # ==========================================================================

    tab_form, tab_preview, tab_search = st.tabs([
        "📝 Form Create DO",
        "🖨️ Preview & PDF Cetak (A5)",
        "🔍 Search, Edit & Relokasi Site"
    ])

    # ==========================================================================
    # TAB 1
    # ==========================================================================

    with tab_form:

        st.subheader(
            "Header Delivery Order"
        )

        # ----------------------------------------------------------------------
        # REFRESH MASTER DATA
        # ----------------------------------------------------------------------

        refresh_col1, refresh_col2 = st.columns(
            [5, 1]
        )

        with refresh_col1:

            st.caption(
                "Material standard dan quantity dibaca dari "
                f"sheet **{STANDARD_CHARGING_TYPE_SHEET}**. "
                "UoM tetap dibaca dari **Master Item**."
            )

        with refresh_col2:

            if st.button(
                "🔄 Refresh Master Data",
                key="btn_refresh_do_master"
            ):

                refresh_master_data()

                st.toast(
                    "Master Data berhasil direfresh.",
                    icon="🔄"
                )

                st.rerun()

        col1, col2, col3 = st.columns(3)

        with col1:

            no_do_auto = get_current_do_number()

            no_do = st.text_input(
                "1. No. DO (Auto)",
                value=no_do_auto,
                disabled=True
            )

            do_date = st.date_input(
                "2. Date",
                datetime.now()
            )

            epc = st.selectbox(
                "3. EPC (Query Sheet)",
                epc_list
                if epc_list
                else ["Pilih EPC..."],
                index=None,
                placeholder="Pilih EPC..."
            )

        with col2:

            charging_type = st.selectbox(
                "4. Charging Type (Master Dropdown)",
                charging_list,
                index=None,
                placeholder="Pilih Charging Type..."
            )

            expedition = st.selectbox(
                "5. Expedition (Master Dropdown)",
                exp_list,
                index=None,
                placeholder="Pilih Ekspedisi..."
            )

            to_name = st.text_input(
                "6. To (Recipient Name)",
                value="",
                placeholder="Contoh: Tsubasa Ozora"
            )

        with col3:

            contact = st.text_input(
                "7. Contact (Phone No.)",
                value="",
                placeholder="Contoh: 081234567890"
            )

            address = st.text_area(
                "8. Address",
                value="",
                placeholder="Contoh: Alamat Tujuan",
                height=110
            )

        st.divider()

        # ==========================================================================
        # FILTER SITE
        # ==========================================================================

        st.subheader(
            "Filter Site & Kalkulasi Material Automatic"
        )

        if epc and charging_type:

            available_sites = load_filtered_sites(
                epc,
                charging_type
            )

        else:

            available_sites = []

        selected_sites = st.multiselect(

            "Alokasi Site (Maksimal 15 Site terpilih):",

            options=available_sites,

            default=[],

            max_selections=MAX_SITE_SELECTION,

            placeholder=(
                "Pilih Alokasi Site..."
                if epc and charging_type
                else
                "⚠️ Silakan pilih EPC dan Charging Type terlebih dahulu..."
            )

        )

        site_count = len(
            selected_sites
        )

        st.info(
            f"📊 Total Site Terpilih: "
            f"**{site_count} Site Allocated** "
            f"(Maksimal {MAX_SITE_SELECTION} Site)"
        )

        # ==========================================================================
        # MATERIAL
        # ==========================================================================

        raw_materials = (

            load_standard_charging_materials(
                charging_type
            )

            if charging_type

            else []

        )

        table_data = []

        for idx, item in enumerate(
            raw_materials,
            start=1
        ):

            # ------------------------------------------------------------------
            # TOTAL QTY
            #
            # Standard Qty x jumlah Site
            #
            # Contoh:
            # Standard = 5
            # Site = 3
            #
            # Total Qty = 15
            #
            # Nilai Total Qty ini tetap ditampilkan pada editor agar user
            # dapat melakukan pengecekan/edit total kebutuhan.
            # ------------------------------------------------------------------

            total_qty = (

                item["std_qty"]

                *

                (
                    site_count
                    if site_count > 0
                    else 1
                )

            )

            # ------------------------------------------------------------------
            # UOM DARI MASTER ITEM
            # ------------------------------------------------------------------

            item_uom = get_master_uom(
                item["code"],
                default=item.get(
                    "uom",
                    "Pcs"
                )
            )

            table_data.append({

                "No":
                    idx,

                "Material Code":
                    item["code"],

                "Material Name":
                    item["name"],

                "Qty":
                    total_qty,

                "UoM":
                    item_uom,

                "Remarks":
                    ""

            })

        df_materials = pd.DataFrame(
            table_data
        )

        st.subheader(
            "Detail Material Item "
            "(Total Qty Akan Didistribusikan Per Site)"
        )

        edited_df = st.data_editor(

            df_materials,

            num_rows="dynamic",

            use_container_width=True,

            key="create_do_material_editor",

            column_config={

                "No":
                    st.column_config.NumberColumn(
                        width="small",
                        disabled=True
                    ),

                "Material Code":
                    st.column_config.TextColumn(
                        disabled=True
                    ),

                "Material Name":
                    st.column_config.TextColumn(
                        disabled=True
                    ),

                "Qty":
                    st.column_config.NumberColumn(
                        "Total Qty (Auto calculated)",
                        help=(
                            "Qty awal = Standard Qty x Total Site. "
                            "Saat disimpan, Qty akan dibreakdown "
                            "menjadi Qty per Site."
                        ),
                        min_value=0,
                        step=1
                    ),

                "UoM":
                    st.column_config.TextColumn(
                        "UoM",
                        disabled=True
                    ),

                "Remarks":
                    st.column_config.TextColumn(
                        "Remarks"
                    )

            }

        )

        # ----------------------------------------------------------------------
        # INFORMASI STANDARD MATERIAL
        # ----------------------------------------------------------------------

        if charging_type:

            standard_count = len(
                raw_materials
            )

            if standard_count > 0:

                st.caption(
                    f"📦 {standard_count} material standard "
                    f"ditemukan dari sheet "
                    f"**{STANDARD_CHARGING_TYPE_SHEET}** "
                    f"untuk charging type **{charging_type}**."
                )

            else:

                st.warning(
                    f"⚠️ Tidak ada material standard untuk "
                    f"**{charging_type}** pada sheet "
                    f"**{STANDARD_CHARGING_TYPE_SHEET}**."
                )

        # ----------------------------------------------------------------------
        # BREAKDOWN PREVIEW
        # ----------------------------------------------------------------------

        if site_count > 0 and not df_materials.empty:

            st.caption(
                "📐 **Breakdown Qty:** "
                "Total Qty ÷ Jumlah Site = Qty per Site. "
                "Qty per Site inilah yang akan tersimpan pada "
                "**DB Material Out**."
            )

        st.divider()

        # ==========================================================================
        # SAVE DO
        # ==========================================================================

        if st.button(
            "🚀 Simpan & Generate Delivery Order",
            type="primary",
            key="btn_create_do"
        ):

            if not epc or epc == "Pilih EPC...":

                st.error(
                    "EPC wajib dipilih!"
                )

            elif not charging_type:

                st.error(
                    "Charging Type wajib dipilih!"
                )

            elif not expedition:

                st.error(
                    "Expedition wajib dipilih!"
                )

            elif not to_name:

                st.error(
                    "Kolom 'To' wajib diisi!"
                )

            elif not address:

                st.error(
                    "Kolom 'Address' wajib diisi!"
                )

            elif site_count == 0:

                st.error(
                    "Pilih minimal 1 Site Allocation!"
                )

            elif site_count > MAX_SITE_SELECTION:

                st.error(
                    f"Maksimal {MAX_SITE_SELECTION} site."
                )

            elif edited_df.empty:

                st.error(
                    "Tidak ada material standard yang tersedia "
                    "untuk Charging Type tersebut."
                )

            else:

                date_str = do_date.strftime(
                    "%Y-%m-%d"
                )

                generated_db_rows = []

                row_counter = 1

                edited_material_rows = (
                    edited_df
                    .to_dict(
                        orient="records"
                    )
                )

                # ==================================================================
                # CREATE DATABASE ROW
                #
                # PERUBAHAN UTAMA:
                #
                # Sebelumnya:
                #
                #     Qty DB = Total Qty
                #
                # sehingga jika:
                #     Standard Qty = 5
                #     Site = 3
                #
                # maka:
                #     Site A = 15
                #     Site B = 15
                #     Site C = 15
                #
                # Sekarang:
                #
                #     Total Qty = 5 x 3 = 15
                #     Qty Per Site = 15 / 3 = 5
                #
                # sehingga:
                #     Site A = 5
                #     Site B = 5
                #     Site C = 5
                #
                # ==================================================================

                for site_name in selected_sites:

                    for mat_item in edited_material_rows:

                        material_code = safe_text(
                            mat_item.get(
                                "Material Code",
                                ""
                            ),
                            default=""
                        )

                        material_name = safe_text(
                            mat_item.get(
                                "Material Name",
                                ""
                            ),
                            default=""
                        )

                        if not material_code:

                            continue

                        # ----------------------------------------------------------
                        # TOTAL QTY DARI EDITOR
                        # ----------------------------------------------------------

                        total_qty = safe_qty(
                            mat_item.get(
                                "Qty",
                                0
                            ),
                            default=0
                        )

                        # ----------------------------------------------------------
                        # BREAKDOWN QTY PER SITE
                        #
                        # Formula:
                        #
                        # Qty Per Site =
                        #       Total Qty / Site Count
                        #
                        # Contoh:
                        #
                        # 15 / 3 = 5
                        # ----------------------------------------------------------

                        qty_per_site = calculate_qty_per_site(
                            total_qty,
                            site_count
                        )

                        # ----------------------------------------------------------
                        # UOM MASTER ITEM
                        # ----------------------------------------------------------

                        uom = get_master_uom(
                            material_code,
                            default=get_uom_from_row(
                                mat_item,
                                default="Pcs"
                            )
                        )

                        remarks = (
                            mat_item.get(
                                "Remarks",
                                ""
                            )
                        )

                        if remarks is None:

                            remarks = ""

                        # ----------------------------------------------------------
                        # SAVE ROW PER SITE
                        # ----------------------------------------------------------

                        generated_db_rows.append({

                            "No":
                                row_counter,

                            "No. DO":
                                no_do,

                            "Delv. Date":
                                date_str,

                            "Material Code":
                                material_code,

                            "Material Name":
                                material_name,

                            # ======================================================
                            # QTY HASIL BREAKDOWN PER SITE
                            # ======================================================

                            "Qty":
                                qty_per_site,

                            "UoM":
                                uom,

                            "Charging Type":
                                charging_type,

                            "Site Alocation":
                                site_name,

                            "Remarks":
                                remarks,

                            "To":
                                to_name,

                            "Phone No.":
                                contact,

                            "Address":
                                address,

                            "EPC":
                                epc,

                            "Date Reloc.":
                                "",

                            "No. DO Reloc.":
                                "",

                            "Qty Reloc.":
                                "",

                            "Site Reloc.":
                                "",

                            "Mitra Reloc.":
                                "",

                            "Remarks Reloc.":
                                ""

                        })

                        row_counter += 1

                # ==========================================================================
                # SAVE + VERIFY
                # ==========================================================================

                if not generated_db_rows:

                    st.error(
                        "Tidak ada material valid yang dapat disimpan."
                    )

                else:

                    with st.spinner(
                        "Menyimpan transaksi ke "
                        "sheet 'DB Material Out'..."
                    ):

                        save_ok, save_error = (
                            save_do_and_verify(
                                generated_db_rows,
                                no_do
                            )
                        )

                    if not save_ok:

                        st.error(
                            "❌ Delivery Order belum berhasil disimpan.\n\n"
                            f"{save_error}"
                        )

                    else:

                        st.session_state.current_do = {

                            "no_do":
                                no_do,

                            "date":
                                date_str,

                            "epc":
                                epc,

                            "charging_type":
                                charging_type,

                            "expedition":
                                expedition,

                            "to":
                                to_name,

                            "contact":
                                contact,

                            "address":
                                address,

                            "sites":
                                selected_sites,

                            "site_count":
                                site_count,

                            "materials":
                                generated_db_rows

                        }

                        st.session_state.do_success_notification = {

                            "no_do":
                                no_do,

                            "site_count":
                                site_count,

                            "material_count":
                                len(generated_db_rows),

                            "timestamp":
                                datetime.now().strftime(
                                    "%Y-%m-%d %H:%M:%S"
                                )

                        }

                        # ==================================================================
                        # CLEAR CACHE
                        # ==================================================================

                        get_used_sites_cached.clear()
                        fetch_raw_query_data_cached.clear()
                        load_epc_list.clear()
                        load_available_relocation_sites.clear()
                        load_master_item_uom.clear()
                        load_standard_charging_type_sheet.clear()
                        load_standard_charging_materials.clear()
                        load_master_dropdown.clear()

                        reset_do_number()

                        st.rerun()

    # ==========================================================================
    # TAB 2
    # ==========================================================================

    with tab_preview:

        st.subheader(
            "Preview PDF Delivery Order (A5 Format)"
        )

        do_data = st.session_state.get(
            "current_do"
        )

        if not REPORTLAB_AVAILABLE:

            st.error(
                "Library `reportlab` belum terpasang. "
                "Jalankan `pip install reportlab`."
            )

        elif not do_data:

            st.warning(
                "Belum ada Delivery Order yang dibuat/dipilih. "
                "Silakan isi form atau cari DO terlebih dahulu."
            )

        else:

            try:

                pdf_data = dict(do_data)

                pdf_materials = []

                for item in do_data.get(
                    "materials",
                    []
                ):

                    item_copy = dict(item)

                    material_code = (
                        item_copy.get(
                            "Material Code",
                            item_copy.get(
                                "code",
                                ""
                            )
                        )
                    )

                    item_copy["UoM"] = get_master_uom(
                        material_code,
                        default=get_uom_from_row(
                            item_copy,
                            default="Pcs"
                        )
                    )

                    # --------------------------------------------------------------
                    # QTY TIDAK DIBAGI LAGI DI SINI.
                    #
                    # Qty yang datang dari DB sudah merupakan Qty Per Site.
                    # --------------------------------------------------------------

                    item_copy["Qty"] = safe_qty(
                        item_copy.get(
                            "Qty",
                            0
                        ),
                        default=0
                    )

                    pdf_materials.append(
                        item_copy
                    )

                pdf_data["materials"] = (
                    pdf_materials
                )

                pdf_bytes = generate_do_a5_pdf(
                    pdf_data
                )

                st.download_button(

                    label=(
                        "🖨️ Download Delivery Order A5 "
                        f"({do_data['no_do'].replace('/', '_')}.pdf)"
                    ),

                    data=pdf_bytes,

                    file_name=(
                        f"DO_"
                        f"{do_data['no_do'].replace('/', '_')}"
                        f"_A5.pdf"
                    ),

                    mime="application/pdf",

                    type="primary",

                    key="download_do_pdf"

                )

            except Exception as e:

                st.error(
                    f"❌ Gagal membuat PDF: {e}"
                )

    # ==========================================================================
    # TAB 3
    # ==========================================================================

    with tab_search:

        st.subheader(
            "🔍 Cari, Edit & Relokasi Site Delivery Order"
        )

        st.caption(
            "Cari DO berdasarkan Nomor DO untuk mengedit data, "
            "merelokasi site material, atau melihat histori relokasi."
        )

        existing_dos = get_all_do_numbers()

        col_s1, col_s2 = st.columns([3, 1])

        with col_s1:

            selected_do_search = st.selectbox(
                "Pilih Nomor DO yang Tersimpan:",
                options=[""] + existing_dos,
                key="selected_do_search"
            )

        with col_s2:

            st.write("")
            st.write("")

            btn_search = st.button(
                "🔎 Cari DO",
                type="primary",
                key="btn_search_do"
            )

        if btn_search and selected_do_search:

            with st.spinner(
                f"Mencari data {selected_do_search}..."
            ):

                found_data = get_do_by_number(
                    selected_do_search
                )

            if found_data:

                found_materials = []

                for item in found_data.get(
                    "materials",
                    []
                ):

                    item_copy = dict(item)

                    material_code = (
                        item_copy.get(
                            "Material Code",
                            ""
                        )
                    )

                    item_copy["UoM"] = get_master_uom(
                        material_code,
                        default=get_uom_from_row(
                            item_copy,
                            default="Pcs"
                        )
                    )

                    item_copy["Qty"] = safe_qty(
                        item_copy.get(
                            "Qty",
                            0
                        ),
                        default=0
                    )

                    found_materials.append(
                        item_copy
                    )

                found_data["materials"] = (
                    found_materials
                )

                st.session_state.edit_do_data = (
                    found_data
                )

                st.success(
                    f"Data {selected_do_search} ditemukan!"
                )

            else:

                st.error(
                    "Data DO tidak ditemukan di database."
                )

        # ==========================================================================
        # EDIT FORM
        # ==========================================================================

        if (
            "edit_do_data"
            in st.session_state
            and
            st.session_state.edit_do_data
        ):

            edit_data = (
                st.session_state.edit_do_data
            )

            st.divider()

            st.subheader(
                f"Edit Data DO: "
                f"{edit_data['no_do']}"
            )

            ecol1, ecol2, ecol3 = st.columns(3)

            with ecol1:

                e_no_do = st.text_input(
                    "No. DO",
                    value=edit_data.get(
                        "no_do",
                        ""
                    ),
                    disabled=True,
                    key="e_no_do"
                )

                e_date = st.text_input(
                    "Delivery Date",
                    value=edit_data.get(
                        "date",
                        ""
                    ),
                    key="e_date"
                )

            with ecol2:

                e_to = st.text_input(
                    "To (Recipient)",
                    value=edit_data.get(
                        "to",
                        ""
                    ),
                    key="e_to"
                )

                e_contact = st.text_input(
                    "Phone No.",
                    value=edit_data.get(
                        "contact",
                        ""
                    ),
                    key="e_contact"
                )

            with ecol3:

                e_epc = st.text_input(
                    "EPC",
                    value=edit_data.get(
                        "epc",
                        ""
                    ),
                    key="e_epc"
                )

                e_address = st.text_area(
                    "Address",
                    value=edit_data.get(
                        "address",
                        ""
                    ),
                    key="e_address",
                    height=100
                )

            st.write(
                "**Material Items per Site Allocation:**"
            )

            df_edit_mat = pd.DataFrame(
                edit_data.get(
                    "materials",
                    []
                )
            )

            df_edit_mat = ensure_relocation_columns(
                df_edit_mat
            )

            if "UoM" not in df_edit_mat.columns:

                df_edit_mat["UoM"] = ""

            if "Material Code" in df_edit_mat.columns:

                df_edit_mat["UoM"] = (
                    df_edit_mat.apply(
                        lambda row: get_master_uom(
                            row.get(
                                "Material Code",
                                ""
                            ),
                            default=get_uom_from_row(
                                row,
                                default="Pcs"
                            )
                        ),
                        axis=1
                    )
                )

            else:

                df_edit_mat["UoM"] = (
                    df_edit_mat["UoM"]
                    .apply(
                        lambda x: safe_uom(
                            x,
                            default="Pcs"
                        )
                    )
                )

            cols_to_show = [

                "No",
                "No. DO",
                "Delv. Date",
                "Material Code",
                "Material Name",
                "Qty",
                "UoM",
                "Charging Type",
                "Site Alocation",
                "Remarks",
                "To",
                "Phone No.",
                "Address",
                "EPC",
                "Date Reloc.",
                "No. DO Reloc.",
                "Qty Reloc.",
                "Site Reloc.",
                "Mitra Reloc.",
                "Remarks Reloc."

            ]

            cols_existing = [
                c
                for c in cols_to_show
                if c in df_edit_mat.columns
            ]

            edited_mat_df = st.data_editor(

                df_edit_mat[cols_existing],

                num_rows="dynamic",

                use_container_width=True,

                key="editor_search_do",

                column_config={

                    "Qty":
                        st.column_config.NumberColumn(
                            "Qty",
                            min_value=0,
                            step=1
                        ),

                    "UoM":
                        st.column_config.TextColumn(
                            "UoM",
                            disabled=True
                        ),

                    "Remarks":
                        st.column_config.TextColumn(
                            "Remarks"
                        ),

                    "Site Alocation":
                        st.column_config.TextColumn(
                            "Site Allocation"
                        )

                }

            )

            edited_mat_df = ensure_relocation_columns(
                edited_mat_df
            )

            # ==========================================================================
            # RELOCATION
            # ==========================================================================

            st.markdown("---")

            st.subheader(
                "🔁 Form Eksekusi Relokasi Site Material"
            )

            st.info(
                "Fitur ini akan memperbarui Kolom O:T "
                "pada DO Asal dan otomatis membuat "
                "baris DO Relokasi Baru di DB."
            )

            with st.expander(
                "📌 Klik di sini untuk Melakukan Relokasi Site",
                expanded=True
            ):

                raw_sites_in_do = []

                if "Site Alocation" in edited_mat_df.columns:

                    raw_sites_in_do.extend(
                        edited_mat_df[
                            "Site Alocation"
                        ]
                        .dropna()
                        .astype(str)
                        .tolist()
                    )

                if "Site Allocation" in edited_mat_df.columns:

                    raw_sites_in_do.extend(
                        edited_mat_df[
                            "Site Allocation"
                        ]
                        .dropna()
                        .astype(str)
                        .tolist()
                    )

                current_do_sites = []

                for site in raw_sites_in_do:

                    clean_site = str(
                        site
                    ).strip()

                    if (
                        clean_site
                        and
                        clean_site not in current_do_sites
                        and
                        not clean_site.isdigit()
                        and
                        clean_site.lower()
                        not in [
                            "none",
                            "nan"
                        ]
                    ):

                        current_do_sites.append(
                            clean_site
                        )

                all_query_sites = (
                    load_available_relocation_sites()
                )

                used_sites_set = set(
                    current_do_sites
                )

                selectable_new_sites = [

                    site

                    for site in all_query_sites

                    if site not in used_sites_set

                ]

                col_r1, col_r2 = st.columns(2)

                with col_r1:

                    selected_site_old = st.selectbox(

                        "Pilih Site Asal yang Ingin Direlokasi:",

                        options=(
                            current_do_sites
                            if current_do_sites
                            else
                            ["Tidak Ada Site"]
                        ),

                        key="reloc_old_site"

                    )

                with col_r2:

                    selected_site_new = st.selectbox(

                        "Nama Site Tujuan Baru (New Site):",

                        options=(
                            selectable_new_sites
                            if selectable_new_sites
                            else
                            [
                                "Tidak ada site baru yang tersedia"
                            ]
                        ),

                        key="reloc_new_site"

                    )

                reloc_mitra = st.text_input(
                    "Mitra Relokasi:",
                    placeholder="Contoh: PT Mitra Jaya",
                    key="reloc_mitra"
                )

                reloc_remarks = st.text_input(
                    "Alasan / Catatan Relokasi:",
                    placeholder=(
                        "Contoh: Perubahan WO Lapangan / "
                        "Re-alloc Site"
                    ),
                    key="reloc_reason"
                )

                # ==========================================================================
                # EXECUTE RELOCATION
                # ==========================================================================

                if st.button(
                    "🔀 Eksekusi Relokasi Site",
                    type="secondary",
                    key="btn_execute_relocation"
                ):

                    if (
                        selected_site_old
                        ==
                        "Tidak Ada Site"
                    ):

                        st.error(
                            "Site asal tidak ditemukan!"
                        )

                    elif (
                        not selected_site_new
                        or
                        selected_site_new
                        ==
                        "Tidak ada site baru yang tersedia"
                    ):

                        st.error(
                            "Silakan pilih Site Tujuan Baru yang valid!"
                        )

                    elif (
                        selected_site_old
                        ==
                        selected_site_new
                    ):

                        st.warning(
                            "Site Asal dan Site Tujuan Baru "
                            "tidak boleh sama!"
                        )

                    else:

                        reloc_date = datetime.now().strftime(
                            "%Y-%m-%d"
                        )

                        reloc_timestamp = datetime.now().strftime(
                            "%Y-%m-%d %H:%M"
                        )

                        try:

                            new_reloc_do_num = (
                                generate_do_number(
                                    is_reloc=True
                                )
                            )

                        except Exception as e:

                            st.error(
                                "❌ Gagal membuat "
                                "Nomor DO Relokasi.\n\n"
                                f"{e}"
                            )

                            new_reloc_do_num = ""

                        if not new_reloc_do_num:

                            st.stop()

                        new_reloc_rows_to_save = []

                        edited_relocation_df = (
                            edited_mat_df.copy()
                        )

                        relocation_count = 0

                        for idx, row in (
                            edited_relocation_df.iterrows()
                        ):

                            site_in_row = (

                                str(
                                    row.get(
                                        "Site Alocation",
                                        ""
                                    )
                                    or
                                    row.get(
                                        "Site Allocation",
                                        ""
                                    )
                                    or
                                    ""
                                )
                                .strip()
                            )

                            if (
                                site_in_row
                                != str(
                                    selected_site_old
                                ).strip()
                            ):

                                continue

                            # ======================================================
                            # UPDATE O:T
                            # ======================================================

                            edited_relocation_df.at[
                                idx,
                                "Date Reloc."
                            ] = reloc_date

                            edited_relocation_df.at[
                                idx,
                                "No. DO Reloc."
                            ] = new_reloc_do_num

                            edited_relocation_df.at[
                                idx,
                                "Qty Reloc."
                            ] = safe_qty(
                                row.get(
                                    "Qty",
                                    0
                                ),
                                default=0
                            )

                            edited_relocation_df.at[
                                idx,
                                "Site Reloc."
                            ] = selected_site_new

                            edited_relocation_df.at[
                                idx,
                                "Mitra Reloc."
                            ] = reloc_mitra

                            edited_relocation_df.at[
                                idx,
                                "Remarks Reloc."
                            ] = reloc_remarks

                            # ======================================================
                            # NEW RELOCATION ROW
                            # ======================================================

                            new_row = row.copy()

                            new_row["No. DO"] = (
                                new_reloc_do_num
                            )

                            new_row["Delv. Date"] = (
                                reloc_date
                            )

                            new_row["Qty"] = safe_qty(
                                row.get(
                                    "Qty",
                                    0
                                ),
                                default=0
                            )

                            new_row["UoM"] = (
                                get_master_uom(
                                    row.get(
                                        "Material Code",
                                        ""
                                    ),
                                    default=get_uom_from_row(
                                        row,
                                        default="Pcs"
                                    )
                                )
                            )

                            new_row["Charging Type"] = (
                                row.get(
                                    "Charging Type",
                                    ""
                                )
                            )

                            new_row["Site Alocation"] = (
                                selected_site_new
                            )

                            new_row["Remarks"] = (
                                row.get(
                                    "Remarks",
                                    ""
                                )
                            )

                            if new_row["Remarks"] is None:

                                new_row["Remarks"] = ""

                            new_row["To"] = e_to

                            new_row["Phone No."] = (
                                e_contact
                            )

                            new_row["Address"] = (
                                e_address
                            )

                            new_row["EPC"] = e_epc

                            # ======================================================
                            # CLEAR O:T
                            # ======================================================

                            new_row["Date Reloc."] = ""
                            new_row["No. DO Reloc."] = ""
                            new_row["Qty Reloc."] = ""
                            new_row["Site Reloc."] = ""
                            new_row["Mitra Reloc."] = ""
                            new_row["Remarks Reloc."] = ""

                            new_reloc_rows_to_save.append(
                                new_row.to_dict()
                            )

                            relocation_count += 1

                        if relocation_count == 0:

                            st.error(
                                "❌ Site asal tidak ditemukan "
                                "pada material DO yang sedang diedit."
                            )

                        elif not new_reloc_rows_to_save:

                            st.error(
                                "❌ Tidak ada material yang dapat "
                                "dibuat sebagai DO Relokasi."
                            )

                        else:

                            with st.spinner(
                                "Memperbarui DO asal..."
                            ):

                                update_result = (
                                    update_do_in_db_material_out(
                                        e_no_do,
                                        edited_relocation_df
                                        .to_dict(
                                            orient="records"
                                        )
                                    )
                                )

                            if not update_result:

                                st.error(
                                    "❌ Gagal memperbarui "
                                    "DO asal. "
                                    "DO Relokasi baru "
                                    "tidak dibuat."
                                )

                            else:

                                with st.spinner(
                                    "Menyimpan DO Relokasi..."
                                ):

                                    save_reloc_result = (
                                        save_do_to_db_material_out(
                                            new_reloc_rows_to_save
                                        )
                                    )

                                if not save_reloc_result:

                                    st.error(
                                        "⚠️ DO asal berhasil diperbarui, "
                                        "tetapi DO Relokasi gagal disimpan.\n\n"
                                        f"Nomor DO Relokasi: "
                                        f"{new_reloc_do_num}"
                                    )

                                else:

                                    get_used_sites_cached.clear()
                                    fetch_raw_query_data_cached.clear()
                                    load_epc_list.clear()
                                    load_available_relocation_sites.clear()
                                    load_master_item_uom.clear()
                                    load_standard_charging_type_sheet.clear()
                                    load_standard_charging_materials.clear()
                                    load_master_dropdown.clear()

                                    st.session_state.relocation_history.append({

                                        "no_do":
                                            e_no_do,

                                        "timestamp":
                                            reloc_timestamp,

                                        "old_site":
                                            selected_site_old,

                                        "new_site":
                                            selected_site_new,

                                        "reloc_do":
                                            new_reloc_do_num,

                                        "reason":
                                            reloc_remarks

                                    })

                                    st.success(
                                        "✅ Relokasi Berhasil!\n\n"
                                        f"DO Asal: **{e_no_do}**\n\n"
                                        f"Site Asal: **{selected_site_old}**\n\n"
                                        f"Site Baru: **{selected_site_new}**\n\n"
                                        f"DO Relokasi: **{new_reloc_do_num}**"
                                    )

                                    st.rerun()

            # ==========================================================================
            # HISTORY
            # ==========================================================================

            do_hist = [
                h
                for h in st.session_state.relocation_history
                if h["no_do"] == e_no_do
            ]

            if do_hist:

                st.markdown(
                    "#### 📜 Audit Trail / Histori Relokasi DO Ini"
                )

                df_hist = pd.DataFrame(
                    do_hist
                )

                st.dataframe(
                    df_hist[
                        [
                            "timestamp",
                            "old_site",
                            "new_site",
                            "reloc_do",
                            "reason"
                        ]
                    ],
                    use_container_width=True
                )

            st.markdown("---")

            # ==========================================================================
            # MANUAL EDIT
            # ==========================================================================

            col_btn1, col_btn2 = st.columns(2)

            with col_btn1:

                if st.button(
                    "💾 Simpan Perubahan Edit Manual DO",
                    type="primary",
                    key="btn_save_manual_edit"
                ):

                    updated_materials = (
                        edited_mat_df
                        .to_dict(
                            orient="records"
                        )
                    )

                    for row in updated_materials:

                        row["Qty"] = safe_qty(
                            row.get(
                                "Qty",
                                0
                            ),
                            default=0
                        )

                        row["UoM"] = get_master_uom(
                            row.get(
                                "Material Code",
                                ""
                            ),
                            default=get_uom_from_row(
                                row,
                                default="Pcs"
                            )
                        )

                        if row.get(
                            "Remarks"
                        ) is None:

                            row["Remarks"] = ""

                    with st.spinner(
                        "Memperbarui database Google Sheets..."
                    ):

                        update_result = (
                            update_do_in_db_material_out(
                                e_no_do,
                                updated_materials
                            )
                        )

                    if update_result:

                        st.success(
                            f"Berhasil memperbarui "
                            f"{e_no_do} di database!"
                        )

                        get_used_sites_cached.clear()
                        load_master_item_uom.clear()

                        st.session_state.current_do = {

                            "no_do":
                                e_no_do,

                            "date":
                                e_date,

                            "epc":
                                e_epc,

                            "charging_type":
                                edit_data.get(
                                    "charging_type",
                                    ""
                                ),

                            "expedition":
                                edit_data.get(
                                    "expedition",
                                    ""
                                ),

                            "to":
                                e_to,

                            "contact":
                                e_contact,

                            "address":
                                e_address,

                            "materials":
                                updated_materials

                        }

                        st.rerun()

                    else:

                        st.error(
                            "❌ Gagal memperbarui DO."
                        )

            # ==========================================================================
            # PREVIEW
            # ==========================================================================

            with col_btn2:

                if st.button(
                    "🖨️ Set Ke Preview & Cetak PDF Baru",
                    key="btn_set_preview"
                ):

                    preview_materials = (
                        edited_mat_df
                        .to_dict(
                            orient="records"
                        )
                    )

                    for row in preview_materials:

                        row["Qty"] = safe_qty(
                            row.get(
                                "Qty",
                                0
                            ),
                            default=0
                        )

                        row["UoM"] = get_master_uom(
                            row.get(
                                "Material Code",
                                ""
                            ),
                            default=get_uom_from_row(
                                row,
                                default="Pcs"
                            )
                        )

                        if row.get(
                            "Remarks"
                        ) is None:

                            row["Remarks"] = ""

                    preview_sites = set()

                    for x in preview_materials:

                        site_value = x.get(
                            "Site Alocation",
                            ""
                        )

                        if site_value is None:

                            site_value = x.get(
                                "Site Allocation",
                                ""
                            )

                        if site_value:

                            site_value = str(
                                site_value
                            ).strip()

                            if site_value:

                                preview_sites.add(
                                    site_value
                                )

                    st.session_state.current_do = {

                        "no_do":
                            edit_data.get(
                                "no_do",
                                ""
                            ),

                        "date":
                            e_date,

                        "epc":
                            e_epc,

                        "charging_type":
                            edit_data.get(
                                "charging_type",
                                ""
                            ),

                        "expedition":
                            edit_data.get(
                                "expedition",
                                ""
                            ),

                        "to":
                            e_to,

                        "contact":
                            e_contact,

                        "address":
                            e_address,

                        "materials":
                            preview_materials,

                        "site_count":
                            len(preview_sites)

                    }

                    st.success(
                        "Data DO telah diset untuk preview. "
                        "Silakan buka tab "
                        "**🖨️ Preview & PDF Cetak (A5)**."
                    )


# ==============================================================================
# ALIAS
# ==============================================================================

show = render
