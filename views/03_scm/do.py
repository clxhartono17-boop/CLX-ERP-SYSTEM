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
        get_query_sheet_data
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
# MASTER DROPDOWN
# ==============================================================================

@st.cache_data(
    ttl=3600,
    show_spinner=False
)
def load_master_dropdown():

    charging_types = DEFAULT_CHARGING_TYPES.copy()

    expeditions = DEFAULT_EXPEDITIONS.copy()

    return charging_types, expeditions


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
            == target_epc
        )

        &

        (
            working_df["_charging_clean"]
            == target_charging
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
# STANDARD MATERIAL
# ==============================================================================

@st.cache_data(
    ttl=3600,
    show_spinner=False
)
def load_standard_charging_materials(
    charging_type
):

    if not charging_type:

        return []

    raw_data = [

        {
            "code": "AC0001",
            "name": "Clamp Conduit",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 5,
                "12S1P": 5
            }
        },

        {
            "code": "AC0002",
            "name": "Kabel Schoen 16",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 10,
                "12S1P": 10
            }
        },

        {
            "code": "AC0004",
            "name": "Kabel Vynil Biru",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 4,
                "12S1P": 4
            }
        },

        {
            "code": "AC0005",
            "name": "Kabel Vynil Hijau",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 2,
                "12S1P": 2
            }
        },

        {
            "code": "AC0006",
            "name": "Kabel Vynil Hitam",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 2,
                "12S1P": 2
            }
        },

        {
            "code": "AC0008",
            "name": "Kabel Vynil Merah",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 4,
                "12S1P": 4
            }
        },

        {
            "code": "AC0009",
            "name": "Kuku Macan 10",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 1,
                "12S1P": 1
            }
        },

        {
            "code": "AC0010",
            "name": "Sok Konektor Grounding 5/8\"",
            "uom": "Pcs",
            "qty_map": {
                "DC20": 1,
                "DC30": 1,
                "DC60": 1
            }
        },

        {
            "code": "MM0001",
            "name": "APAR 3Kg",
            "uom": "Pcs",
            "qty_map": {
                "DC20": 1,
                "DC30": 1,
                "DC60": 1
            }
        },

        {
            "code": "MM0002",
            "name": "Box APAR",
            "uom": "Pcs",
            "qty_map": {
                "DC20": 1,
                "DC30": 1,
                "DC60": 1
            }
        },

        {
            "code": "MM0003",
            "name": "Combiner 125A",
            "uom": "Unit",
            "qty_map": {
                "DC60": 1
            }
        },

        {
            "code": "MM0004",
            "name": "Combiner 63A BSS",
            "uom": "Pcs",
            "qty_map": {
                "12S1P": 1
            }
        },

        {
            "code": "MM0005",
            "name": "Combiner 40A 3P",
            "uom": "Pcs",
            "qty_map": {
                "DC20": 1
            }
        },

        {
            "code": "MM0006",
            "name": "Combiner 40A BSS",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 1
            }
        },

        {
            "code": "MM0007",
            "name": "Combiner 63A",
            "uom": "Pcs",
            "qty_map": {
                "DC30": 1
            }
        },

        {
            "code": "MM0009",
            "name": "Conduit Anaconda 1\"",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 10,
                "12S1P": 10
            }
        },

        {
            "code": "MM0011",
            "name": "Kabel Grounding 6",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 5,
                "12S1P": 5
            }
        },

        {
            "code": "MM0012",
            "name": "Kabel Power NYY 4x10",
            "uom": "Pcs",
            "qty_map": {
                "DC20": 12
            }
        },

        {
            "code": "MM0013",
            "name": "Kabel Power NYY 4x16",
            "uom": "Pcs",
            "qty_map": {
                "DC30": 12
            }
        },

        {
            "code": "MM0014",
            "name": "Kabel Power NYY 4x25mm",
            "uom": "Pcs",
            "qty_map": {
                "DC60": 12
            }
        },

        {
            "code": "MM0016",
            "name": "Kabel Power NYYHY 3x10",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 10
            }
        },

        {
            "code": "MM0017",
            "name": "NYA 10mm",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 5,
                "12S1P": 5,
                "DC20": 15,
                "DC30": 15
            }
        },

        {
            "code": "MM0018",
            "name": "NYA 16mm",
            "uom": "Pcs",
            "qty_map": {
                "DC60": 15
            }
        },

        {
            "code": "MM0020",
            "name": "Stick Rod 2m",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 1,
                "12S1P": 1,
                "DC20": 1,
                "DC30": 1,
                "DC60": 1
            }
        },

        {
            "code": "MM0021",
            "name": "Stick Rod 1.5m",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 1,
                "12S1P": 1,
                "DC20": 1,
                "DC30": 1,
                "DC60": 1
            }
        },

        {
            "code": "MM0022",
            "name": "Stick Rod 1m",
            "uom": "Pcs",
            "qty_map": {
                "6S1P": 1,
                "12S1P": 1,
                "DC20": 1,
                "DC30": 1,
                "DC60": 1
            }
        },

        {
            "code": "MM0023",
            "name": "Wheel Stopper",
            "uom": "Pcs",
            "qty_map": {
                "DC20": 2,
                "DC30": 2,
                "DC60": 2
            }
        },

        {
            "code": "MM0024",
            "name": "Kabel Power NYY 4x35",
            "uom": "Pcs",
            "qty_map": {
                "DC60": 12
            }
        }

    ]

    result = []

    for item in raw_data:

        std_qty = item["qty_map"].get(
            charging_type,
            0
        )

        if std_qty > 0:

            result.append({

                "code":
                    item["code"],

                "name":
                    item["name"],

                "std_qty":
                    std_qty,

                "uom":
                    item["uom"]

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
# SAFE QTY
# ==============================================================================

def safe_qty(value, default=0):

    """
    Mengubah Qty dari data editor menjadi angka.

    Penting:
    - 0 tetap dianggap 0
    - tidak menggunakan `or default`
      karena 0 adalah nilai valid.
    """

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
# PDF GENERATOR
# ==============================================================================

def generate_do_a5_pdf(data):

    if not REPORTLAB_AVAILABLE:

        raise RuntimeError(
            "ReportLab belum terpasang."
        )

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

    title_style = ParagraphStyle(

        "T",

        fontName="Helvetica-Bold",

        fontSize=10,

        textColor=colors.HexColor("#1a365d")

    )

    subtitle_style = ParagraphStyle(

        "ST",

        fontName="Helvetica",

        fontSize=6,

        textColor=colors.HexColor("#4a5568"),

        leading=7

    )

    body_style = ParagraphStyle(

        "B",

        fontName="Helvetica",

        fontSize=6.5,

        leading=8,

        textColor=colors.HexColor("#2d3748")

    )

    body_bold = ParagraphStyle(

        "BB",

        fontName="Helvetica-Bold",

        fontSize=6.5,

        leading=8,

        textColor=colors.HexColor("#1a365d")

    )

    header_table_style = ParagraphStyle(

        "HT",

        fontName="Helvetica-Bold",

        fontSize=6.5,

        textColor=colors.white,

        alignment=1

    )

    logo_path = "assets/logo.png"

    if os.path.exists(logo_path):

        logo_img = RLImage(
            logo_path,
            width=90,
            height=25
        )

    else:

        logo_img = Paragraph(
            "<b>PT. CLX</b>",
            title_style
        )

    company_info = [

        Paragraph(
            "<b>PT. Connectivity Leads excellence</b>",
            title_style
        ),

        Paragraph(
            "Jl. M Ali 2 No. 19 RT 007 RW 004 Tanah Baru, "
            "Beji, Kota Depok, Jawa barat 16426",
            subtitle_style
        ),

        Paragraph(
            "E: clx.central@gmail.com | T: +62 821-4858-1879",
            subtitle_style
        )

    ]

    head_table = Table(

        [[logo_img, company_info]],

        colWidths=[95, 295]

    )

    head_table.setStyle(

        TableStyle([

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),

            (
                "LINEBELOW",
                (0, 0),
                (-1, -1),
                1,
                colors.HexColor("#1a365d")
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                4
            )

        ])

    )

    elements.append(head_table)

    elements.append(
        Spacer(1, 5)
    )

    # --------------------------------------------------------------------------
    # TO BOX
    # --------------------------------------------------------------------------

    to_box = [

        [
            Paragraph("<b>To</b>", body_bold),
            ""
        ],

        [
            Paragraph("Name:", body_style),
            Paragraph(
                str(data.get("to", "")),
                body_bold
            )
        ],

        [
            Paragraph("Phone No.:", body_style),
            Paragraph(
                str(data.get("contact", "")),
                body_style
            )
        ],

        [
            Paragraph("Address:", body_style),
            Paragraph(
                str(data.get("address", "")),
                body_style
            )
        ]

    ]

    to_table = Table(

        to_box,

        colWidths=[45, 140]

    )

    to_table.setStyle(

        TableStyle([

            (
                "BOX",
                (0, 0),
                (-1, -1),
                0.5,
                colors.HexColor("#cbd5e0")
            ),

            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor("#edf2f7")
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

    # --------------------------------------------------------------------------
    # META BOX
    # --------------------------------------------------------------------------

    meta_box = [

        [

            Paragraph(
                "<b>DELIVERY ORDER</b>",
                ParagraphStyle(
                    "DO",
                    fontName="Helvetica-Bold",
                    fontSize=8,
                    alignment=1,
                    textColor=colors.HexColor("#1a365d")
                )
            ),

            ""

        ],

        [
            Paragraph("No. DO:", body_bold),
            Paragraph(
                str(data.get("no_do", "")),
                body_bold
            )
        ],

        [
            Paragraph("Date:", body_style),
            Paragraph(
                str(data.get("date", "")),
                body_style
            )
        ],

        [
            Paragraph("EPC:", body_style),
            Paragraph(
                str(data.get("epc", "")),
                body_style
            )
        ],

        [
            Paragraph("Charging Type:", body_style),
            Paragraph(
                str(data.get("charging_type", "-")),
                body_style
            )
        ],

        [
            Paragraph("Expedition:", body_style),
            Paragraph(
                str(data.get("expedition", "-")),
                body_style
            )
        ]

    ]

    meta_table = Table(

        meta_box,

        colWidths=[65, 140]

    )

    meta_table.setStyle(

        TableStyle([

            (
                "SPAN",
                (0, 0),
                (1, 0)
            ),

            (
                "BOX",
                (0, 0),
                (-1, -1),
                0.5,
                colors.HexColor("#cbd5e0")
            ),

            (
                "BACKGROUND",
                (0, 0),
                (1, 0),
                colors.HexColor("#e2e8f0")
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

    top_info_table = Table(

        [[to_table, meta_table]],

        colWidths=[190, 200]

    )

    top_info_table.setStyle(

        TableStyle([

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP"
            )

        ])

    )

    elements.append(top_info_table)

    elements.append(
        Spacer(1, 5)
    )

    # --------------------------------------------------------------------------
    # MATERIAL TABLE
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

        # ----------------------------------------------------------------------
        # IMPORTANT:
        # Site Allocation hanya mengambil kolom Site Alocation / Site Allocation.
        # Remarks TIDAK digunakan sebagai fallback.
        # ----------------------------------------------------------------------

        site = (

            item.get("Site Alocation")

            if item.get("Site Alocation") is not None

            else item.get("Site Allocation", "")

        )

        if site is None:

            site = ""

        uom = (

            item.get("UoM")

            if item.get("UoM") is not None

            else item.get("uom", "Pcs")

        )

        if uom is None or str(uom).strip() == "":

            uom = "Pcs"

        qty = safe_qty(
            item.get("Qty", 0),
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

        site = (

            material.get("Site Alocation")

            if material.get("Site Alocation") is not None

            else material.get("Site Allocation", "")

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

        col1, col2, col3 = st.columns(3)

        # ----------------------------------------------------------------------
        # COLUMN 1
        # ----------------------------------------------------------------------

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

        # ----------------------------------------------------------------------
        # COLUMN 2
        # ----------------------------------------------------------------------

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

        # ----------------------------------------------------------------------
        # COLUMN 3
        # ----------------------------------------------------------------------

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

        st.subheader(
            "Filter Site & Kalkulasi Material Automatic"
        )

        # ----------------------------------------------------------------------
        # SITE FILTER
        # ----------------------------------------------------------------------

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

        # ----------------------------------------------------------------------
        # MATERIAL
        # ----------------------------------------------------------------------

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

            total_qty = (

                item["std_qty"]

                *

                (
                    site_count
                    if site_count > 0
                    else 1
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
                    item["uom"],

                "Remarks":
                    ""

            })

        df_materials = pd.DataFrame(
            table_data
        )

        st.subheader(
            "Detail Material Item "
            "(Akan Didistribusikan per Site)"
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
                            "Qty awal = Std Qty x Total Site. "
                            "Qty dapat diedit manual."
                        ),
                        min_value=0,
                        step=1
                    ),

                "UoM":
                    st.column_config.TextColumn(
                        disabled=True
                    ),

                "Remarks":
                    st.column_config.TextColumn(
                        "Remarks"
                    )

            }

        )

        st.divider()

        # ----------------------------------------------------------------------
        # SAVE DO
        # ----------------------------------------------------------------------

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

            else:

                date_str = do_date.strftime(
                    "%Y-%m-%d"
                )

                generated_db_rows = []

                row_counter = 1

                # ==================================================================
                # IMPORTANT FIX:
                #
                # Sebelumnya sistem menggunakan:
                #
                #     mat_item["std_qty"]
                #
                # sehingga Qty yang diedit user di data_editor diabaikan.
                #
                # Sekarang sistem menggunakan edited_df.
                # ==================================================================

                edited_material_rows = (
                    edited_df
                    .to_dict(
                        orient="records"
                    )
                )

                # ------------------------------------------------------------------
                # GENERATE ROW PER SITE
                # ------------------------------------------------------------------

                for site_name in selected_sites:

                    for mat_item in edited_material_rows:

                        material_code = (
                            mat_item.get(
                                "Material Code",
                                ""
                            )
                        )

                        material_name = (
                            mat_item.get(
                                "Material Name",
                                ""
                            )
                        )

                        qty = safe_qty(
                            mat_item.get(
                                "Qty",
                                0
                            ),
                            default=0
                        )

                        uom = (
                            mat_item.get(
                                "UoM",
                                ""
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

                        # ==========================================================
                        # FIX #1:
                        #
                        # Remarks TIDAK lagi diisi site_name.
                        #
                        # Site hanya masuk ke:
                        #     Site Alocation
                        #
                        # Remarks tetap mengambil:
                        #     mat_item["Remarks"]
                        # ==========================================================

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
                            # FIX #2:
                            #
                            # Qty sekarang berasal dari edited_df.
                            #
                            # Jika user mengubah Qty menjadi 0,
                            # maka database menerima 0.
                            # ======================================================

                            "Qty":
                                qty,

                            "UoM":
                                uom,

                            "Charging Type":
                                charging_type,

                            "Site Alocation":
                                site_name,

                            # ======================================================
                            # REMARKS ASLI DARI DATA EDITOR
                            # ======================================================

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

                # ------------------------------------------------------------------
                # SAVE
                # ------------------------------------------------------------------

                with st.spinner(
                    "Menyimpan transaksi ke "
                    "sheet 'DB Material Out'..."
                ):

                    save_result = (
                        save_do_to_db_material_out(
                            generated_db_rows
                        )
                    )

                if not save_result:

                    st.error(
                        "❌ Gagal menyimpan Delivery Order."
                    )

                else:

                    # ----------------------------------------------------------------
                    # CURRENT DO
                    # ----------------------------------------------------------------

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

                    # ----------------------------------------------------------------
                    # SUCCESS NOTIFICATION
                    # ----------------------------------------------------------------

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

                    # ----------------------------------------------------------------
                    # INVALIDATE CACHE
                    # ----------------------------------------------------------------

                    get_used_sites_cached.clear()

                    fetch_raw_query_data_cached.clear()

                    load_epc_list.clear()

                    load_available_relocation_sites.clear()

                    # ----------------------------------------------------------------
                    # RESET DO NUMBER
                    # ----------------------------------------------------------------

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

                pdf_bytes = generate_do_a5_pdf(
                    do_data
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

        # ----------------------------------------------------------------------
        # GET DO LIST
        # ----------------------------------------------------------------------

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

        # ----------------------------------------------------------------------
        # SEARCH
        # ----------------------------------------------------------------------

        if btn_search and selected_do_search:

            with st.spinner(

                f"Mencari data {selected_do_search}..."

            ):

                found_data = get_do_by_number(
                    selected_do_search
                )

            if found_data:

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

        # ----------------------------------------------------------------------
        # EDIT FORM
        # ----------------------------------------------------------------------

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

            # ==================================================================
            # RELOCATION
            # ==================================================================

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

                # --------------------------------------------------------------
                # CURRENT SITE
                # --------------------------------------------------------------

                raw_sites_in_do = []

                # ==============================================================
                # FIX:
                #
                # HANYA membaca Site Alocation.
                #
                # Jangan membaca Remarks sebagai site.
                # ==============================================================

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

                # --------------------------------------------------------------
                # AVAILABLE NEW SITE
                # --------------------------------------------------------------

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

                # --------------------------------------------------------------
                # EXECUTE RELOCATION
                # --------------------------------------------------------------

                if st.button(

                    "🔀 Eksekusi Relokasi Site",

                    type="secondary",

                    key="btn_execute_relocation"

                ):

                    if (
                        selected_site_old
                        == "Tidak Ada Site"
                    ):

                        st.error(
                            "Site asal tidak ditemukan!"
                        )

                    elif (

                        not selected_site_new

                        or

                        selected_site_new
                        == "Tidak ada site baru yang tersedia"

                    ):

                        st.error(
                            "Silakan pilih Site Tujuan Baru yang valid!"
                        )

                    elif (

                        selected_site_old
                        == selected_site_new

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

                        # ------------------------------------------------------
                        # GENERATE RELOCATION DO
                        # ------------------------------------------------------

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

                        # ------------------------------------------------------
                        # FIND SELECTED SITE
                        # ------------------------------------------------------

                        relocation_count = 0

                        for idx, row in (
                            edited_relocation_df.iterrows()
                        ):

                            # ==================================================
                            # FIX:
                            #
                            # Site hanya dari Site Alocation.
                            # Remarks tidak digunakan.
                            # ==================================================

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

                            # --------------------------------------------------
                            # UPDATE O:T
                            # --------------------------------------------------

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

                            # --------------------------------------------------
                            # CREATE NEW RELOCATION ROW
                            # --------------------------------------------------

                            new_row = row.copy()

                            new_row["No. DO"] = (
                                new_reloc_do_num
                            )

                            new_row["Delv. Date"] = (
                                reloc_date
                            )

                            new_row["Material Code"] = (
                                row.get(
                                    "Material Code",
                                    ""
                                )
                            )

                            new_row["Material Name"] = (
                                row.get(
                                    "Material Name",
                                    ""
                                )
                            )

                            # ==================================================
                            # Qty tetap mengambil Qty hasil edit.
                            # Jika 0 -> tetap 0.
                            # ==================================================

                            new_row["Qty"] = safe_qty(
                                row.get(
                                    "Qty",
                                    0
                                ),
                                default=0
                            )

                            new_row["UoM"] = (

                                row.get("UoM")

                                if row.get("UoM") is not None

                                else row.get(
                                    "uom",
                                    "Pcs"
                                )

                            )

                            if (
                                new_row["UoM"] is None
                                or
                                str(
                                    new_row["UoM"]
                                ).strip() == ""
                            ):

                                new_row["UoM"] = "Pcs"

                            new_row["Charging Type"] = (
                                row.get(
                                    "Charging Type",
                                    ""
                                )
                            )

                            # ==================================================
                            # SITE BARU HANYA MASUK KE SITE ALOCATION
                            # ==================================================

                            new_row["Site Alocation"] = (
                                selected_site_new
                            )

                            # ==================================================
                            # FIX:
                            #
                            # Remarks tetap Remarks.
                            #
                            # Tidak lagi:
                            # new_row["Remarks"] = selected_site_new
                            # ==================================================

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

                            # --------------------------------------------------
                            # CLEAR O:T
                            # --------------------------------------------------

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

                        # ------------------------------------------------------
                        # VALIDATION
                        # ------------------------------------------------------

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

                            # --------------------------------------------------
                            # UPDATE OLD DO
                            # --------------------------------------------------

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

                                # ----------------------------------------------
                                # SAVE NEW DO
                                # ----------------------------------------------

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

                                    # ------------------------------------------
                                    # CACHE INVALIDATION
                                    # ------------------------------------------

                                    get_used_sites_cached.clear()

                                    fetch_raw_query_data_cached.clear()

                                    load_epc_list.clear()

                                    load_available_relocation_sites.clear()

                                    # ------------------------------------------
                                    # HISTORY
                                    # ------------------------------------------

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

                                        f"Site Asal: "
                                        f"**{selected_site_old}**\n\n"

                                        f"Site Baru: "
                                        f"**{selected_site_new}**\n\n"

                                        f"DO Relokasi: "
                                        f"**{new_reloc_do_num}**"

                                    )

                                    st.rerun()

            # ==================================================================
            # HISTORY
            # ==================================================================

            do_hist = [

                h

                for h in
                st.session_state.relocation_history

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

            # ==================================================================
            # MANUAL EDIT
            # ==================================================================

            col_btn1, col_btn2 = st.columns(2)

            with col_btn1:

                if st.button(

                    "💾 Simpan Perubahan Edit Manual DO",

                    type="primary",

                    key="btn_save_manual_edit"

                ):

                    # ==========================================================
                    # Pastikan Qty yang diedit benar-benar dikonversi sebagai
                    # angka dan nilai 0 tidak hilang.
                    # ==========================================================

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

                        st.session_state.current_do = {

                            "no_do":
                                e_no_do,

                            "date":
                                e_date,

                            "epc":
                                e_epc,

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

            # ==================================================================
            # PREVIEW
            # ==================================================================

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

                    # ==========================================================
                    # Normalisasi Qty preview
                    # ==========================================================

                    for row in preview_materials:

                        row["Qty"] = safe_qty(
                            row.get(
                                "Qty",
                                0
                            ),
                            default=0
                        )

                        if row.get(
                            "Remarks"
                        ) is None:

                            row["Remarks"] = ""

                    # ==========================================================
                    # SITE COUNT:
                    #
                    # Hanya dari Site Alocation.
                    # Tidak lagi dari Remarks.
                    # ==========================================================

                    preview_sites = set()

                    for x in preview_materials:

                        site_value = (

                            x.get(
                                "Site Alocation",
                                ""
                            )

                            if x.get(
                                "Site Alocation"
                            ) is not None

                            else x.get(
                                "Site Allocation",
                                ""
                            )

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
