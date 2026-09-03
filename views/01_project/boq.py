import datetime
import io
import os
import re
import sys

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
import streamlit as st


# ==============================================================================
# SAFE IMPORT FOR SERVICES & ROOT DIR CONFIGURATION
# ==============================================================================

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, "../../"))

if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

CWD = os.getcwd()

if CWD not in sys.path:
    sys.path.insert(0, CWD)

try:
    from services.gsheet import get_google_sheet_connection
except ModuleNotFoundError:

    def get_google_sheet_connection():
        return None


# ==============================================================================
# TEMPLATE CONFIGURATION
# ==============================================================================

OLD_TEMPLATE_FILENAME = "Template BOQ Vgreen.xlsx"

NEW_TEMPLATE_FILENAME = "New Template BOQ Sept 2026 All Charger.xlsx"

TEMPLATE_OPTIONS = {
    "Template BOQ Vgreen Lama": OLD_TEMPLATE_FILENAME,
    "New Template BOQ Sept 2026": NEW_TEMPLATE_FILENAME,
}

TEMPLATE_DB_VALUES = {
    "Template BOQ Vgreen Lama": "BOQ LAMA",
    "New Template BOQ Sept 2026": "BOQ BARU",
}

DB_TEMPLATE_VALUES = [
    "BOQ LAMA",
    "BOQ BARU",
]


# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================

def parse_price(val):
    """
    Mengubah berbagai format nominal menjadi float.

    Contoh:
        Rp. 1.000.000 -> 1000000
        1,000,000     -> 1000000
        1.000.000     -> 1000000
        -             -> 0
    """

    if pd.isna(val):
        return 0.0

    if isinstance(val, str):

        if val.strip() == "":
            return 0.0

        if val.strip().lower() in [
            "nan",
            "none",
            "-",
        ]:
            return 0.0

    if isinstance(val, (int, float)):
        try:
            return float(val)
        except Exception:
            return 0.0

    s = str(val).strip()

    s = (
        s.replace("Rp.", "")
        .replace("Rp", "")
        .replace("IDR", "")
        .replace("idr", "")
        .strip()
    )

    if s in ["", "-", "nan", "None"]:
        return 0.0

    # Indonesia:
    # 1.234.567,89
    if "," in s and "." in s:

        if s.rfind(",") > s.rfind("."):
            s = (
                s.replace(".", "")
                .replace(",", ".")
            )
        else:
            s = s.replace(",", "")

    elif "," in s:

        parts = s.split(",")

        if len(parts) == 2 and len(parts[1]) <= 2:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")

    elif "." in s:

        # Nominal seperti:
        # 1.000.000
        s = s.replace(".", "")

    try:
        return float(s)

    except Exception:
        return 0.0


def parse_qty_num(val):
    """
    Parse Qty / Volume menjadi numeric.
    """

    if pd.isna(val):
        return 0.0

    if isinstance(val, (int, float)):
        try:
            return float(val)
        except Exception:
            return 0.0

    s = str(val).strip()

    if s.lower() in [
        "",
        "-",
        "nan",
        "none",
    ]:
        return 0.0

    try:

        if "," in s and "." in s:

            if s.rfind(",") > s.rfind("."):

                return float(
                    s.replace(".", "")
                    .replace(",", ".")
                )

            return float(
                s.replace(",", "")
            )

        if "," in s:

            parts = s.split(",")

            if (
                len(parts) == 2
                and len(parts[1]) <= 3
            ):

                return float(
                    s.replace(",", ".")
                )

            return float(
                s.replace(",", "")
            )

        return float(s)

    except Exception:

        match = re.search(
            r"^-?[\d.,]+",
            s,
        )

        if match:

            raw = match.group(0)

            try:

                if "," in raw and "." not in raw:
                    return float(
                        raw.replace(",", ".")
                    )

                if "." in raw and "," not in raw:

                    # Jika ada lebih dari satu titik,
                    # anggap thousands separator.
                    if raw.count(".") > 1:
                        return float(
                            raw.replace(".", "")
                        )

                    # Satu titik bisa decimal.
                    return float(raw)

            except Exception:
                pass

    return 0.0


def format_currency(value):
    """
    Format nominal untuk tampilan Streamlit/PDF.
    """

    value = parse_price(value)

    return (
        f"Rp. {value:,.0f}"
        .replace(",", ".")
    )


def map_to_standard_province(raw_province):

    if not raw_province or pd.isna(raw_province):
        return "JAVA"

    p = str(raw_province).strip().upper()

    java_keywords = [
        "JAVA",
        "JAWA",
        "BANTEN",
        "DKI",
        "JAKARTA",
        "JABODETABEK",
        "YOGYAKARTA",
        "DIY",
        "WEST JAVA",
        "EAST JAVA",
        "CENTRAL JAVA",
    ]

    sumatera_keywords = [
        "SUMATERA",
        "SUMATRA",
        "ACEH",
        "MEDAN",
        "RIU",
        "RIAU",
        "JAMBI",
        "LAMPUNG",
        "BENGKULU",
        "PALEMBANG",
        "PADANG",
        "BABEL",
        "BANGKA",
    ]

    kalimantan_keywords = [
        "KALIMANTAN",
        "BORNEO",
        "PONTIANAK",
        "BANJARMASIN",
        "BALIKPAPAN",
        "SAMARINDA",
    ]

    sulawesi_keywords = [
        "SULAWESI",
        "CELEBES",
        "MAKASSAR",
        "MANADO",
        "PALU",
        "GORONTALO",
        "MAMUJU",
        "POLEWALI",
    ]

    bali_nusa_keywords = [
        "BALI",
        "NUSA",
        "NTB",
        "NTT",
        "LOMBOK",
        "DENPASAR",
        "SUMBAWA",
        "FLORES",
    ]

    if any(k in p for k in java_keywords):
        return "JAVA"

    elif any(k in p for k in sumatera_keywords):
        return "SUMATERA"

    elif any(k in p for k in kalimantan_keywords):
        return "KALIMANTAN"

    elif any(k in p for k in sulawesi_keywords):
        return "SULAWESI"

    elif any(k in p for k in bali_nusa_keywords):
        return "BALI NUSATENGGARA"

    return "JAVA"


def generate_boq_number(sequence_num=1):

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

    return (
        f"{sequence_num:04d}/CLX/BOQ/"
        f"{month_roman}/{now.year}"
    )


# ==============================================================================
# TEMPLATE PATH FINDER
# ==============================================================================

def get_template_path(filename):

    possible_paths = [
        os.path.join(
            ROOT_DIR,
            "assets",
            "templates",
            filename,
        ),
        os.path.join(
            CWD,
            "assets",
            "templates",
            filename,
        ),
        os.path.join(
            os.path.dirname(CURRENT_DIR),
            "assets",
            "templates",
            filename,
        ),
        os.path.join(
            CURRENT_DIR,
            "assets",
            "templates",
            filename,
        ),
    ]

    for path in possible_paths:

        if os.path.exists(path):
            return path

    template_dirs = [
        os.path.join(
            ROOT_DIR,
            "assets",
            "templates",
        ),
        os.path.join(
            CWD,
            "assets",
            "templates",
        ),
        os.path.join(
            os.path.dirname(CURRENT_DIR),
            "assets",
            "templates",
        ),
        os.path.join(
            CURRENT_DIR,
            "assets",
            "templates",
        ),
    ]

    target_lower = filename.strip().lower()

    for directory in template_dirs:

        if not os.path.isdir(directory):
            continue

        try:

            for file_name in os.listdir(directory):

                if (
                    file_name.strip().lower()
                    == target_lower
                ):

                    return os.path.join(
                        directory,
                        file_name,
                    )

        except Exception:
            continue

    return None


# ==============================================================================
# TEMPLATE / CHARGER NORMALIZATION
# ==============================================================================

def normalize_charger_type(charging_type):

    raw = (
        str(charging_type or "")
        .strip()
        .upper()
    )

    charger_map = {

        "20 KW": "DC20",
        "20KW": "DC20",
        "DC20": "DC20",

        "30 KW": "DC30",
        "30KW": "DC30",
        "DC30": "DC30",

        "60 KW": "DC60",
        "60KW": "DC60",
        "DC60": "DC60",

        "120 KW": "DC120",
        "120KW": "DC120",
        "DC120": "DC120",

        "6S1P": "6S1P",
        "12S1P": "12S1P",
        "12S3P": "12S3P",

        "7KW": "7KW",
        "7 KW": "7KW",

        "22KW": "22KW",
        "22 KW": "22KW",
    }

    return charger_map.get(
        raw,
        raw,
    )


def get_new_template_sheet_name(
    charging_type,
    province_str,
):

    charger_key = normalize_charger_type(
        charging_type
    )

    region = map_to_standard_province(
        province_str
    )

    # BSS
    if charger_key in [
        "6S1P",
        "12S1P",
        "12S3P",
    ]:
        return charger_key

    # EVCS
    if charger_key in [
        "DC20",
        "DC30",
        "DC60",
    ]:

        if region == "JAVA":
            return charger_key

        return f"{charger_key} (NON JAVA)"

    return charger_key


def template_db_value(template_name):

    return TEMPLATE_DB_VALUES.get(
        template_name,
        "BOQ LAMA",
    )


def template_name_from_db_value(value):

    value = str(
        value or ""
    ).strip().upper()

    if value == "BOQ BARU":
        return "New Template BOQ Sept 2026"

    return "Template BOQ Vgreen Lama"


# ==============================================================================
# TEMPLATE HEADER DETECTION
# ==============================================================================

def detect_template_header(
    df_raw,
    max_rows=30,
):

    max_check = min(
        len(df_raw),
        max_rows,
    )

    best_idx = None
    best_score = -1

    for idx in range(max_check):

        values = [
            str(v)
            .strip()
            .upper()
            for v in df_raw.iloc[idx].tolist()
        ]

        score = 0

        if "NO" in values:
            score += 5

        if any(
            "ITEM" in v
            for v in values
        ):
            score += 5

        if any(
            "VOLUME" in v
            or "QTY" in v
            or "QUANTITY" in v
            for v in values
        ):
            score += 3

        if any(
            "UNIT PRICE" in v
            for v in values
        ):
            score += 3

        if any(
            "TOTAL PRICE" in v
            for v in values
        ):
            score += 3

        if score > best_score:

            best_score = score
            best_idx = idx

    return best_idx


def find_column_by_keywords(
    header_values,
    keywords,
    exclude_indices=None,
):

    exclude_indices = (
        exclude_indices
        or set()
    )

    for idx, value in enumerate(
        header_values
    ):

        if idx in exclude_indices:
            continue

        text = (
            str(value)
            .strip()
            .upper()
        )

        for keyword in keywords:

            if keyword in text:
                return idx

    return None


def build_new_template_column_mapping(
    df_raw,
    header_row_idx,
):

    header_values = [
        str(v)
        .strip()
        .upper()
        for v in df_raw.iloc[
            header_row_idx
        ].tolist()
    ]

    used = set()

    no_idx = find_column_by_keywords(
        header_values,
        ["NO"],
        used,
    )

    if no_idx is not None:
        used.add(no_idx)

    item_idx = find_column_by_keywords(
        header_values,
        ["ITEM"],
        used,
    )

    if item_idx is not None:
        used.add(item_idx)

    volume_idx = find_column_by_keywords(
        header_values,
        [
            "VOLUME",
            "QTY",
            "QUANTITY",
        ],
        used,
    )

    if volume_idx is not None:
        used.add(volume_idx)

    # Cari SATUAN/UOM secara spesifik terlebih dahulu
    satuan_idx = None

    for idx, value in enumerate(
        header_values
    ):

        if idx in used:
            continue

        text = str(value).strip().upper()

        if text in [
            "SATUAN",
            "UOM",
            "UNIT",
        ]:

            satuan_idx = idx
            break

    if satuan_idx is not None:
        used.add(satuan_idx)

    merk_idx = find_column_by_keywords(
        header_values,
        [
            "MERK",
            "BRAND",
        ],
        used,
    )

    if merk_idx is not None:
        used.add(merk_idx)

    # UNIT PRICE harus dicari sebelum TOTAL
    unit_price_idx = None

    for idx, value in enumerate(
        header_values
    ):

        if idx in used:
            continue

        text = str(value).strip().upper()

        if text == "UNIT PRICE":
            unit_price_idx = idx
            break

    if unit_price_idx is None:

        unit_price_idx = find_column_by_keywords(
            header_values,
            ["UNIT PRICE"],
            used,
        )

    if unit_price_idx is not None:
        used.add(unit_price_idx)

    total_price_idx = None

    for idx, value in enumerate(
        header_values
    ):

        if idx in used:
            continue

        text = str(value).strip().upper()

        if text == "TOTAL PRICE":
            total_price_idx = idx
            break

    if total_price_idx is None:

        total_price_idx = find_column_by_keywords(
            header_values,
            ["TOTAL PRICE"],
            used,
        )

    # ==========================================================================
    # FALLBACK
    # ==========================================================================

    if no_idx is None:
        no_idx = 0

    if item_idx is None:
        item_idx = 1

    if volume_idx is None:

        if df_raw.shape[1] >= 8:
            volume_idx = 3
        else:
            volume_idx = 2

    if satuan_idx is None:

        if df_raw.shape[1] >= 8:
            satuan_idx = 4
        else:
            satuan_idx = 3

    if merk_idx is None:

        if df_raw.shape[1] >= 8:
            merk_idx = 5
        else:
            merk_idx = 4

    if unit_price_idx is None:

        if df_raw.shape[1] >= 8:
            unit_price_idx = 6
        else:
            unit_price_idx = 5

    if total_price_idx is None:

        if df_raw.shape[1] >= 8:
            total_price_idx = 7
        else:
            total_price_idx = 6

    return [
        no_idx,
        item_idx,
        volume_idx,
        satuan_idx,
        merk_idx,
        unit_price_idx,
        total_price_idx,
    ]


# ==============================================================================
# CACHED DATA FETCHING
# ==============================================================================

@st.cache_data(
    ttl=120,
    show_spinner=False,
)
def get_all_saved_boq():

    try:

        sh = get_google_sheet_connection()

        if not sh:
            return []

        worksheet = sh.worksheet(
            "DB BOQ"
        )

        rows = worksheet.get_all_values()

        if len(rows) <= 1:
            return []

        headers = [
            str(h).strip()
            for h in rows[0]
        ]

        data = []

        for idx, r in enumerate(
            rows[1:],
            start=2,
        ):

            if any(r):

                item = {
                    "row_idx": idx
                }

                for h_idx, h in enumerate(
                    headers
                ):

                    item[h] = (
                        r[h_idx]
                        if h_idx < len(r)
                        else ""
                    )

                data.append(item)

        return data

    except Exception as e:

        st.error(
            f"Gagal mengambil data DB BOQ: {e}"
        )

        return []


@st.cache_data(
    ttl=120,
    show_spinner=False,
)
def get_existing_saved_site_charger_pairs():

    saved_boqs = (
        get_all_saved_boq()
    )

    saved_pairs = set()

    for item in saved_boqs:

        site_name = (
            str(
                item.get(
                    "Site Name",
                    "",
                )
            )
            .strip()
            .lower()
        )

        charger_type = (
            str(
                item.get(
                    "Charger Type",
                    "",
                )
            )
            .strip()
            .lower()
        )

        if site_name and charger_type:

            saved_pairs.add(
                (
                    site_name,
                    charger_type,
                )
            )

    return saved_pairs


@st.cache_data(
    ttl=120,
    show_spinner=False,
)
def fetch_query_site_options(
    exclude_saved=True,
):

    site_options = []
    site_data_map = {}

    existing_saved_pairs = (
        get_existing_saved_site_charger_pairs()
        if exclude_saved
        else set()
    )

    try:

        sh = get_google_sheet_connection()

        if sh:

            worksheet = sh.worksheet(
                "Query"
            )

            data_query = (
                worksheet.get_all_values()
            )

            if len(data_query) > 1:

                for row in data_query[1:]:

                    if len(row) <= 5:
                        continue

                    epc_val = (
                        str(row[1]).strip()
                        if len(row) > 1
                        else "-"
                    )

                    charger_val = (
                        str(row[2]).strip()
                        if len(row) > 2
                        else "DC20"
                    )

                    status_val = (
                        str(row[3]).strip()
                        if len(row) > 3
                        else ""
                    )

                    site_val = (
                        str(row[5]).strip()
                        if len(row) > 5
                        else ""
                    )

                    address_val = (
                        str(row[6]).strip()
                        if len(row) > 6
                        else "-"
                    )

                    raw_prov_val = (
                        str(row[8]).strip()
                        if len(row) > 8
                        else ""
                    )

                    status_upper = (
                        status_val.upper()
                    )

                    if (
                        "DROP"
                        in status_upper
                        or "CANCEL"
                        in status_upper
                    ):
                        continue

                    pair_key = (
                        site_val.strip().lower(),
                        charger_val.strip().lower(),
                    )

                    if (
                        exclude_saved
                        and pair_key
                        in existing_saved_pairs
                    ):
                        continue

                    if (
                        site_val
                        and site_val.lower()
                        not in [
                            "nan",
                            "",
                            "none",
                            "project / location name",
                        ]
                    ):

                        std_province = (
                            map_to_standard_province(
                                raw_prov_val
                            )
                        )

                        display_label = (
                            f"{site_val} "
                            f"({charger_val})"
                        )

                        if (
                            display_label
                            not in site_data_map
                        ):

                            site_options.append(
                                display_label
                            )

                        site_data_map[
                            display_label
                        ] = {
                            "site_name": site_val,
                            "epc": (
                                epc_val
                                if epc_val
                                else "-"
                            ),
                            "address": (
                                address_val
                                if address_val
                                else "-"
                            ),
                            "charger": (
                                charger_val
                                if charger_val
                                else "DC20"
                            ),
                            "province": std_province,
                            "raw_province": (
                                raw_prov_val
                                if raw_prov_val
                                else "-"
                            ),
                            "status": status_val,
                        }

    except Exception as e:

        st.sidebar.warning(
            "⚠️ Gagal membaca sheet "
            f"'Query': {e}"
        )

    if not site_options:

        placeholder = (
            "(Semua kombinasi Site & "
            "Charger telah dibuatkan BOQ)"
        )

        site_options = [
            placeholder
        ]

        site_data_map = {
            placeholder: {
                "site_name": "-",
                "epc": "-",
                "address": "-",
                "charger": "-",
                "province": "JAVA",
                "raw_province": "-",
                "status": "-",
            }
        }

    return (
        site_options,
        site_data_map,
    )


# ==============================================================================
# BOQ CALCULATION ENGINE
# ==============================================================================

def calculate_detail_total(
    quantity,
    unit_price,
):
    """
    Kalkulasi TOTAL PRICE detail.

    TOTAL PRICE selalu:
        Quantity / Volume x Unit Price
    """

    qty = parse_qty_num(quantity)
    price = parse_price(unit_price)

    return qty * price


def recalculate_boq_totals(
    df_boq,
):
    """
    ENGINE KALKULASI BOQ FINAL.

    RULE:

    1. DETAIL:
       TOTAL PRICE = Unit/Volume x UNIT PRICE

    2. PARENT:
       Jika Parent mempunyai Volume > 0
       DAN UNIT PRICE > 0:
           TOTAL PRICE = Volume x UNIT PRICE

       Jika tidak:
           TOTAL PRICE = SUM seluruh detail
           sampai Parent berikutnya.

    3. SUB TOTAL:
       SUM seluruh Parent A-H yang tersedia.

    4. VAT:
       11%

    5. GRAND TOTAL:
       Sub Total + VAT
    """

    if (
        df_boq is None
        or df_boq.empty
    ):

        return (
            df_boq,
            0.0,
            0.0,
            0.0,
        )

    df = df_boq.copy()

    required_columns = [
        "NO",
        "Item",
        "Unit/Volume",
        "Satuan/Uom",
        "MERK",
        "UNIT PRICE",
        "TOTAL PRICE",
    ]

    for col in required_columns:

        if col not in df.columns:
            df[col] = ""

    parent_headers = [
        "A",
        "B",
        "C",
        "D",
        "E",
        "F",
        "G",
        "H",
    ]

    # --------------------------------------------------------------------------
    # NORMALISASI UNIT PRICE
    # --------------------------------------------------------------------------

    for idx in df.index:

        df.loc[
            idx,
            "UNIT PRICE",
        ] = parse_price(
            df.loc[
                idx,
                "UNIT PRICE",
            ]
        )

    # --------------------------------------------------------------------------
    # PASS 1 - DETAIL
    # --------------------------------------------------------------------------

    parent_indices = []

    for idx, row in df.iterrows():

        no_val = (
            str(
                row.get(
                    "NO",
                    "",
                )
            )
            .strip()
            .upper()
        )

        if no_val in parent_headers:

            parent_indices.append(idx)

            continue

        qty = parse_qty_num(
            row.get(
                "Unit/Volume",
                0,
            )
        )

        unit_price = parse_price(
            row.get(
                "UNIT PRICE",
                0,
            )
        )

        detail_total = (
            qty * unit_price
        )

        df.loc[
            idx,
            "TOTAL PRICE",
        ] = detail_total

    # --------------------------------------------------------------------------
    # PASS 2 - PARENT
    # --------------------------------------------------------------------------

    for parent_position, parent_idx in enumerate(
        parent_indices
    ):

        parent_row = df.loc[
            parent_idx
        ]

        parent_qty = parse_qty_num(
            parent_row.get(
                "Unit/Volume",
                0,
            )
        )

        parent_unit_price = parse_price(
            parent_row.get(
                "UNIT PRICE",
                0,
            )
        )

        if (
            parent_position
            + 1
            < len(parent_indices)
        ):

            next_parent_idx = (
                parent_indices[
                    parent_position + 1
                ]
            )

            section_detail_indices = (
                df.index[
                    (
                        df.index
                        > parent_idx
                    )
                    & (
                        df.index
                        < next_parent_idx
                    )
                ]
            )

        else:

            section_detail_indices = (
                df.index[
                    df.index
                    > parent_idx
                ]
            )

        if (
            parent_qty != 0
            and parent_unit_price != 0
        ):

            parent_total = (
                parent_qty
                * parent_unit_price
            )

        else:

            parent_total = 0.0

            for detail_idx in (
                section_detail_indices
            ):

                detail_no = (
                    str(
                        df.loc[
                            detail_idx,
                            "NO",
                        ]
                    )
                    .strip()
                    .upper()
                )

                if detail_no in parent_headers:
                    continue

                parent_total += parse_price(
                    df.loc[
                        detail_idx,
                        "TOTAL PRICE",
                    ]
                )

        df.loc[
            parent_idx,
            "TOTAL PRICE",
        ] = parent_total

    # --------------------------------------------------------------------------
    # SUB TOTAL
    # --------------------------------------------------------------------------

    sub_total = 0.0

    for idx in parent_indices:

        sub_total += parse_price(
            df.loc[
                idx,
                "TOTAL PRICE",
            ]
        )

    # --------------------------------------------------------------------------
    # VAT
    # --------------------------------------------------------------------------

    vat_amount = (
        sub_total * 0.11
    )

    # --------------------------------------------------------------------------
    # GRAND TOTAL
    # --------------------------------------------------------------------------

    grand_total = (
        sub_total
        + vat_amount
    )

    return (
        df,
        sub_total,
        vat_amount,
        grand_total,
    )


# ==============================================================================
# BOQ LOADER
# ==============================================================================

@st.cache_data(
    ttl=600,
    show_spinner=False,
)
def load_boq_dataframe(
    charging_type,
    province_str,
    template_name="Template BOQ Vgreen Lama",
):

    region_normalized = (
        map_to_standard_province(
            province_str
        )
    )

    raw_charging_key = (
        str(
            charging_type or ""
        )
        .upper()
        .strip()
    )

    normalized_charger = (
        normalize_charger_type(
            raw_charging_key
        )
    )

    # ==========================================================================
    # VALIDASI TEMPLATE
    # ==========================================================================

    if (
        template_name
        not in TEMPLATE_OPTIONS
    ):

        template_name = (
            "Template BOQ Vgreen Lama"
        )

    template_filename = (
        TEMPLATE_OPTIONS[
            template_name
        ]
    )

    excel_path = get_template_path(
        template_filename
    )

    if not excel_path:

        st.error(
            "⚠️ File template Excel "
            "tidak ditemukan!\n\n"
            f"Template yang dipilih: "
            f"`{template_name}`\n"
            f"Nama file: "
            f"`{template_filename}`"
        )

        return (
            None,
            normalized_charger,
            region_normalized,
        )

    # ==========================================================================
    # TARGET SHEET
    # ==========================================================================

    if (
        template_name
        == "Template BOQ Vgreen Lama"
    ):

        sheet_map = {

            "20 KW": "DC20",
            "20KW": "DC20",
            "DC20": "DC20",

            "30 KW": "DC30",
            "30KW": "DC30",
            "DC30": "DC30",

            "60 KW": "DC60",
            "60KW": "DC60",
            "DC60": "DC60",

            "120 KW": "DC120",
            "120KW": "DC120",
            "DC120": "DC120",

            "6S1P": "6S1P",
            "12S1P": "12S1P",
            "12S3P": "12S3P",

            "7KW": "7KW",
            "7 KW": "7KW",

            "22KW": "22KW",
            "22 KW": "22KW",
        }

        target_sheet = sheet_map.get(
            raw_charging_key,
            normalized_charger,
        )

        region_col_indices = {
            "JAVA": 0,
            "SUMATERA": 8,
            "BALI NUSATENGGARA": 16,
            "BALI NUSA TENGGARA": 16,
            "KALIMANTAN": 25,
            "SULAWESI": 33,
        }

        start_idx = (
            region_col_indices.get(
                region_normalized,
                0,
            )
        )

    else:

        target_sheet = (
            get_new_template_sheet_name(
                normalized_charger,
                region_normalized,
            )
        )

        start_idx = 0

    # ==========================================================================
    # READ EXCEL
    # ==========================================================================

    try:

        xls = pd.ExcelFile(
            excel_path
        )

        sheet_found = next(
            (
                s
                for s in xls.sheet_names
                if s.strip().lower()
                == target_sheet.strip().lower()
            ),
            None,
        )

        if not sheet_found:

            st.error(
                f"⚠️ Sheet `{target_sheet}` "
                "tidak ditemukan di file "
                f"`{template_filename}`."
            )

            st.info(
                "Sheet yang tersedia:\n\n"
                + ", ".join(
                    xls.sheet_names
                )
            )

            return (
                None,
                target_sheet,
                region_normalized,
            )

        df_raw = pd.read_excel(
            xls,
            sheet_name=sheet_found,
            header=None,
        )

        # ==========================================================================
        # DETECT HEADER
        # ==========================================================================

        header_row_idx = (
            detect_template_header(
                df_raw
            )
        )

        if header_row_idx is None:
            header_row_idx = 5

        data_start_idx = (
            header_row_idx + 1
        )

        # ==========================================================================
        # COLUMN MAPPING
        # ==========================================================================

        if (
            template_name
            == "Template BOQ Vgreen Lama"
        ):

            if (
                start_idx + 6
                >= df_raw.shape[1]
            ):

                start_idx = 0

            col_indices = [
                start_idx + i
                for i in range(7)
            ]

        else:

            col_indices = (
                build_new_template_column_mapping(
                    df_raw,
                    header_row_idx,
                )
            )

        if any(
            i >= df_raw.shape[1]
            for i in col_indices
        ):

            st.error(
                "⚠️ Struktur kolom template "
                f"`{sheet_found}` tidak sesuai."
            )

            return (
                None,
                target_sheet,
                region_normalized,
            )

        # ==========================================================================
        # EXTRACT
        # ==========================================================================

        df_boq = df_raw.iloc[
            data_start_idx:,
            col_indices,
        ].copy()

        df_boq.columns = [
            "NO",
            "Item",
            "Unit/Volume",
            "Satuan/Uom",
            "MERK",
            "UNIT PRICE",
            "TOTAL PRICE",
        ]

        # ==========================================================================
        # CLEAN ITEM
        # ==========================================================================

        df_boq["Item"] = (
            df_boq["Item"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        df_boq = df_boq[
            (
                df_boq["Item"]
                != ""
            )
            &
            (
                df_boq["Item"]
                .str.lower()
                != "nan"
            )
        ].reset_index(
            drop=True
        )

        # ==========================================================================
        # CLEAN PRICE
        # ==========================================================================

        for c in [
            "UNIT PRICE",
            "TOTAL PRICE",
        ]:

            df_boq[c] = (
                df_boq[c]
                .apply(parse_price)
            )

        # ==========================================================================
        # CLEAN OTHER
        # ==========================================================================

        for c in [
            "NO",
            "Unit/Volume",
            "Satuan/Uom",
            "MERK",
        ]:

            df_boq[c] = (
                df_boq[c]
                .fillna("-")
                .astype(str)
                .str.strip()
            )

            df_boq[c] = (
                df_boq[c]
                .replace(
                    {
                        "nan": "-",
                        "NaN": "-",
                        "None": "-",
                    }
                )
            )

        # ==========================================================================
        # INITIAL AUTO CALCULATION
        # ==========================================================================

        (
            df_boq,
            _sub_total,
            _vat,
            _grand_total,
        ) = recalculate_boq_totals(
            df_boq
        )

        return (
            df_boq,
            target_sheet,
            region_normalized,
        )

    except Exception as e:

        st.error(
            "Gagal membaca Excel Template: "
            f"{e}"
        )

        return (
            None,
            target_sheet,
            region_normalized,
        )


# ==============================================================================
# EDITABLE SECTION DETECTION
# ==============================================================================

def get_section_indices(
    df_boq,
    section_no,
):

    if (
        df_boq is None
        or df_boq.empty
    ):
        return []

    parent_headers = [
        "A",
        "B",
        "C",
        "D",
        "E",
        "F",
        "G",
        "H",
    ]

    section_no = (
        str(section_no)
        .strip()
        .upper()
    )

    indices = []

    active = False

    for idx, row in df_boq.iterrows():

        no_val = (
            str(
                row.get(
                    "NO",
                    "",
                )
            )
            .strip()
            .upper()
        )

        if no_val == section_no:

            active = True
            indices.append(idx)

            continue

        if (
            active
            and no_val in parent_headers
        ):

            break

        if active:
            indices.append(idx)

    return indices


def get_editable_sections(
    charger_type,
):

    charger_key = normalize_charger_type(
        charger_type
    )

    # ==========================================================================
    # BSS
    # ==========================================================================

    if charger_key in [
        "6S1P",
        "12S1P",
        "12S3P",
    ]:

        return [
            "A",
            "D",
        ]

    # ==========================================================================
    # EVCS
    # ==========================================================================

    if charger_key in [
        "DC20",
        "DC30",
        "DC60",
    ]:

        return [
            "A",
            "G",
        ]

    return []


# ==============================================================================
# EDIT BOQ SECTIONS
# ==============================================================================

def edit_boq_sections(
    df_boq,
    charger_type,
    widget_prefix,
):

    if (
        df_boq is None
        or df_boq.empty
    ):
        return df_boq

    editable_sections = (
        get_editable_sections(
            charger_type
        )
    )

    if not editable_sections:
        return df_boq

    df = df_boq.copy()

    st.markdown(
        "### ✏️ Edit Detail BOQ"
    )

    st.info(
        "💡 Ubah **Item, Volume, Satuan, "
        "MERK, atau UNIT PRICE**. "
        "**TOTAL PRICE otomatis dihitung "
        "= Volume × UNIT PRICE**."
    )

    for section_no in editable_sections:

        section_indices = (
            get_section_indices(
                df,
                section_no,
            )
        )

        if not section_indices:
            continue

        section_df = (
            df.loc[
                section_indices
            ].copy()
        )

        st.markdown(
            f"#### Point {section_no}"
        )

        editor_columns = [
            "NO",
            "Item",
            "Unit/Volume",
            "Satuan/Uom",
            "MERK",
            "UNIT PRICE",
            "TOTAL PRICE",
        ]

        editor_df = (
            section_df[
                editor_columns
            ].copy()
        )

        # ----------------------------------------------------------------------
        # COLUMN CONFIG
        #
        # NO dan TOTAL PRICE readonly.
        # Semua bagian lainnya editable.
        # ----------------------------------------------------------------------

        column_config = {

            "NO": st.column_config.TextColumn(
                "NO",
                disabled=True,
            ),

            "Item": st.column_config.TextColumn(
                "Item",
                disabled=False,
            ),

            "Unit/Volume": st.column_config.TextColumn(
                "Volume",
                disabled=False,
            ),

            "Satuan/Uom": st.column_config.TextColumn(
                "Satuan",
                disabled=False,
            ),

            "MERK": st.column_config.TextColumn(
                "MERK",
                disabled=False,
            ),

            "UNIT PRICE": st.column_config.NumberColumn(
                "UNIT PRICE",
                min_value=0.0,
                step=1.0,
                format="%.0f",
                disabled=False,
            ),

            "TOTAL PRICE": st.column_config.NumberColumn(
                "TOTAL PRICE",
                format="%.0f",
                disabled=True,
            ),
        }

        edited_section = st.data_editor(
            editor_df,
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            column_config=column_config,
            key=(
                f"{widget_prefix}_"
                f"section_{section_no}"
            ),
        )

        # ----------------------------------------------------------------------
        # COPY EDIT HASIL USER
        # ----------------------------------------------------------------------

        for position, idx in enumerate(
            section_indices
        ):

            if position >= len(
                edited_section
            ):
                continue

            edited_row = (
                edited_section.iloc[
                    position
                ]
            )

            for col in [
                "Item",
                "Unit/Volume",
                "Satuan/Uom",
                "MERK",
                "UNIT PRICE",
            ]:

                df.loc[
                    idx,
                    col,
                ] = edited_row[
                    col
                ]

            # ------------------------------------------------------------------
            # DETAIL TOTAL
            # ------------------------------------------------------------------

            no_val = (
                str(
                    df.loc[
                        idx,
                        "NO",
                    ]
                )
                .strip()
                .upper()
            )

            if no_val not in [
                "A",
                "B",
                "C",
                "D",
                "E",
                "F",
                "G",
                "H",
            ]:

                df.loc[
                    idx,
                    "TOTAL PRICE",
                ] = calculate_detail_total(
                    df.loc[
                        idx,
                        "Unit/Volume",
                    ],
                    df.loc[
                        idx,
                        "UNIT PRICE",
                    ],
                )

    # --------------------------------------------------------------------------
    # FINAL RECALCULATION
    # --------------------------------------------------------------------------

    (
        df,
        _sub_total,
        _vat,
        _grand_total,
    ) = recalculate_boq_totals(
        df
    )

    return df


# ==============================================================================
# DATABASE BOQ
# ==============================================================================

def ensure_db_boq_headers(
    worksheet,
):

    rows = worksheet.get_all_values()

    headers = (
        rows[0]
        if rows
        else []
    )

    required_headers = [
        "No",
        "BOQ No.",
        "Site Name",
        "Charger Type",
        "BOQ Amount Exc. PPN",
        "BOQ Amount inc. PPN",
        "EPC Name",
        "Template",
    ]

    if not headers:

        worksheet.update(
            "A1:H1",
            [required_headers],
        )

        return required_headers

    current_headers = [
        str(h).strip()
        for h in headers
    ]

    if len(current_headers) < 8:

        worksheet.update_cell(
            1,
            8,
            "Template",
        )

    elif (
        current_headers[7]
        != "Template"
    ):

        worksheet.update_cell(
            1,
            8,
            "Template",
        )

    return required_headers


def save_to_db_boq(
    site_name,
    charger_capacity,
    sub_total,
    grand_total,
    epc_name="-",
    template_name="Template BOQ Vgreen Lama",
):

    try:

        sh = get_google_sheet_connection()

        if not sh:
            return None

        try:

            worksheet = sh.worksheet(
                "DB BOQ"
            )

        except Exception:

            worksheet = sh.add_worksheet(
                title="DB BOQ",
                rows=1000,
                cols=10,
            )

        ensure_db_boq_headers(
            worksheet
        )

        existing_rows = (
            worksheet.get_all_values()
        )

        no_urut = (
            len(
                [
                    r
                    for r in existing_rows[1:]
                    if any(r)
                ]
            )
            + 1
            if len(existing_rows) > 1
            else 1
        )

        boq_no = generate_boq_number(
            sequence_num=no_urut
        )

        db_template = (
            template_db_value(
                template_name
            )
        )

        new_row = [
            no_urut,
            boq_no,
            str(site_name),
            str(charger_capacity),
            sub_total,
            grand_total,
            str(epc_name),
            db_template,
        ]

        worksheet.append_row(
            new_row
        )

        return boq_no

    except Exception as e:

        st.error(
            "❌ Gagal menyimpan ke DB BOQ: "
            f"{e}"
        )

        return None


def update_db_boq_row(
    row_idx,
    old_site_name,
    new_site_name,
    charger_capacity,
    sub_total,
    grand_total,
    epc_name,
    template_name=None,
):

    try:

        sh = get_google_sheet_connection()

        if not sh:
            return False

        worksheet = sh.worksheet(
            "DB BOQ"
        )

        ensure_db_boq_headers(
            worksheet
        )

        worksheet.update_cell(
            row_idx,
            3,
            new_site_name,
        )

        worksheet.update_cell(
            row_idx,
            4,
            charger_capacity,
        )

        worksheet.update_cell(
            row_idx,
            5,
            sub_total,
        )

        worksheet.update_cell(
            row_idx,
            6,
            grand_total,
        )

        worksheet.update_cell(
            row_idx,
            7,
            epc_name,
        )

        if template_name:

            worksheet.update_cell(
                row_idx,
                8,
                template_db_value(
                    template_name
                ),
            )

        update_google_sheet_summary(
            old_site_name,
            new_site_name,
            sub_total,
            grand_total,
        )

        return True

    except Exception as e:

        st.error(
            "❌ Gagal memperbarui DB BOQ: "
            f"{e}"
        )

        return False


def update_google_sheet_summary(
    old_site_name,
    new_site_name,
    sub_total,
    grand_total,
):

    try:

        sh = get_google_sheet_connection()

        if not sh:
            return False

        sheet_names = [
            ws.title
            for ws in sh.worksheets()
        ]

        if (
            "Sum Project"
            not in sheet_names
        ):
            return False

        worksheet = sh.worksheet(
            "Sum Project"
        )

        data_sum = (
            worksheet.get_all_values()
        )

        for idx, row in enumerate(
            data_sum[1:],
            start=2,
        ):

            if (
                len(row) > 2
                and str(row[2])
                .strip()
                .lower()
                ==
                str(old_site_name)
                .strip()
                .lower()
            ):

                worksheet.update_cell(
                    idx,
                    3,
                    new_site_name,
                )

                if len(row) >= 19:

                    worksheet.update_cell(
                        idx,
                        18,
                        sub_total,
                    )

                    worksheet.update_cell(
                        idx,
                        19,
                        grand_total,
                    )

                return True

    except Exception as e:

        st.caption(
            "ℹ️ Info: Sheet 'Sum Project' "
            f"belum ter-update ({e})"
        )

    return False


# ==============================================================================
# PDF GENERATOR
# ==============================================================================

def generate_boq_pdf(
    site_name,
    site_location,
    charger_capacity,
    region,
    df_boq,
    sub_total,
    vat,
    grand_total,
):

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=10,
        leftMargin=10,
        topMargin=10,
        bottomMargin=10,
    )

    elements = []

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Heading1"],
        fontSize=9,
        leading=10,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor(
            "#111111"
        ),
        spaceAfter=1,
    )

    sub_style = ParagraphStyle(
        "DocSub",
        parent=styles["Normal"],
        fontSize=7,
        leading=8,
        fontName="Helvetica",
        textColor=colors.HexColor(
            "#333333"
        ),
        spaceAfter=0,
    )

    table_text = ParagraphStyle(
        "TableText",
        parent=styles["Normal"],
        fontSize=5.5,
        leading=6.5,
        fontName="Helvetica",
    )

    table_text_bold = ParagraphStyle(
        "TableTextBold",
        parent=table_text,
        fontName="Helvetica-Bold",
    )

    table_header = ParagraphStyle(
        "TableHeader",
        parent=styles["Normal"],
        fontSize=6,
        leading=7.5,
        fontName="Helvetica-Bold",
        textColor=colors.white,
    )

    elements.append(
        Paragraph(
            f"CHARGING WORK {charger_capacity}",
            title_style,
        )
    )

    elements.append(
        Paragraph(
            f"<b>Site Name:</b> "
            f"{site_name}",
            sub_style,
        )
    )

    elements.append(
        Paragraph(
            f"<b>Site Location:</b> "
            f"{site_location}",
            sub_style,
        )
    )

    elements.append(
        Paragraph(
            f"<b>NEW PLAN BOQ VGREEN - "
            f"{region} ISLAND</b>",
            ParagraphStyle(
                "SubHeader",
                parent=title_style,
                fontSize=7.5,
                leading=8.5,
                spaceAfter=2,
                spaceBefore=1,
            ),
        )
    )

    table_data = [
        [
            Paragraph(
                "NO",
                table_header,
            ),
            Paragraph(
                "Item",
                table_header,
            ),
            Paragraph(
                "Unit/Vol",
                table_header,
            ),
            Paragraph(
                "Satuan",
                table_header,
            ),
            Paragraph(
                "MERK",
                table_header,
            ),
            Paragraph(
                "UNIT PRICE",
                table_header,
            ),
            Paragraph(
                "TOTAL PRICE",
                table_header,
            ),
        ]
    ]

    parent_headers = [
        "A",
        "B",
        "C",
        "D",
        "E",
        "F",
        "G",
        "H",
    ]

    # Pastikan PDF menggunakan kalkulasi terbaru
    (
        pdf_df,
        _,
        _,
        _,
    ) = recalculate_boq_totals(
        df_boq
    )

    for _, row in pdf_df.iterrows():

        no_str = (
            str(
                row.get(
                    "NO",
                    "",
                )
            )
            .strip()
            .upper()
        )

        is_parent = (
            no_str
            in parent_headers
        )

        up_val = parse_price(
            row.get(
                "UNIT PRICE",
                0,
            )
        )

        tp_val = parse_price(
            row.get(
                "TOTAL PRICE",
                0,
            )
        )

        up_str = (
            format_currency(
                up_val
            )
            if up_val != 0
            else (
                "-"
                if not is_parent
                else ""
            )
        )

        tp_str = (
            format_currency(
                tp_val
            )
            if tp_val != 0
            else "-"
        )

        style_to_use = (
            table_text_bold
            if is_parent
            else table_text
        )

        table_data.append(
            [
                Paragraph(
                    no_str,
                    style_to_use,
                ),

                Paragraph(
                    str(
                        row.get(
                            "Item",
                            "",
                        )
                    ),
                    style_to_use,
                ),

                Paragraph(
                    str(
                        row.get(
                            "Unit/Volume",
                            "",
                        )
                    ),
                    style_to_use,
                ),

                Paragraph(
                    str(
                        row.get(
                            "Satuan/Uom",
                            "",
                        )
                    ),
                    style_to_use,
                ),

                Paragraph(
                    str(
                        row.get(
                            "MERK",
                            "",
                        )
                    ),
                    style_to_use,
                ),

                Paragraph(
                    up_str,
                    style_to_use,
                ),

                Paragraph(
                    tp_str,
                    style_to_use,
                ),
            ]
        )

    table_data.append(
        [
            "",
            "",
            Paragraph(
                "<b>Sub Total:</b>",
                table_text_bold,
            ),
            "",
            "",
            "",
            Paragraph(
                f"<b>{format_currency(sub_total)}</b>",
                table_text_bold,
            ),
        ]
    )

    table_data.append(
        [
            "",
            "",
            Paragraph(
                "<b>VAT 11%</b>",
                table_text_bold,
            ),
            "",
            "",
            "",
            Paragraph(
                f"<b>{format_currency(vat)}</b>",
                table_text_bold,
            ),
        ]
    )

    table_data.append(
        [
            "",
            "",
            Paragraph(
                "<b>Total Contractor Price</b>",
                table_text_bold,
            ),
            "",
            "",
            "",
            Paragraph(
                f"<b>{format_currency(grand_total)}</b>",
                table_text_bold,
            ),
        ]
    )

    col_widths = [
        18,
        260,
        45,
        45,
        67,
        70,
        70,
    ]

    t = Table(
        table_data,
        colWidths=col_widths,
    )

    t.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor(
                        "#2C3E50"
                    ),
                ),

                (
                    "ALIGN",
                    (0, 0),
                    (-1, -1),
                    "LEFT",
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
                    (-1, -4),
                    0.3,
                    colors.HexColor(
                        "#CCCCCC"
                    ),
                ),

                (
                    "BACKGROUND",
                    (0, -3),
                    (-1, -1),
                    colors.HexColor(
                        "#F8F9F9"
                    ),
                ),

                (
                    "LINEABOVE",
                    (0, -3),
                    (-1, -3),
                    0.8,
                    colors.HexColor(
                        "#2C3E50"
                    ),
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    0.5,
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    0.5,
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    2,
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    2,
                ),
            ]
        )
    )

    elements.append(t)

    doc.build(elements)

    buffer.seek(0)

    return buffer.getvalue()


# ==============================================================================
# MAIN RENDER
# ==============================================================================

def render():

    st.title(
        "📝 Quotation & BOQ Manager"
    )

    tab_create, tab_edit = st.tabs(
        [
            "➕ Buat BOQ Baru",
            "✏️ Edit / Reuse / Re-Download BOQ",
        ]
    )

    # ==========================================================================
    # FETCH SITE
    # ==========================================================================

    site_options, site_data_map = (
        fetch_query_site_options(
            exclude_saved=True
        )
    )

    PLACEHOLDER_OPTION = (
        "-- Pilih Site Name & Charger --"
    )

    dropdown_options = [
        PLACEHOLDER_OPTION
    ] + sorted(
        list(site_options)
    )

    # ==========================================================================
    # TAB CREATE
    # ==========================================================================

    with tab_create:

        st.subheader(
            "1. Informasi Site & Spesifikasi"
        )

        # ----------------------------------------------------------------------
        # LAST SAVED
        # ----------------------------------------------------------------------

        if (
            "last_saved_info"
            in st.session_state
            and st.session_state[
                "last_saved_info"
            ]
        ):

            last_info = (
                st.session_state[
                    "last_saved_info"
                ]
            )

            st.success(
                f"🎉 **BOQ Berhasil Disimpan!** "
                f"Nomor BOQ: "
                f"`{last_info['boq_no']}` "
                f"untuk "
                f"**{last_info['site_name']}**"
            )

            st.download_button(
                label=(
                    "📥 Download PDF BOQ "
                    "Terakhir Disimpan "
                    f"({last_info['site_name']})"
                ),
                data=last_info[
                    "pdf_bytes"
                ],
                file_name=last_info[
                    "filename"
                ],
                mime="application/pdf",
                key="btn_download_last_saved",
            )

            st.markdown("---")

        # ----------------------------------------------------------------------
        # TEMPLATE
        # ----------------------------------------------------------------------

        st.markdown(
            "### 📑 Pilih Template BOQ"
        )

        template_name = st.selectbox(
            "Template yang digunakan",
            list(
                TEMPLATE_OPTIONS.keys()
            ),
            index=1,
            key="create_boq_template",
        )

        if (
            template_name
            == "New Template BOQ Sept 2026"
        ):

            st.info(
                "🆕 Template September 2026 aktif. "
                "EVCS DC20/DC30/DC60 akan otomatis "
                "memilih sheet JAVA atau NON JAVA "
                "berdasarkan Province."
            )

        else:

            st.caption(
                "Template lama menggunakan "
                "struktur sheet dan mapping wilayah "
                "seperti sebelumnya."
            )

        # ----------------------------------------------------------------------
        # SITE
        # ----------------------------------------------------------------------

        col_s1, col_s2 = st.columns(2)

        with col_s1:

            selected_label = st.selectbox(
                "Pilih Site Name & Charger "
                "(Filtered: Drop/Cancel & "
                "Saved BOQ Pair Excluded)",
                dropdown_options,
                index=0,
                key="create_site_selector",
            )

        if (
            selected_label
            == PLACEHOLDER_OPTION
        ):

            st.info(
                "💡 Silakan pilih "
                "**Site Name & Charger** "
                "dari dropdown di atas."
            )

        else:

            current_meta = (
                site_data_map.get(
                    selected_label,
                    {
                        "site_name": selected_label,
                        "epc": "-",
                        "address": "-",
                        "charger": "-",
                        "province": "JAVA",
                        "raw_province": "-",
                    },
                )
            )

            selected_site = (
                current_meta.get(
                    "site_name",
                    selected_label,
                )
            )

            with col_s2:

                site_address = st.text_input(
                    "Address (Kolom G)",
                    value=current_meta[
                        "address"
                    ],
                    key=(
                        f"address_"
                        f"{selected_label}"
                    ),
                )

            col_s3, col_s4, col_s5 = (
                st.columns(
                    [1.5, 1.5, 1]
                )
            )

            with col_s3:

                charging_type = st.text_input(
                    "Charging Type (Kolom C)",
                    value=current_meta[
                        "charger"
                    ],
                    key=(
                        f"charger_"
                        f"{selected_label}"
                    ),
                )

            with col_s4:

                province = st.text_input(
                    "Province Standar "
                    "(Kolom I Mapped)",
                    value=current_meta[
                        "province"
                    ],
                    key=(
                        f"province_"
                        f"{selected_label}"
                    ),
                )

                st.caption(
                    "📍 Raw Province Sheet: "
                    f"`{current_meta.get('raw_province', '-')}`"
                )

            with col_s5:

                epc_name = st.text_input(
                    "EPC Name (Kolom B)",
                    value=current_meta.get(
                        "epc",
                        "-",
                    ),
                    key=(
                        f"epc_"
                        f"{selected_label}"
                    ),
                )

            saved_pairs = (
                get_existing_saved_site_charger_pairs()
            )

            current_pair_key = (
                selected_site.strip().lower(),
                charging_type.strip().lower(),
            )

            is_already_saved = (
                current_pair_key
                in saved_pairs
                or selected_site == "-"
            )

            if (
                is_already_saved
                and selected_site != "-"
            ):

                st.warning(
                    f"⚠️ Kombinasi Site "
                    f"`{selected_site}` dengan "
                    f"Charger `{charging_type}` "
                    "sudah pernah dibuatkan BOQ."
                )

            # ------------------------------------------------------------------
            # LOAD TEMPLATE
            # ------------------------------------------------------------------

            (
                df_boq,
                target_sheet,
                region_normalized,
            ) = load_boq_dataframe(
                charging_type,
                province,
                template_name,
            )

            if (
                df_boq is not None
                and not df_boq.empty
            ):

                st.subheader(
                    f"2. Table BOQ "
                    f"({target_sheet} - "
                    f"{region_normalized})"
                )

                st.caption(
                    f"📑 Template aktif: "
                    f"**{template_name}**"
                )

                # ==================================================================
                # TEMPLATE BOQ LAMA
                #
                # LOGIKA LAMA TETAP DIPERTAHANKAN.
                #
                # Hanya CABLING AND ACCESSORIES
                # INSTALLATION Point 1-3 yang editable.
                # ==================================================================

                if (
                    template_name
                    == "Template BOQ Vgreen Lama"
                ):

                    st.info(
                        "💡 Khusus "
                        "**CABLING AND ACCESSORIES "
                        "INSTALLATION** "
                        "(Poin 1-3), Qty/Volume "
                        "dapat diubah. "
                        "TOTAL PRICE akan otomatis "
                        "mengikuti Qty × UNIT PRICE."
                    )

                    e_start = False
                    editable_indices = []

                    for idx, row in df_boq.iterrows():

                        no_val = (
                            str(
                                row.get(
                                    "NO",
                                    "",
                                )
                            )
                            .strip()
                            .upper()
                        )

                        item_text = (
                            str(
                                row.get(
                                    "Item",
                                    "",
                                )
                            ).upper()
                        )

                        if (
                            "CABLING AND ACCESSORIES"
                            in item_text
                            or no_val == "E"
                        ):

                            e_start = True
                            continue

                        elif no_val in [
                            "A",
                            "B",
                            "C",
                            "D",
                            "F",
                        ]:

                            e_start = False

                        if (
                            e_start
                            and no_val
                            in [
                                "1",
                                "2",
                                "3",
                            ]
                        ):

                            editable_indices.append(
                                idx
                            )

                    col_e1, col_e2, col_e3 = (
                        st.columns(3)
                    )

                    for idx in editable_indices:

                        no_val = str(
                            df_boq.loc[
                                idx,
                                "NO",
                            ]
                        )

                        item_name = str(
                            df_boq.loc[
                                idx,
                                "Item",
                            ]
                        )

                        current_qty = (
                            parse_qty_num(
                                df_boq.loc[
                                    idx,
                                    "Unit/Volume",
                                ]
                            )
                        )

                        col_target = (
                            col_e1
                            if no_val == "1"
                            else (
                                col_e2
                                if no_val == "2"
                                else col_e3
                            )
                        )

                        with col_target:

                            new_qty = st.number_input(
                                (
                                    f"Qty Poin "
                                    f"{no_val}: "
                                    f"{item_name[:25]}..."
                                ),
                                min_value=0.0,
                                value=float(
                                    current_qty
                                ),
                                step=1.0,
                                key=(
                                    f"qty_e_"
                                    f"{no_val}_"
                                    f"{selected_label}_"
                                    f"{template_name}"
                                ),
                            )

                            df_boq.loc[
                                idx,
                                "Unit/Volume",
                            ] = (
                                str(
                                    int(new_qty)
                                )
                                if new_qty.is_integer()
                                else new_qty
                            )

                            df_boq.loc[
                                idx,
                                "TOTAL PRICE",
                            ] = calculate_detail_total(
                                df_boq.loc[
                                    idx,
                                    "Unit/Volume",
                                ],
                                df_boq.loc[
                                    idx,
                                    "UNIT PRICE",
                                ],
                            )

                # ==================================================================
                # TEMPLATE BOQ BARU
                #
                # PERUBAHAN UTAMA:
                #
                # EVCS:
                #   A + G editable
                #
                # BSS:
                #   A + D editable
                #
                # Berlaku untuk:
                #   DC20
                #   DC30
                #   DC60
                #   DC20 (NON JAVA)
                #   DC30 (NON JAVA)
                #   DC60 (NON JAVA)
                #   6S1P
                #   12S1P
                #   12S3P
                #
                # TOTAL PRICE tetap readonly.
                # ==================================================================

                else:

                    normalized_create_charger = (
                        normalize_charger_type(
                            charging_type
                        )
                    )

                    if normalized_create_charger in [
                        "DC20",
                        "DC30",
                        "DC60",
                    ]:

                        st.info(
                            "🆕 **Template BOQ Baru - EVCS**\n\n"
                            "Point **A dan G** dapat diedit "
                            "secara penuh. Anda dapat mengubah "
                            "**Item, Volume, Satuan, MERK, "
                            "dan UNIT PRICE**. "
                            "**TOTAL PRICE otomatis dihitung "
                            "Volume × UNIT PRICE**."
                        )

                    elif normalized_create_charger in [
                        "6S1P",
                        "12S1P",
                        "12S3P",
                    ]:

                        st.info(
                            "🆕 **Template BOQ Baru - BSS**\n\n"
                            "Point **A dan D** dapat diedit "
                            "secara penuh. Anda dapat mengubah "
                            "**Item, Volume, Satuan, MERK, "
                            "dan UNIT PRICE**. "
                            "**TOTAL PRICE otomatis dihitung "
                            "Volume × UNIT PRICE**."
                        )

                    # --------------------------------------------------------------
                    # PANGGIL EDITOR KHUSUS TEMPLATE BARU
                    # --------------------------------------------------------------

                    df_boq = edit_boq_sections(
                        df_boq,
                        charging_type,
                        (
                            f"create_new_template_"
                            f"{selected_label}_"
                            f"{normalized_create_charger}"
                        ),
                    )

                # ------------------------------------------------------------------
                # FINAL AUTO CALCULATION
                # ------------------------------------------------------------------

                (
                    df_boq,
                    sub_total,
                    vat_amount,
                    grand_total,
                ) = recalculate_boq_totals(
                    df_boq
                )

                # ------------------------------------------------------------------
                # DISPLAY
                # ------------------------------------------------------------------

                display_df = (
                    df_boq.copy()
                )

                for c in [
                    "UNIT PRICE",
                    "TOTAL PRICE",
                ]:

                    display_df[c] = (
                        display_df[c]
                        .apply(
                            lambda x:
                            format_currency(x)
                            if parse_price(x)
                            != 0
                            else "-"
                        )
                    )

                st.dataframe(
                    display_df,
                    use_container_width=True,
                    hide_index=True,
                )

                st.markdown("---")

                c_res1, c_res2 = (
                    st.columns(
                        [2, 1]
                    )
                )

                with c_res1:

                    st.success(
                        f"✅ Tabel BOQ Aktif: "
                        f"**{selected_site}** | "
                        f"Tipe Charger: "
                        f"**{charging_type}** | "
                        f"Wilayah: "
                        f"**{region_normalized}**"
                    )

                    st.caption(
                        f"Template: "
                        f"**{template_name}** | "
                        f"DB Value: "
                        f"**{template_db_value(template_name)}** | "
                        f"Sheet: "
                        f"**{target_sheet}**"
                    )

                with c_res2:

                    st.metric(
                        label=(
                            "Total Contractor Price "
                            "(Inc. VAT 11%)"
                        ),
                        value=format_currency(
                            grand_total
                        ),
                    )

                # ------------------------------------------------------------------
                # PDF
                # ------------------------------------------------------------------

                st.subheader(
                    "3. Action & Generate PDF"
                )

                col_btn1, col_btn2 = (
                    st.columns(2)
                )

                pdf_bytes = (
                    generate_boq_pdf(
                        selected_site,
                        site_address,
                        charging_type,
                        region_normalized,
                        df_boq,
                        sub_total,
                        vat_amount,
                        grand_total,
                    )
                )

                safe_site_name = re.sub(
                    r'[\\/*?:"<>|]',
                    "_",
                    str(
                        selected_site
                    ),
                )

                safe_charger = re.sub(
                    r'[\\/*?:"<>|]',
                    "_",
                    str(
                        charging_type
                    ),
                )

                filename_pdf = (
                    f"BOQ_"
                    f"{safe_charger}_"
                    f"{region_normalized}_"
                    f"{safe_site_name}.pdf"
                )

                # ------------------------------------------------------------------
                # SAVE
                # ------------------------------------------------------------------

                with col_btn1:

                    if st.button(
                        "🚀 Simpan ke Database (DB BOQ)",
                        type="primary",
                        disabled=is_already_saved,
                        key=(
                            "btn_save_boq_"
                            f"{selected_label}_"
                            f"{template_name}"
                        ),
                    ):

                        if is_already_saved:

                            st.error(
                                "❌ Gagal Simpan! "
                                "Kombinasi Site + Charger "
                                "sudah terdaftar."
                            )

                        else:

                            with st.spinner(
                                "Memproses penyimpanan..."
                            ):

                                boq_no = (
                                    save_to_db_boq(
                                        selected_site,
                                        charging_type,
                                        sub_total,
                                        grand_total,
                                        epc_name,
                                        template_name,
                                    )
                                )

                                if boq_no:

                                    update_google_sheet_summary(
                                        selected_site,
                                        selected_site,
                                        sub_total,
                                        grand_total,
                                    )

                                    st.cache_data.clear()

                                    st.session_state[
                                        "last_saved_info"
                                    ] = {
                                        "boq_no": boq_no,
                                        "site_name": selected_site,
                                        "pdf_bytes": pdf_bytes,
                                        "filename": filename_pdf,
                                    }

                                    st.rerun()

                # ------------------------------------------------------------------
                # DOWNLOAD
                # ------------------------------------------------------------------

                with col_btn2:

                    st.download_button(
                        label=(
                            "📥 Download PDF BOQ "
                            "(Draft Preview)"
                        ),
                        data=pdf_bytes,
                        file_name=filename_pdf,
                        mime="application/pdf",
                        key=(
                            "dl_active_"
                            f"{selected_site}_"
                            f"{template_name}"
                        ),
                    )

    # ==========================================================================
    # TAB EDIT
    # ==========================================================================

    with tab_edit:

        st.subheader(
            "✏️ Edit, Reuse, & "
            "Re-Download BOQ Tersimpan"
        )

        saved_boq_list = (
            get_all_saved_boq()
        )

        if not saved_boq_list:

            st.info(
                "ℹ️ Belum ada data BOQ "
                "yang tersimpan di "
                "`DB BOQ`."
            )

        else:

            boq_options = {

                (
                    f"{item.get('BOQ No.', '-')}: "
                    f"{item.get('Site Name', '-')} "
                    f"[{item.get('Charger Type', '-')}]"
                ): item

                for item in saved_boq_list
            }

            selected_boq_key = st.selectbox(
                "Pilih Nomor BOQ yang Ingin "
                "Di-edit / Re-assign / Re-Download",
                sorted(
                    list(
                        boq_options.keys()
                    )
                ),
                key="edit_boq_selector",
            )

            selected_data = (
                boq_options[
                    selected_boq_key
                ]
            )

            st.markdown("---")

            st.caption(
                f"📌 Editing Row: "
                f"`{selected_data['row_idx']}` | "
                f"No. BOQ: "
                f"`{selected_data.get('BOQ No.', '-')}`"
            )

            old_site_name = (
                selected_data.get(
                    "Site Name",
                    "",
                )
            )

            old_charger_type = (
                selected_data.get(
                    "Charger Type",
                    "",
                )
            )

            # ------------------------------------------------------------------
            # TEMPLATE
            # ------------------------------------------------------------------

            stored_template_value = (
                selected_data.get(
                    "Template",
                    "",
                )
            )

            stored_template_value = (
                str(
                    stored_template_value
                    or ""
                )
                .strip()
                .upper()
            )

            if (
                stored_template_value
                == "BOQ BARU"
            ):

                default_template_index = 1

            else:

                default_template_index = 0

            st.markdown(
                "### 📑 Template BOQ"
            )

            edit_template_name = st.selectbox(
                "Pilih Template untuk "
                "Edit / Re-Download PDF",
                list(
                    TEMPLATE_OPTIONS.keys()
                ),
                index=default_template_index,
                key=(
                    f"edit_boq_template_"
                    f"{selected_boq_key}"
                ),
            )

            if stored_template_value:

                st.caption(
                    f"📌 Template tersimpan di "
                    f"DB BOQ kolom H: "
                    f"**{stored_template_value}**"
                )

            else:

                st.caption(
                    "ℹ️ BOQ lama belum memiliki "
                    "record template di kolom H. "
                    "Template dapat dipilih kembali."
                )

            # ------------------------------------------------------------------
            # QUERY SITE
            # ------------------------------------------------------------------

            all_site_options, all_site_map = (
                fetch_query_site_options(
                    exclude_saved=False
                )
            )

            matched_meta = next(
                (
                    meta

                    for label, meta
                    in all_site_map.items()

                    if (
                        meta[
                            "site_name"
                        ]
                        .strip()
                        .lower()
                        ==
                        old_site_name
                        .strip()
                        .lower()
                    )

                    and (
                        meta[
                            "charger"
                        ]
                        .strip()
                        .lower()
                        ==
                        old_charger_type
                        .strip()
                        .lower()
                    )
                ),
                {
                    "address": "-",
                    "province": "JAVA",
                    "epc": "-",
                },
            )

            edit_address = matched_meta[
                "address"
            ]

            edit_province = matched_meta[
                "province"
            ]

            # ------------------------------------------------------------------
            # BASIC DETAILS
            # ------------------------------------------------------------------

            col_e1, col_e2 = (
                st.columns(2)
            )

            with col_e1:

                use_existing_query_site = (
                    st.checkbox(
                        "Ganti/Re-assign dengan "
                        "Site Aktif dari Sheet Query",
                        value=False,
                        key=(
                            f"use_query_site_"
                            f"{selected_boq_key}"
                        ),
                    )
                )

                if use_existing_query_site:

                    selected_reassign_label = (
                        st.selectbox(
                            "Pilih Site & Charger Pengganti",
                            sorted(
                                list(
                                    all_site_options
                                )
                            ),
                            key=(
                                f"reassign_site_"
                                f"{selected_boq_key}"
                            ),
                        )
                    )

                    reassign_meta = (
                        all_site_map[
                            selected_reassign_label
                        ]
                    )

                    edit_site_name = (
                        reassign_meta[
                            "site_name"
                        ]
                    )

                    edit_charger_type = (
                        reassign_meta[
                            "charger"
                        ]
                    )

                    edit_epc_name = (
                        reassign_meta[
                            "epc"
                        ]
                    )

                    edit_address = (
                        reassign_meta[
                            "address"
                        ]
                    )

                    edit_province = (
                        reassign_meta[
                            "province"
                        ]
                    )

                    st.info(
                        "💡 Menimpa BOQ No. "
                        f"`{selected_data.get('BOQ No.', '-')}` "
                        "ke Site Baru: "
                        f"**{edit_site_name} "
                        f"({edit_charger_type})**"
                    )

                else:

                    edit_site_name = (
                        st.text_input(
                            "Site Name",
                            value=old_site_name,
                            key=(
                                f"edit_site_"
                                f"{selected_boq_key}"
                            ),
                        )
                    )

                    edit_charger_type = (
                        st.text_input(
                            "Charger Type",
                            value=old_charger_type,
                            key=(
                                f"edit_charger_"
                                f"{selected_boq_key}"
                            ),
                        )
                    )

                    edit_epc_name = (
                        st.text_input(
                            "EPC Name",
                            value=selected_data.get(
                                "EPC Name",
                                "-",
                            ),
                            key=(
                                f"edit_epc_"
                                f"{selected_boq_key}"
                            ),
                        )
                    )

            with col_e2:

                edit_sub_total_manual = (
                    parse_price(
                        selected_data.get(
                            "BOQ Amount Exc. PPN",
                            0,
                        )
                    )
                )

                edit_grand_total_manual = (
                    parse_price(
                        selected_data.get(
                            "BOQ Amount inc. PPN",
                            0,
                        )
                    )
                )

                st.caption(
                    "Nilai di bawah akan "
                    "otomatis mengikuti "
                    "kalkulasi detail template."
                )

                st.metric(
                    "DB Sub Total Saat Ini",
                    format_currency(
                        edit_sub_total_manual
                    ),
                )

                st.metric(
                    "DB Grand Total Saat Ini",
                    format_currency(
                        edit_grand_total_manual
                    ),
                )

            # ------------------------------------------------------------------
            # LOAD SELECTED TEMPLATE
            # ------------------------------------------------------------------

            (
                df_boq_edit,
                target_sheet_edit,
                reg_norm,
            ) = load_boq_dataframe(
                edit_charger_type,
                edit_province,
                edit_template_name,
            )

            if (
                df_boq_edit is not None
                and not df_boq_edit.empty
            ):

                st.markdown("---")

                st.subheader(
                    "✏️ Edit Detail BOQ"
                )

                st.caption(
                    f"Template: "
                    f"**{edit_template_name}** | "
                    f"Sheet: "
                    f"**{target_sheet_edit}** | "
                    f"Region: "
                    f"**{reg_norm}**"
                )

                # ------------------------------------------------------------------
                # EDIT DETAIL
                # ------------------------------------------------------------------

                df_boq_edit = (
                    edit_boq_sections(
                        df_boq_edit,
                        edit_charger_type,
                        (
                            f"edit_detail_"
                            f"{selected_boq_key}_"
                            f"{edit_template_name}_"
                            f"{edit_charger_type}"
                        ),
                    )
                )

                # ------------------------------------------------------------------
                # FINAL RECALCULATION
                # ------------------------------------------------------------------

                (
                    df_boq_edit,
                    edit_sub_total,
                    vat_edit,
                    edit_grand_total,
                ) = recalculate_boq_totals(
                    df_boq_edit
                )

                st.markdown("---")

                # ------------------------------------------------------------------
                # PREVIEW TABLE
                # ------------------------------------------------------------------

                st.subheader(
                    "👁️ Preview BOQ"
                )

                display_edit_df = (
                    df_boq_edit.copy()
                )

                for c in [
                    "UNIT PRICE",
                    "TOTAL PRICE",
                ]:

                    display_edit_df[c] = (
                        display_edit_df[c]
                        .apply(
                            lambda x:
                            format_currency(x)
                            if parse_price(x)
                            != 0
                            else "-"
                        )
                    )

                st.dataframe(
                    display_edit_df,
                    use_container_width=True,
                    hide_index=True,
                )

                # ------------------------------------------------------------------
                # SUMMARY
                # ------------------------------------------------------------------

                summary_c1, summary_c2, summary_c3 = (
                    st.columns(3)
                )

                with summary_c1:

                    st.metric(
                        "Sub Total / Exc. PPN",
                        format_currency(
                            edit_sub_total
                        ),
                    )

                with summary_c2:

                    st.metric(
                        "PPN 11%",
                        format_currency(
                            vat_edit
                        ),
                    )

                with summary_c3:

                    st.metric(
                        "Grand Total / Inc. PPN",
                        format_currency(
                            edit_grand_total
                        ),
                    )

                # ------------------------------------------------------------------
                # ACTION
                # ------------------------------------------------------------------

                st.markdown("---")

                col_act1, col_act2 = (
                    st.columns(2)
                )

                # ------------------------------------------------------------------
                # SAVE UPDATE
                # ------------------------------------------------------------------

                with col_act1:

                    if st.button(
                        "💾 Save & Update BOQ Database",
                        type="primary",
                        key=(
                            f"btn_update_boq_"
                            f"{selected_boq_key}"
                        ),
                    ):

                        with st.spinner(
                            "Memperbarui DB BOQ "
                            "& Sum Project..."
                        ):

                            success = (
                                update_db_boq_row(
                                    row_idx=selected_data[
                                        "row_idx"
                                    ],
                                    old_site_name=(
                                        old_site_name
                                    ),
                                    new_site_name=(
                                        edit_site_name
                                    ),
                                    charger_capacity=(
                                        edit_charger_type
                                    ),
                                    sub_total=(
                                        edit_sub_total
                                    ),
                                    grand_total=(
                                        edit_grand_total
                                    ),
                                    epc_name=(
                                        edit_epc_name
                                    ),
                                    template_name=(
                                        edit_template_name
                                    ),
                                )
                            )

                            if success:

                                st.cache_data.clear()

                                st.success(
                                    f"🎉 BOQ "
                                    f"`{selected_data.get('BOQ No.', '-')}` "
                                    "berhasil diperbarui."
                                )

                                st.rerun()

                # ------------------------------------------------------------------
                # RE-DOWNLOAD PDF
                # ------------------------------------------------------------------

                with col_act2:

                    pdf_bytes_edit = (
                        generate_boq_pdf(
                            edit_site_name,
                            edit_address,
                            edit_charger_type,
                            reg_norm,
                            df_boq_edit,
                            edit_sub_total,
                            vat_edit,
                            edit_grand_total,
                        )
                    )

                    safe_edit_site = re.sub(
                        r'[\\/*?:"<>|]',
                        "_",
                        str(
                            edit_site_name
                        ),
                    )

                    safe_edit_charger = re.sub(
                        r'[\\/*?:"<>|]',
                        "_",
                        str(
                            edit_charger_type
                        ),
                    )

                    safe_boq_no = re.sub(
                        r'[\\/*?:"<>|]',
                        "_",
                        str(
                            selected_data.get(
                                "BOQ No.",
                                "",
                            )
                        ),
                    )

                    st.download_button(
                        label=(
                            "📥 Re-Download PDF BOQ "
                            f"({selected_data.get('BOQ No.', '-')})"
                        ),
                        data=pdf_bytes_edit,
                        file_name=(
                            f"BOQ_"
                            f"{safe_boq_no}_"
                            f"{safe_edit_site}_"
                            f"{safe_edit_charger}.pdf"
                        ),
                        mime="application/pdf",
                        key=(
                            f"download_edit_"
                            f"{selected_boq_key}_"
                            f"{edit_template_name}"
                        ),
                    )


# ==============================================================================
# RUN
# ==============================================================================

if __name__ == "__main__":
    render()
