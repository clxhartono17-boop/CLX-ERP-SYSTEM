import ast
import datetime
import hashlib
import io
import json
import os
import re
import sys
from xml.sax.saxutils import escape as xml_escape

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, A5, portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.platypus import (
    Image,
    KeepInFrame,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
import streamlit as st


# ==============================================================================
# SAFE IMPORT / ROOT CONFIGURATION
# ==============================================================================
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, "../"))
ROOT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, "../../"))

for _p in [ROOT_DIR, PARENT_DIR]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

CWD = os.getcwd()

if CWD not in sys.path:
    sys.path.insert(0, CWD)

try:
    from services.gsheet import get_google_sheet_connection
except ModuleNotFoundError:
    def get_google_sheet_connection():
        return None


# ==============================================================================
# TEMPLATE & COMPANY CONSTANTS
# ==============================================================================
OLD_TEMPLATE_FILENAME = "Template BOQ Vgreen.xlsx"
NEW_TEMPLATE_FILENAME = "New Template BOQ Sept 2026 All Charger.xlsx"

BOQ_EDITOR_VERSION = "QTY_TEMPLATE_AUTO_TOTAL_V15"

COMPANY_OPTIONS = ["CLX", "Sunrise"]

COMPANY_CONFIGS = {
    "CLX": {
        "code": "CLX",
        "name": "PT. Connectivity Leads eXcellence",
        "subtitle": "EV Charging & Battery Swapping Infrastructure Services",
        "logo_candidates": ["CLX.png", "clx.png"],
        "primary_hex": "#1B365D",
        "border_primary_hex": "#1B365D",
        "accent_hex": "#0284C7",
        "badge_text_hex": "#0284C7",
        "title_text_hex": "#1B365D",
        "section_bg_hex": "#DCE7F5",
        "subsection_bg_hex": "#EDF2F7",
        "zebra_bg_hex": "#F8FAFC",
        "card_bg_hex": "#F1F5F9",
        "header_text_hex": "#FFFFFF",
        "section_text_hex": "#1B365D",
        "grand_text_hex": "#FFFFFF",
        "ms_header_hex": "#2C3E50",
        "ms_header_text_hex": "#FFFFFF",
    },
    "Sunrise": {
        "code": "SI",
        "name": "PT. Sunrise Internusa",
        "subtitle": "EV Charging & Battery Swapping Infrastructure Services",
        "logo_candidates": ["SUNRISE.png", "Sunrise.png", "sunrise.png"],
        "primary_hex": "#FACC15",
        "border_primary_hex": "#CA8A04",
        "accent_hex": "#EAB308",
        "badge_text_hex": "#B45309",
        "title_text_hex": "#854D0E",
        "section_bg_hex": "#FDE047",
        "subsection_bg_hex": "#FEF9C3",
        "zebra_bg_hex": "#FFFBEB",
        "card_bg_hex": "#FEFCE8",
        "header_text_hex": "#0F172A",
        "section_text_hex": "#0F172A",
        "grand_text_hex": "#0F172A",
        "ms_header_hex": "#FACC15",
        "ms_header_text_hex": "#0F172A",
    },
}

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

DB_BOQ_SHEET = "DB BOQ"
DB_BOQ_MS_SHEET = "DB BOQ MS"
QUERY_SHEET = "Query"
SUM_PROJECT_SHEET = "Sum Project"

DB_BOQ_HEADERS = [
    "No",
    "BOQ No.",
    "Site Name",
    "Charger Type",
    "BOQ Amount Exc. PPN",
    "BOQ Amount inc. PPN",
    "EPC Name",
    "Template",
    "Items JSON",
]

BOQ_MS_HEADERS = [
    "No",
    "Date",
    "No. BOQ",
    "SN Machine",
    "Site Name",
    "Charging Type",
    "Item Name",
    "Price",
    "Periode",
]

BOQ_MS_HELPER_HEADERS = [
    "Qty",
    "Template",
    "Region",
    "Issue Note",
    "Photo Evident 1",
    "Photo Evident 2",
]


# ==============================================================================
# GENERIC HELPERS
# ==============================================================================
def normalize_company_key(company):
    c = clean_text(company).upper()
    if "SUNRISE" in c or c in {"SI", "SR"}:
        return "Sunrise"
    return "CLX"


def get_company_config(company):
    return COMPANY_CONFIGS[normalize_company_key(company)]


def detect_company_from_boq_no(boq_no):
    s = clean_text(boq_no).upper()
    if "/SI/" in s or "/SUNRISE/" in s or "/SR/" in s:
        return "Sunrise"
    return "CLX"


def detect_company_from_context(boq_no="", epc_name=""):
    b = clean_text(boq_no).upper()
    e = clean_text(epc_name).upper()
    if "/SI/" in b or "/SUNRISE/" in b or "/SR/" in b or "SUNRISE" in e:
        return "Sunrise"
    return "CLX"


def replace_boq_company_code(boq_no, company):
    raw = clean_text(boq_no)
    if not raw or raw == "-":
        return raw
    target_code = get_company_config(company)["code"]
    updated = re.sub(
        r"^(\d+)/(?:CLX|SI|SUNRISE|SR)(/BOQ/)",
        rf"\1/{target_code}\2",
        raw,
        flags=re.IGNORECASE,
    )
    return updated


def parse_price(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0

    if isinstance(val, (int, float)):
        return float(val)

    s = str(val).strip()

    if not s or s.upper() in ["#VALUE!", "#REF!", "#N/A", "NONE", "NAN"]:
        return 0.0

    s = (
        s.replace("Rp", "")
        .replace("rp", "")
        .replace(" ", "")
    )

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")

    elif "," in s:
        tail = s.rsplit(",", 1)[-1]

        if len(tail) <= 2:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")

    else:
        if (
            s.count(".") > 1
            or (
                s.count(".") == 1
                and len(s.rsplit(".", 1)[-1]) == 3
            )
        ):
            s = s.replace(".", "")

    try:
        return float(s)
    except Exception:
        m = re.search(r"[-+]?\d+(?:\.\d+)?", s)

        return float(m.group()) if m else 0.0


def parse_qty_num(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0

    if isinstance(val, (int, float)):
        return float(val)

    s = str(val).strip().replace(" ", "")

    if not s or s.upper() in ["#VALUE!", "#REF!", "#N/A", "NONE", "NAN"]:
        return 0.0

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")

    elif "," in s:
        s = s.replace(",", ".")

    try:
        return float(s)
    except Exception:
        return 0.0


def format_currency(value):
    try:
        num = float(value)
        if abs(num) < 0.005:
            return "Rp. 0"
        if num < 0:
            return f"-Rp. {abs(num):,.0f}".replace(",", ".")
        return f"Rp. {num:,.0f}".replace(",", ".")
    except Exception:
        return "Rp. 0"


def _format_qty_clean(val):
    s = clean_text(val)
    if not s:
        return ""
    try:
        num = float(s.replace(",", "."))
        if num.is_integer():
            return f"{int(num):,}".replace(",", ".")
        return f"{num:g}"
    except Exception:
        return s


def _esc_pdf_text(val):
    return xml_escape(clean_text(val))


def clean_text(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""

    s = str(value).strip()
    if s.upper() in ["#VALUE!", "#REF!", "#N/A", "NONE", "NAN"]:
        return ""

    return s


def map_to_standard_province(raw_province):
    s = clean_text(raw_province).upper()

    if not s:
        return "JAVA"

    s_norm = re.sub(
        r"[^A-Z0-9]+",
        " ",
        s
    ).strip()

    java_terms = {
        "JAKARTA",
        "DKI JAKARTA",
        "BANTEN",
        "JAWA",
        "JAWA BARAT",
        "JAWA TENGAH",
        "JAWA TIMUR",
        "WEST JAVA",
        "CENTRAL JAVA",
        "EAST JAVA",
        "YOGYAKARTA",
        "SPECIAL REGION OF YOGYAKARTA",
        "DI YOGYAKARTA",
        "D I YOGYAKARTA",
    }

    if (
        s_norm in java_terms
        or any(
            term in s_norm
            for term in [
                "WEST JAVA",
                "CENTRAL JAVA",
                "EAST JAVA",
                "JAWA BARAT",
                "JAWA TENGAH",
                "JAWA TIMUR",
            ]
        )
    ):
        return "JAVA"

    sumatera_terms = [
        "SUMATERA",
        "SUMATRA",
        "ACEH",
        "NANGGROE ACEH",
        "RIAU",
        "JAMBI",
        "BENGKULU",
        "LAMPUNG",
        "BANGKA",
        "BELITUNG",
        "KEPULAUAN RIAU",
        "RIAU ISLANDS",
        "NORTH SUMATERA",
        "NORTH SUMATRA",
        "WEST SUMATERA",
        "WEST SUMATRA",
        "SOUTH SUMATERA",
        "SOUTH SUMATRA",
    ]

    if any(x in s_norm for x in sumatera_terms):
        return "SUMATERA"

    bali_nt_terms = [
        "BALI",
        "NUSA TENGGARA",
        "NUSA TENGGARA BARAT",
        "NUSA TENGGARA TIMUR",
        "NTB",
        "NTT",
        "WEST NUSA TENGGARA",
        "EAST NUSA TENGGARA",
    ]

    if any(x in s_norm for x in bali_nt_terms):
        return "BALI NUSATENGGARA"

    if (
        "KALIMANTAN" in s_norm
        or "BORNEO" in s_norm
    ):
        return "KALIMANTAN"

    sulawesi_terms = [
        "SULAWESI",
        "GORONTALO",
        "NORTH SULAWESI",
        "CENTRAL SULAWESI",
        "SOUTH SULAWESI",
        "SOUTHEAST SULAWESI",
        "WEST SULAWESI",
    ]

    if any(x in s_norm for x in sulawesi_terms):
        return "SULAWESI"

    return s_norm


def normalize_charger_type(charging_type):
    s = (
        clean_text(charging_type)
        .upper()
        .replace(" ", "")
    )

    aliases = {
        "20KW": "DC20",
        "DC20KW": "DC20",
        "DC20": "DC20",

        "30KW": "DC30",
        "DC30KW": "DC30",
        "DC30": "DC30",

        "60KW": "DC60",
        "DC60KW": "DC60",
        "DC60": "DC60",

        "6S1P": "6S1P",
        "BSS6S1P": "6S1P",

        "12S1P": "12S1P",
        "BSS12S1P": "12S1P",

        "12S3P": "12S3P",
        "BSS12S3P": "12S3P",

        "7KW": "7KW",
        "AC7": "7KW",

        "22KW": "22KW",
        "AC22": "22KW",

        "DC120": "DC120",
        "120KW": "DC120",
    }

    return aliases.get(
        s,
        clean_text(charging_type).strip()
    )


def generate_boq_number(sequence_num=1, company="CLX"):
    now = datetime.datetime.now()

    roman = {
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
    }[now.month]

    company_code = get_company_config(company)["code"]
    return f"{int(sequence_num):04d}/{company_code}/BOQ/{roman}/{now.year}"


def generate_boq_ms_number(sequence_num=1, date_obj=None, company="CLX"):
    now = date_obj or datetime.datetime.now()

    roman = {
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
    }[now.month]

    company_code = get_company_config(company)["code"]
    return f"{int(sequence_num):04d}/{company_code}/BOQ/MS/{roman}/{now.year}"


def _candidate_asset_dirs():
    dirs = [
        os.path.join(CURRENT_DIR, "assets", "templates"),
        os.path.join(PARENT_DIR, "assets", "templates"),
        os.path.join(ROOT_DIR, "assets", "templates"),
        os.path.join(CWD, "assets", "templates"),
        os.path.join(CWD, "CLX-ERP-SYSTEM", "assets", "templates"),
        os.path.join(CWD, "assets", "templates", "boq"),
    ]
    seen = set()
    unique_dirs = []
    for d in dirs:
        norm = os.path.abspath(d)
        if norm not in seen:
            seen.add(norm)
            unique_dirs.append(norm)
    return unique_dirs


def get_template_path(filename):
    asset_dirs = _candidate_asset_dirs()

    for base_dir in asset_dirs:
        candidate = os.path.join(base_dir, filename)
        if os.path.exists(candidate):
            return candidate

    base, ext = os.path.splitext(filename)
    variants = [
        f"{base}(4){ext}",
        f"{base} (4){ext}",
        f"{base}(3){ext}",
        f"{base} (3){ext}",
    ]

    for variant in variants:
        for base_dir in asset_dirs:
            candidate = os.path.join(base_dir, variant)
            if os.path.exists(candidate):
                return candidate

    return os.path.join(asset_dirs[0], filename)


def get_logo_path(company="CLX"):
    cfg = get_company_config(company)
    logo_filenames = cfg["logo_candidates"]
    target_lower = {f.lower() for f in logo_filenames}
    asset_dirs = _candidate_asset_dirs()

    for base_dir in asset_dirs:
        for fname in logo_filenames:
            candidate = os.path.join(base_dir, fname)
            if os.path.isfile(candidate):
                return candidate

    for base_dir in asset_dirs:
        if os.path.isdir(base_dir):
            try:
                for entry in os.listdir(base_dir):
                    if entry.lower() in target_lower:
                        full_path = os.path.join(base_dir, entry)
                        if os.path.isfile(full_path):
                            return full_path
            except Exception:
                pass

    return None


def template_db_value(template_name):
    return TEMPLATE_DB_VALUES.get(
        template_name,
        template_name,
    )


def template_name_from_db_value(value):
    v = clean_text(value).upper()

    if v == "BOQ LAMA":
        return "Template BOQ Vgreen Lama"

    return "New Template BOQ Sept 2026"


def get_new_template_sheet_name(
    charging_type,
    province_str,
):
    c = normalize_charger_type(charging_type)
    region = map_to_standard_province(province_str)

    supported_regional_chargers = {
        "DC20",
        "DC30",
        "DC60",
        "6S1P",
        "12S1P",
        "12S3P",
    }

    if c not in supported_regional_chargers:
        return c

    if region == "JAVA":
        return c

    non_java_regions = {
        "SUMATERA",
        "BALI NUSATENGGARA",
        "KALIMANTAN",
        "SULAWESI",
    }

    if region in non_java_regions:
        return f"{c} (NON JAVA)"

    raise ValueError(
        f"Province/Region tidak dikenali untuk template BOQ baru: "
        f"'{province_str}' -> '{region}'. "
    )


def detect_template_header(
    raw,
    keywords=("NO", "ITEM"),
):
    for i in range(
        min(30, len(raw))
    ):
        row = [
            clean_text(x).upper()
            for x in raw.iloc[i].tolist()
        ]

        if all(
            any(k in cell for cell in row)
            for k in keywords
        ):
            return i

    return 0


def find_column_by_keywords(
    columns,
    keywords,
):
    for c in columns:
        s = clean_text(c).upper()
        if any(
            k == s or (len(k) > 2 and k in s)
            for k in keywords
        ):
            return c

    for c in columns:
        s = clean_text(c).upper()
        for k in keywords:
            if k in s:
                return c

    return None


def _normalize_template_headers(headers):
    return [
        clean_text(x)
        for x in list(headers)
    ]


def build_new_template_column_mapping(columns):
    cols = _normalize_template_headers(columns)

    mapping = {}

    mapping["NO"] = find_column_by_keywords(
        cols,
        ["NO"],
    )

    mapping["ITEM"] = find_column_by_keywords(
        cols,
        ["ITEM", "DESCRIPTION", "MATERIAL"],
    )

    mapping["UNIT_VOLUME"] = find_column_by_keywords(
        cols,
        [
            "UNIT/VOLUME",
            "UNIT/VOL",
            "UNIT VOL",
            "QTY/VOL",
            "QTY",
            "VOLUME",
        ],
    )

    mapping["SATUAN"] = find_column_by_keywords(
        cols,
        [
            "SATUAN/UOM",
            "SATUAN",
            "UOM",
            "LOT",
        ],
    )

    if (
        mapping["SATUAN"] is None
        and mapping["UNIT_VOLUME"] is not None
    ):
        try:
            qty_pos = cols.index(
                mapping["UNIT_VOLUME"]
            )

            if qty_pos + 1 < len(cols):
                mapping["SATUAN"] = cols[
                    qty_pos + 1
                ]
        except Exception:
            pass

    mapping["MERK"] = find_column_by_keywords(
        cols,
        ["MERK", "BRAND"],
    )

    mapping["UNIT_PRICE"] = find_column_by_keywords(
        cols,
        ["UNIT PRICE", "PRICE"],
    )

    mapping["TOTAL_PRICE"] = find_column_by_keywords(
        cols,
        ["TOTAL PRICE", "TOTAL"],
    )

    return mapping


# ==============================================================================
# SERIALISASI & RESTORASI SNAPSHOT ITEM BOQ (OPSI 1 - KOLOM ITEMS JSON)
# ==============================================================================
def serialize_boq_items_json(df_boq):
    """
    Serialisasi daftar Qty/Volume setiap item pada df_boq menjadi string JSON
    ringkas agar dapat disimpan pada kolom 'Items JSON' di sheet DB BOQ.
    """
    if df_boq is None or df_boq.empty:
        return "[]"

    records = []
    for idx, row in df_boq.iterrows():
        no_val = clean_text(row.get("NO", ""))
        item_val = clean_text(row.get("Item", ""))
        up_val = parse_price(row.get("UNIT PRICE", 0))
        raw_qty = row.get("Unit/Volume", "")

        if up_val != 0:
            q_num = parse_qty_num(raw_qty)
            q_val = int(q_num) if float(q_num).is_integer() else round(float(q_num), 4)
        else:
            q_val = clean_text(raw_qty)

        records.append(
            {
                "r": int(idx),
                "n": no_val,
                "i": item_val,
                "q": q_val,
            }
        )

    payload = json.dumps(records, ensure_ascii=False, separators=(",", ":"))

    # Jika melebihi batas karakter sel Google Sheets (50.000 karakter), simpan
    # versi lebih ringkas tanpa teks item panjang.
    if len(payload) > 48000:
        compact = [
            {"r": rec["r"], "n": rec["n"], "q": rec["q"]}
            for rec in records
        ]
        payload = json.dumps(compact, ensure_ascii=False, separators=(",", ":"))

    return payload


def apply_saved_boq_items_json(df_boq, items_json_str):
    """
    Memulihkan nilai Unit/Volume (Qty) dari string JSON yang tersimpan pada DB BOQ
    ke dalam DataFrame BOQ hasil load template, lalu menghitung ulang seluruh total.
    """
    if df_boq is None or df_boq.empty:
        return df_boq

    raw_json = clean_text(items_json_str)
    if not raw_json or raw_json in {"[]", "{}"}:
        return df_boq

    try:
        saved_items = json.loads(raw_json)
    except Exception:
        return df_boq

    if not isinstance(saved_items, list) or not saved_items:
        return df_boq

    df = df_boq.copy()
    for col in ["Unit/Volume", "UNIT PRICE", "TOTAL PRICE"]:
        if col in df.columns:
            df[col] = df[col].astype(object)

    # Index lookup berdasarkan (NO, Item) sebagai fallback jika indeks baris bergeser
    lookup_by_no_item = {}
    for idx, row in df.iterrows():
        key = (
            clean_text(row.get("NO", "")).upper(),
            clean_text(row.get("Item", "")).upper(),
        )
        lookup_by_no_item.setdefault(key, []).append(idx)

    used_indices = set()

    for rec in saved_items:
        if not isinstance(rec, dict):
            continue

        row_pos = rec.get("r", rec.get("idx"))
        no_saved = clean_text(rec.get("n", rec.get("NO", ""))).upper()
        item_saved = clean_text(rec.get("i", rec.get("Item", ""))).upper()
        qty_saved = rec.get("q", rec.get("qty"))

        target_idx = None

        # 1. Cocokkan posisi baris & NO (dan Item jika tersedia)
        if isinstance(row_pos, int) and 0 <= row_pos < len(df) and row_pos not in used_indices:
            cur_no = clean_text(df.loc[row_pos, "NO"]).upper()
            cur_item = clean_text(df.loc[row_pos, "Item"]).upper()
            if cur_no == no_saved and (not item_saved or cur_item == item_saved):
                target_idx = row_pos

        # 2. Fallback: cocokkan kombinasi (NO, Item) yang belum terpakai
        if target_idx is None and item_saved:
            candidates = lookup_by_no_item.get((no_saved, item_saved), [])
            for cand in candidates:
                if cand not in used_indices:
                    target_idx = cand
                    break

        if target_idx is None:
            continue

        used_indices.add(target_idx)
        unit_price = parse_price(df.loc[target_idx, "UNIT PRICE"])

        if unit_price != 0:
            q_num = parse_qty_num(qty_saved)
            df.loc[target_idx, "Unit/Volume"] = (
                int(q_num) if float(q_num).is_integer() else q_num
            )
            df.loc[target_idx, "TOTAL PRICE"] = calculate_detail_total(q_num, unit_price)
        else:
            if qty_saved is not None and clean_text(qty_saved) != "":
                df.loc[target_idx, "Unit/Volume"] = clean_text(qty_saved)

    df, _, _, _ = recalculate_boq_totals(df)
    return df


# ==============================================================================
# GOOGLE SHEETS HELPERS
# ==============================================================================
def _get_book():
    try:
        return get_google_sheet_connection()
    except Exception:
        return None


def _get_ws(sheet_name):
    book = _get_book()

    if book is None:
        return None

    try:
        return book.worksheet(sheet_name)
    except Exception:
        try:
            return book.open_worksheet(sheet_name)
        except Exception:
            return None


def _values_to_df(values):
    if not values:
        return pd.DataFrame()

    header = list(values[0])
    width = len(header)

    rows = []

    for r in values[1:]:
        rr = (
            list(r)
            + [""] * max(
                0,
                width - len(r),
            )
        )

        rows.append(
            rr[:width]
        )

    return pd.DataFrame(
        rows,
        columns=header,
    )


@st.cache_data(
    ttl=120,
    show_spinner=False,
)
def read_sheet_df(sheet_name):
    ws = _get_ws(sheet_name)

    if ws is None:
        return pd.DataFrame()

    try:
        return _values_to_df(
            ws.get_all_values()
        )
    except Exception:
        return pd.DataFrame()


# ==============================================================================
# DB BOQ
# ==============================================================================
def ensure_db_boq_headers(worksheet):
    headers = DB_BOQ_HEADERS

    try:
        vals = worksheet.get_all_values()
        first = vals[0] if vals else []

        if first[:len(headers)] != headers:
            worksheet.update(
                "A1:I1",
                [headers],
            )
    except Exception:
        pass

    return headers


@st.cache_data(
    ttl=120,
    show_spinner=False,
)
def get_all_saved_boq():
    df = read_sheet_df(DB_BOQ_SHEET)

    if df.empty:
        return []

    cols = list(df.columns)

    def col(letter, fallback):
        idx = ord(letter) - 65
        if idx < len(cols):
            return cols[idx]
        return fallback

    (
        no_c,
        boq_c,
        site_c,
        charger_c,
        dpp_c,
        inc_c,
        epc_c,
        temp_c,
        items_c,
    ) = [
        col(
            chr(65 + i),
            "",
        )
        for i in range(9)
    ]

    out = []

    for idx, row in df.iterrows():
        if not any(
            clean_text(v)
            for v in row.tolist()
        ):
            continue

        out.append(
            {
                "row_idx": idx + 2,
                "No": clean_text(row.get(no_c, "")),
                "BOQ No.": clean_text(row.get(boq_c, "")),
                "Site Name": clean_text(row.get(site_c, "")),
                "Charger Type": clean_text(row.get(charger_c, "")),
                "BOQ Amount Exc. PPN": parse_price(row.get(dpp_c, 0)),
                "BOQ Amount inc. PPN": parse_price(row.get(inc_c, 0)),
                "EPC Name": clean_text(row.get(epc_c, "")),
                "Template": clean_text(row.get(temp_c, "")),
                "Items JSON": clean_text(row.get(items_c, "")) if items_c else "",
            }
        )

    return out


def _next_boq_sequence():
    max_no = 0
    for r in get_all_saved_boq():
        m = re.match(r"^(\d+)", clean_text(r.get("No", "")))
        if m:
            max_no = max(max_no, int(m.group(1)))

        m2 = re.match(
            r"^(\d+)/(?:CLX|SI|SUNRISE|SR)/BOQ/",
            clean_text(r.get("BOQ No.", "")).upper(),
        )
        if m2:
            max_no = max(max_no, int(m2.group(1)))

    return max_no + 1


@st.cache_data(
    ttl=120,
    show_spinner=False,
)
def get_existing_saved_site_charger_pairs():
    return {
        (
            x["Site Name"].strip().lower(),
            normalize_charger_type(x["Charger Type"]),
        )
        for x in get_all_saved_boq()
        if x["Site Name"]
    }


@st.cache_data(
    ttl=120,
    show_spinner=False,
)
def fetch_query_site_options(
    exclude_saved=True,
):
    df = read_sheet_df(QUERY_SHEET)

    if df.empty:
        return [], {}

    def at(pos):
        if pos < len(df.columns):
            return df.columns[pos]
        return None

    (
        epc_c,
        charger_c,
        status_c,
        site_c,
        addr_c,
        province_c,
    ) = (
        at(1),
        at(2),
        at(3),
        at(5),
        at(6),
        at(8),
    )

    if site_c is None:
        return [], {}

    saved = (
        get_existing_saved_site_charger_pairs()
        if exclude_saved
        else set()
    )

    options = []
    data = {}

    for _, r in df.iterrows():
        site = clean_text(r.get(site_c, ""))
        charger = normalize_charger_type(r.get(charger_c, ""))
        status = (
            clean_text(r.get(status_c, "")).upper()
            if status_c
            else ""
        )

        if not site or status in {"DROP", "CANCEL"}:
            continue

        key = (
            site.lower(),
            charger,
        )

        if exclude_saved and key in saved:
            continue

        display = f"{site} ({charger})"

        if display not in data:
            data[display] = {
                "Site Name": site,
                "Charging Type": charger,
                "Address": (
                    clean_text(r.get(addr_c, ""))
                    if addr_c
                    else ""
                ),
                "Province": (
                    clean_text(r.get(province_c, ""))
                    if province_c
                    else ""
                ),
                "EPC Name": (
                    clean_text(r.get(epc_c, ""))
                    if epc_c
                    else ""
                ),
            }

            options.append(display)

    return options, data


def save_to_db_boq(
    site_name,
    charger_capacity,
    sub_total,
    grand_total,
    epc_name="-",
    template_name="Template BOQ Vgreen Lama",
    company="CLX",
    df_boq=None,
):
    try:
        sh = get_google_sheet_connection()

        if not sh:
            return None

        try:
            worksheet = sh.worksheet("DB BOQ")
        except Exception:
            worksheet = sh.add_worksheet(
                title="DB BOQ",
                rows=1000,
                cols=10,
            )

        ensure_db_boq_headers(worksheet)

        existing_rows = worksheet.get_all_values()

        max_no = 0
        for r in existing_rows[1:]:
            if not r:
                continue
            m = re.match(r"^(\d+)", clean_text(r[0]))
            if m:
                max_no = max(max_no, int(m.group(1)))
            if len(r) > 1:
                m2 = re.match(
                    r"^(\d+)/(?:CLX|SI|SUNRISE|SR)/BOQ/",
                    clean_text(r[1]).upper(),
                )
                if m2:
                    max_no = max(max_no, int(m2.group(1)))

        no_urut = max_no + 1

        boq_no = generate_boq_number(sequence_num=no_urut, company=company)
        items_json = serialize_boq_items_json(df_boq)

        new_row = [
            no_urut,
            boq_no,
            str(site_name),
            str(charger_capacity),
            sub_total,
            grand_total,
            str(epc_name),
            template_db_value(template_name),
            items_json,
        ]

        worksheet.append_row(new_row)

        return boq_no

    except Exception as e:
        st.error(f"❌ Gagal menyimpan ke DB BOQ: {e}")
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
    new_boq_no=None,
    df_boq=None,
):
    try:
        sh = get_google_sheet_connection()

        if not sh:
            return False

        worksheet = sh.worksheet("DB BOQ")
        ensure_db_boq_headers(worksheet)

        if new_boq_no:
            worksheet.update_cell(row_idx, 2, new_boq_no)

        worksheet.update_cell(row_idx, 3, new_site_name)
        worksheet.update_cell(row_idx, 4, charger_capacity)
        worksheet.update_cell(row_idx, 5, sub_total)
        worksheet.update_cell(row_idx, 6, grand_total)
        worksheet.update_cell(row_idx, 7, epc_name)

        if template_name:
            worksheet.update_cell(
                row_idx,
                8,
                template_db_value(template_name),
            )

        if df_boq is not None:
            worksheet.update_cell(
                row_idx,
                9,
                serialize_boq_items_json(df_boq),
            )

        update_google_sheet_summary(
            old_site_name,
            new_site_name,
            sub_total,
            grand_total,
        )

        return True

    except Exception as e:
        st.error(f"❌ Gagal memperbarui DB BOQ: {e}")
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

        try:
            worksheet = sh.worksheet("Sum Project")
        except Exception:
            return False

        data_sum = worksheet.get_all_values()

        for idx, row in enumerate(
            data_sum[1:],
            start=2,
        ):
            if (
                len(row) > 2
                and str(row[2]).strip().lower()
                == str(old_site_name).strip().lower()
            ):
                worksheet.update_cell(idx, 3, new_site_name)

                if len(row) >= 19:
                    worksheet.update_cell(idx, 18, sub_total)
                    worksheet.update_cell(idx, 19, grand_total)

                return True

    except Exception as e:
        st.caption(f"ℹ️ Info: Sheet 'Sum Project' belum ter-update ({e})")

    return False


def calculate_detail_total(
    quantity,
    unit_price,
):
    return (
        parse_qty_num(quantity)
        * parse_price(unit_price)
    )


# ==============================================================================
# FIXED BOQ CALCULATION ENGINE
# ==============================================================================
def recalculate_boq_totals(df_boq):
    if df_boq is None or df_boq.empty:
        return (df_boq, 0.0, 0.0, 0.0)

    df = df_boq.copy()

    required = [
        "NO", "Item", "Unit/Volume", "Satuan/Uom", "MERK", "UNIT PRICE", "TOTAL PRICE",
    ]

    for col in required:
        if col not in df.columns:
            df[col] = ""

    drop_idx = []
    for i, row in df.iterrows():
        no = clean_text(row.get("NO", "")).upper()
        item = clean_text(row.get("Item", "")).upper()
        unit = clean_text(row.get("Satuan/Uom", ""))
        merk = clean_text(row.get("MERK", ""))
        qty_val = parse_qty_num(row.get("Unit/Volume", 0))
        up_val = parse_price(row.get("UNIT PRICE", 0))

        if no == "I" and ("TOTAL" in item or "KESELURUHAN" in item or item == ""):
            drop_idx.append(i)
        elif not no and any(x in item for x in ["SUB TOTAL", "SUBTOTAL", "GRAND TOTAL", "VAT", "PPN"]):
            drop_idx.append(i)
        elif "TOTAL" in no and not item:
            drop_idx.append(i)
        elif not no and not item and not unit and not merk and qty_val == 0 and up_val == 0:
            drop_idx.append(i)

    if drop_idx:
        df = df.drop(index=drop_idx).reset_index(drop=True)

    alpha_parent_labels = {"A", "B", "C", "D", "E", "F", "G", "H", "J"}
    subsection_labels = {
        "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X",
        "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX",
    }

    line_totals = [0.0 for _ in range(len(df))]
    is_detail = [False for _ in range(len(df))]

    for pos, (_, row) in enumerate(df.iterrows()):
        unit_price = parse_price(row.get("UNIT PRICE", 0))
        qty = parse_qty_num(row.get("Unit/Volume", 0))

        if unit_price != 0:
            total = qty * unit_price
            line_totals[pos] = total
            is_detail[pos] = True

    all_nos = [clean_text(r.get("NO", "")).upper() for _, r in df.iterrows()]
    has_roman_ii = "II" in all_nos

    top_positions = []
    for pos, (_, row) in enumerate(df.iterrows()):
        no = clean_text(row.get("NO", "")).upper()
        if no in alpha_parent_labels:
            top_positions.append(pos)
        elif no == "I" and not has_roman_ii:
            top_positions.append(pos)

    subsection_positions = []
    for pos, (_, row) in enumerate(df.iterrows()):
        if pos in top_positions:
            continue
        no = clean_text(row.get("NO", "")).upper()
        item = clean_text(row.get("Item", "")).upper()

        if (
            no in subsection_labels
            and item not in {"TOTAL", "GRAND TOTAL", "SUB TOTAL", "SUBTOTAL"}
        ):
            subsection_positions.append(pos)

    display_totals = [None for _ in range(len(df))]
    for pos in range(len(df)):
        if is_detail[pos]:
            display_totals[pos] = line_totals[pos]

    for sub_index, sub_pos in enumerate(subsection_positions):
        next_sub_pos = (
            subsection_positions[sub_index + 1]
            if (sub_index + 1 < len(subsection_positions))
            else len(df)
        )
        next_top_after = [p for p in top_positions if p > sub_pos]
        next_top_pos = min(next_top_after) if next_top_after else len(df)
        end_pos = min(next_sub_pos, next_top_pos)

        subtotal = 0.0
        for i in range(sub_pos + 1, end_pos):
            if is_detail[i]:
                subtotal += line_totals[i]

        display_totals[sub_pos] = subtotal

    for top_index, top_pos in enumerate(top_positions):
        row = df.iloc[top_pos]
        unit_price = parse_price(row.get("UNIT PRICE", 0))

        if unit_price != 0:
            display_totals[top_pos] = line_totals[top_pos]
            continue

        next_top_pos = (
            top_positions[top_index + 1]
            if (top_index + 1 < len(top_positions))
            else len(df)
        )

        child_subsections = [
            p for p in subsection_positions
            if (top_pos < p < next_top_pos)
        ]

        if child_subsections:
            first_subsection = min(child_subsections)
            direct_subtotal = sum(
                line_totals[i]
                for i in range(top_pos + 1, first_subsection)
                if is_detail[i]
            )
            sub_totals_sum = sum(
                display_totals[sub_p] or 0.0
                for sub_p in child_subsections
            )
            display_totals[top_pos] = direct_subtotal + sub_totals_sum
            continue

        subtotal = 0.0
        for i in range(top_pos + 1, next_top_pos):
            if i in subsection_positions:
                continue
            if is_detail[i]:
                subtotal += line_totals[i]

        display_totals[top_pos] = subtotal

    for pos, (_, row) in enumerate(df.iterrows()):
        if display_totals[pos] is not None:
            continue
        existing = parse_price(row.get("TOTAL PRICE", 0))
        if existing != 0:
            display_totals[pos] = existing
        else:
            display_totals[pos] = None

    df["TOTAL PRICE"] = display_totals

    subtotal = sum(
        line_totals[pos]
        for pos in range(len(df))
        if is_detail[pos]
    )
    vat = subtotal * 0.11
    grand = subtotal + vat

    return (df, subtotal, vat, grand)


def _old_template_offsets(region):
    return {
        "JAVA": 0,
        "SUMATERA": 8,
        "BALI NUSATENGGARA": 16,
        "KALIMANTAN": 25,
        "SULAWESI": 33,
    }.get(
        region,
        0,
    )


def _first_series_by_header(
    data,
    source,
):
    if (
        source is None
        or data is None
        or data.empty
    ):
        return None

    try:
        positions = [
            i
            for i, col in enumerate(data.columns)
            if col == source
        ]

        if positions:
            return data.iloc[:, positions[0]]
    except Exception:
        pass

    try:
        value = data[source]
        if isinstance(value, pd.DataFrame):
            return value.iloc[:, 0]
        return value
    except Exception:
        return None


def _standardize_template_df(
    raw,
    header_row=None,
):
    if raw.empty:
        return pd.DataFrame(
            columns=[
                "NO",
                "Item",
                "Unit/Volume",
                "Satuan/Uom",
                "MERK",
                "UNIT PRICE",
                "TOTAL PRICE",
            ]
        )

    raw = raw.copy()

    if header_row is None:
        header_row = detect_template_header(raw)

    headers = _normalize_template_headers(
        raw.iloc[header_row].tolist()
    )

    data = raw.iloc[header_row + 1:].copy()
    data.columns = headers

    data = data.loc[
        :,
        [c != "" for c in data.columns],
    ]

    mapping = build_new_template_column_mapping(data.columns)

    out = pd.DataFrame(index=data.index)

    for target, source in mapping.items():
        if source is None:
            continue

        series = _first_series_by_header(data, source)
        if series is not None:
            out[target] = series.to_numpy()

    rename = {
        "NO": "NO",
        "ITEM": "Item",
        "UNIT_VOLUME": "Unit/Volume",
        "SATUAN": "Satuan/Uom",
        "MERK": "MERK",
        "UNIT_PRICE": "UNIT PRICE",
        "TOTAL_PRICE": "TOTAL PRICE",
    }

    out = out.rename(columns=rename)

    for c in [
        "NO",
        "Item",
        "Unit/Volume",
        "Satuan/Uom",
        "MERK",
        "UNIT PRICE",
        "TOTAL PRICE",
    ]:
        if c not in out.columns:
            out[c] = ""

    return out[
        [
            "NO",
            "Item",
            "Unit/Volume",
            "Satuan/Uom",
            "MERK",
            "UNIT PRICE",
            "TOTAL PRICE",
        ]
    ].reset_index(drop=True)


def _excel_number_from_value(
    value,
    cell_map=None,
    resolving=None,
):
    if (
        value is None
        or (
            isinstance(value, float)
            and pd.isna(value)
        )
    ):
        return None

    if (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
    ):
        try:
            return float(value)
        except Exception:
            return None

    s = str(value).strip()

    if not s or s.upper() in ["#VALUE!", "#REF!", "#N/A"]:
        return None

    if not s.startswith("="):
        try:
            return parse_qty_num(s)
        except Exception:
            return None

    expr = (
        s[1:]
        .strip()
        .replace("$", "")
    )

    cell_map = cell_map or {}
    resolving = set(resolving or set())

    ref_re = re.compile(
        r"(?<![A-Za-z0-9_])"
        r"([A-Z]{1,3})(\d+)"
        r"(?![A-Za-z0-9_])",
        re.I,
    )

    def replace_ref(match):
        ref = (
            match.group(1).upper()
            + match.group(2)
        )

        if ref in resolving:
            raise ValueError("Circular formula reference")

        if ref not in cell_map:
            raise KeyError(ref)

        resolved = _excel_number_from_value(
            cell_map[ref],
            cell_map,
            resolving | {ref},
        )

        if resolved is None:
            raise ValueError(f"Non-numeric cell {ref}")

        return str(resolved)

    try:
        expr = ref_re.sub(replace_ref, expr)

        if not re.fullmatch(r"[0-9eE+\-*/().\s]+", expr):
            return None

        tree = ast.parse(expr, mode="eval")

        allowed = (
            ast.Expression,
            ast.BinOp,
            ast.UnaryOp,
            ast.Add,
            ast.Sub,
            ast.Mult,
            ast.Div,
            ast.Pow,
            ast.USub,
            ast.UAdd,
            ast.Constant,
            ast.Load,
        )

        for node in ast.walk(tree):
            if not isinstance(node, allowed):
                return None
            if (
                isinstance(node, ast.Constant)
                and not isinstance(node.value, (int, float))
            ):
                return None

        return float(
            eval(
                compile(tree, "<boq-formula>", "eval"),
                {"__builtins__": {}},
                {},
            )
        )
    except Exception:
        return None


def _build_excel_cell_map(data_raw, header_row):
    cell_map = {}

    for raw_i, row in data_raw.iterrows():
        excel_row = raw_i + header_row + 2

        for j, value in enumerate(row.tolist(), start=1):
            letters = ""
            n = j

            while n:
                n, rem = divmod(n - 1, 26)
                letters = chr(65 + rem) + letters

            cell_map[f"{letters}{excel_row}"] = value

    return cell_map


def _restore_template_formula_values(df, raw):
    if df is None or df.empty or raw is None or raw.empty:
        return df

    out = df.copy()
    header_row = detect_template_header(raw)

    if header_row >= len(raw):
        return out

    headers = _normalize_template_headers(
        raw.iloc[header_row].tolist()
    )

    data_raw = raw.iloc[header_row + 1:].reset_index(drop=True)

    if data_raw.empty:
        return out

    mapping = build_new_template_column_mapping(headers)
    cell_map = _build_excel_cell_map(data_raw, header_row)

    def raw_column_position(target):
        source = mapping.get(target)
        if source is None or source not in headers:
            return None
        return headers.index(source)

    qty_pos = raw_column_position("UNIT_VOLUME")
    price_pos = raw_column_position("UNIT_PRICE")
    total_pos = raw_column_position("TOTAL_PRICE")

    row_count = min(len(out), len(data_raw))

    for i in range(row_count):
        if qty_pos is not None:
            raw_qty = data_raw.iloc[i, qty_pos]
            resolved_qty = _excel_number_from_value(raw_qty, cell_map)
            if resolved_qty is not None:
                out.at[i, "Unit/Volume"] = resolved_qty

        if price_pos is not None:
            raw_price = data_raw.iloc[i, price_pos]
            resolved_price = _excel_number_from_value(raw_price, cell_map)
            if resolved_price is not None:
                out.at[i, "UNIT PRICE"] = resolved_price

        if total_pos is not None:
            raw_total = data_raw.iloc[i, total_pos]
            resolved_total = _excel_number_from_value(raw_total, cell_map)
            if resolved_total is not None:
                out.at[i, "TOTAL PRICE"] = resolved_total

    return out


# ==============================================================================
# LOAD BOQ
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
    filename = TEMPLATE_OPTIONS.get(template_name, template_name)
    path = get_template_path(filename)

    if not os.path.exists(path):
        raise FileNotFoundError(f"Template tidak ditemukan: {path}")

    c = normalize_charger_type(charging_type)
    region = map_to_standard_province(province_str)

    if template_name == "Template BOQ Vgreen Lama":
        sheet_map = {
            "DC20": "DC20",
            "DC30": "DC30",
            "DC60": "DC60",
            "DC120": "DC120",
            "6S1P": "6S1P",
            "12S1P": "12S1P",
            "12S3P": "12S3P",
            "7KW": "7KW",
            "22KW": "22KW",
        }

        sheet = sheet_map.get(c, c)
        raw = pd.read_excel(path, sheet_name=sheet, header=None)
        off = _old_template_offsets(region)

        if off and len(raw) > off:
            raw = raw.iloc[off:].reset_index(drop=True)

        header_row = detect_template_header(raw)

        if header_row == 0 and raw.shape[1] >= 7:
            first = [
                clean_text(x).upper()
                for x in raw.iloc[0].tolist()
            ]

            if not any("ITEM" in x for x in first):
                raw.columns = [
                    "NO",
                    "Item",
                    "Unit/Volume",
                    "Satuan/Uom",
                    "MERK",
                    "UNIT PRICE",
                    "TOTAL PRICE",
                ] + list(raw.columns[7:])

                df = raw.iloc[:, :7].copy()
                df.columns = [
                    "NO",
                    "Item",
                    "Unit/Volume",
                    "Satuan/Uom",
                    "MERK",
                    "UNIT PRICE",
                    "TOTAL PRICE",
                ]
            else:
                df = _standardize_template_df(raw, header_row)
        else:
            df = _standardize_template_df(raw, header_row)
    else:
        expected_sheet = get_new_template_sheet_name(c, province_str)
        xl = pd.ExcelFile(path)
        available_sheets = list(xl.sheet_names)

        sheet = next(
            (x for x in available_sheets if x == expected_sheet),
            None,
        )

        if sheet is None:
            sheet = next(
                (x for x in available_sheets if x.upper() == expected_sheet.upper()),
                None,
            )

        if sheet is None:
            raise ValueError(
                "Regional template mismatch.\n"
                f"Province source : '{province_str}'\n"
                f"Mapped region   : '{region}'\n"
                f"Charging type   : '{c}'\n"
                f"Required sheet  : '{expected_sheet}'\n"
                f"Available sheets: {', '.join(available_sheets)}"
            )

        raw = pd.read_excel(path, sheet_name=sheet, header=None)
        df = _standardize_template_df(raw)
        df = _restore_template_formula_values(df, raw)

    for col in [
        "NO",
        "Item",
        "Unit/Volume",
        "Satuan/Uom",
        "MERK",
    ]:
        df[col] = df[col].apply(clean_text)

    df["UNIT PRICE"] = df["UNIT PRICE"].apply(parse_price)
    df["TOTAL PRICE"] = df["TOTAL PRICE"].apply(parse_price)

    df, _, _, _ = recalculate_boq_totals(df)

    return (
        df.reset_index(drop=True),
        sheet,
        region,
    )


# ==============================================================================
# SECTION EDITOR
# ==============================================================================
def get_section_indices(df_boq, section_no):
    if (
        df_boq is None
        or df_boq.empty
        or "NO" not in df_boq.columns
    ):
        return []

    parent_labels = {
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J"
    }

    target = clean_text(section_no).upper()
    indices = []
    active = False

    for i, v in df_boq["NO"].items():
        no = clean_text(v).upper()

        if no == target:
            active = True
            indices.append(i)
            continue

        if active and no in parent_labels:
            break

        if active:
            indices.append(i)

    return indices


def get_editable_sections(charger_type):
    c = normalize_charger_type(charger_type)

    if c in {"6S1P", "12S1P", "12S3P"}:
        return ["A", "D"]

    if c in {"DC20", "DC30", "DC60"}:
        return ["A", "G"]

    return []


def _editor_signature(section_df):
    payload = (
        section_df[
            [
                "NO",
                "Item",
                "Unit/Volume",
                "Satuan/Uom",
                "MERK",
                "UNIT PRICE",
            ]
        ]
        .fillna("")
        .astype(str)
        .to_json(orient="records")
    )

    return hashlib.sha1(payload.encode("utf-8")).hexdigest()


def edit_boq_sections(
    df_boq,
    charger_type,
    widget_prefix,
):
    if df_boq is None or df_boq.empty:
        return df_boq

    editable_sections = get_editable_sections(charger_type)

    if not editable_sections:
        return df_boq

    df = df_boq.copy()

    for col in [
        "Unit/Volume",
        "UNIT PRICE",
        "TOTAL PRICE",
    ]:
        if col in df.columns:
            df[col] = df[col].astype(object)

    st.markdown("### ✏️ Edit Qty / Volume BOQ")
    st.info(
        "💡 Qty / Volume awal mengikuti template BOQ yang dipilih. "
        "Hanya Qty / Volume pada Point yang ditentukan yang dapat diubah."
    )

    for section_no in editable_sections:
        section_indices = get_section_indices(df, section_no)

        if not section_indices:
            continue

        section_df = df.loc[section_indices].copy()

        st.markdown(f"#### Point {section_no}")

        editor_df = section_df[
            [
                "NO",
                "Item",
                "Unit/Volume",
                "Satuan/Uom",
                "MERK",
                "UNIT PRICE",
                "TOTAL PRICE",
            ]
        ].copy()

        editor_df["Unit/Volume"] = editor_df["Unit/Volume"].apply(clean_text)
        editor_df["Satuan/Uom"] = editor_df["Satuan/Uom"].apply(clean_text)
        editor_df["MERK"] = editor_df["MERK"].apply(clean_text)
        editor_df["UNIT PRICE"] = editor_df["UNIT PRICE"].apply(parse_price).astype(float)
        editor_df["TOTAL PRICE"] = editor_df["TOTAL PRICE"].apply(parse_price).astype(float)

        editor_key = (
            f"boq_qty_editor_v15_"
            f"{hashlib.sha1(str(widget_prefix).encode()).hexdigest()[:12]}_"
            f"{section_no}"
        )

        signature = _editor_signature(section_df)
        signature_key = f"{editor_key}__template_signature"
        old_signature = st.session_state.get(signature_key)

        if old_signature != signature:
            st.session_state.pop(editor_key, None)
            st.session_state[signature_key] = signature

        column_config = {
            "NO": st.column_config.TextColumn("NO", disabled=True),
            "Item": st.column_config.TextColumn("Item", disabled=True),
            "Unit/Volume": st.column_config.TextColumn("Qty / Volume", disabled=False),
            "Satuan/Uom": st.column_config.TextColumn("Satuan", disabled=True),
            "MERK": st.column_config.TextColumn("MERK", disabled=True),
            "UNIT PRICE": st.column_config.NumberColumn("UNIT PRICE", format="Rp %.0f", disabled=True),
            "TOTAL PRICE": st.column_config.NumberColumn("TOTAL PRICE", format="Rp %.0f", disabled=True),
        }

        edited_df = st.data_editor(
            editor_df,
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            column_config=column_config,
            disabled=[
                "NO",
                "Item",
                "Satuan/Uom",
                "MERK",
                "UNIT PRICE",
                "TOTAL PRICE",
            ],
            key=editor_key,
        )

        for position, idx in enumerate(section_indices):
            if position >= len(edited_df):
                continue

            raw_qty = edited_df.iloc[position]["Unit/Volume"]
            unit_price = parse_price(df.loc[idx, "UNIT PRICE"])

            if unit_price == 0:
                df.loc[idx, "Unit/Volume"] = clean_text(raw_qty)
                continue

            new_qty = parse_qty_num(raw_qty)
            df.loc[idx, "Unit/Volume"] = new_qty
            df.loc[idx, "TOTAL PRICE"] = calculate_detail_total(new_qty, unit_price)

    df, _, _, _ = recalculate_boq_totals(df)
    return df


# ==============================================================================
# BOQ PDF (MODERN CORPORATE LAYOUT — FIT ONE PAGE A4, CLX & SUNRISE THEMES)
# ==============================================================================
class _BoqNumberedCanvas(pdf_canvas.Canvas):
    def __init__(self, *args, **kwargs):
        self.company_name = kwargs.pop(
            "company_name", "PT. Connectivity Leads eXcellence"
        )
        self.doc_ref = kwargs.pop("doc_ref", "")
        self.footer_line_hex = kwargs.pop("footer_line_hex", "#CBD5E1")
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_footer(total_pages)
            super().showPage()
        super().save()

    def _draw_footer(self, total_pages):
        page_width, _ = A4
        margin_x = 22
        y_line = 20
        y_text = 10

        self.saveState()
        self.setStrokeColor(colors.HexColor(self.footer_line_hex))
        self.setLineWidth(0.9)
        self.line(margin_x, y_line, page_width - margin_x, y_line)

        self.setFont("Helvetica", 7)
        self.setFillColor(colors.HexColor("#64748B"))
        left_label = f"{self.company_name}  |  Bill of Quantity (BOQ)"
        if self.doc_ref and self.doc_ref != "-":
            left_label += f"  |  Ref: {self.doc_ref}"
        self.drawString(margin_x, y_text, left_label)
        self.drawRightString(
            page_width - margin_x,
            y_text,
            f"Halaman {self._pageNumber} dari {total_pages}",
        )
        self.restoreState()


def generate_boq_pdf(
    site_name,
    site_location,
    charger_capacity,
    region,
    df_boq,
    sub_total,
    vat,
    grand_total,
    boq_no="-",
    epc_name="-",
    date_obj=None,
    company=None,
):
    buffer = io.BytesIO()
    page_w, page_h = A4
    margin_x = 22
    top_margin = 16
    bottom_margin = 26

    avail_w = page_w - (2 * margin_x)
    avail_h = page_h - top_margin - bottom_margin
    safe_target_h = avail_h - 14.0

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=margin_x,
        leftMargin=margin_x,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
    )

    dt = date_obj or datetime.datetime.now()
    date_str = dt.strftime("%d %B %Y")

    resolved_company = company or detect_company_from_context(boq_no=boq_no, epc_name=epc_name)
    company_key = normalize_company_key(resolved_company)
    cfg = get_company_config(company_key)

    PRIMARY_BG = colors.HexColor(cfg["primary_hex"])
    BORDER_PRIMARY = colors.HexColor(cfg["border_primary_hex"])
    ACCENT_COLOR = colors.HexColor(cfg["accent_hex"])
    TITLE_TEXT_COLOR = colors.HexColor(cfg["title_text_hex"])
    SECTION_BG = colors.HexColor(cfg["section_bg_hex"])
    SUBSECTION_BG = colors.HexColor(cfg["subsection_bg_hex"])
    ZEBRA_BG = colors.HexColor(cfg["zebra_bg_hex"])
    CARD_BG = colors.HexColor(cfg["card_bg_hex"])
    HEADER_TEXT_COLOR = colors.HexColor(cfg["header_text_hex"])
    SECTION_TEXT_COLOR = colors.HexColor(cfg["section_text_hex"])
    GRAND_TEXT_COLOR = colors.HexColor(cfg["grand_text_hex"])
    BORDER_LIGHT = colors.HexColor("#CBD5E1")
    TEXT_DARK = colors.HexColor("#0F172A")
    TEXT_MUTED = colors.HexColor("#475569")

    styles = getSampleStyleSheet()
    pdf_df, _, _, _ = recalculate_boq_totals(df_boq)

    alpha_parents = {"A", "B", "C", "D", "E", "F", "G", "H", "J"}
    roman_subsections = {
        "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X",
        "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX",
    }
    all_nos = [clean_text(x).upper() for x in pdf_df["NO"].tolist()]
    has_roman_ii = "II" in all_nos

    logo_path = get_logo_path(company_key)
    col_widths = [24, 217, 36, 38, 86, 75, 75]

    def _build_scaled_elements(scale):
        def sz(base, floor=5.4):
            return max(floor, round(base * scale, 2))

        def pad(base, floor=0.7):
            return max(floor, round(base * scale, 2))

        company_style = ParagraphStyle(
            f"DocCompanyTitle_{company_key}_{scale}",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=sz(10.5, 8.5),
            leading=sz(12.5, 10.0),
            textColor=TITLE_TEXT_COLOR,
        )
        company_sub_style = ParagraphStyle(
            f"DocCompanySub_{company_key}_{scale}",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=sz(7.8, 6.5),
            leading=sz(9.5, 8.0),
            textColor=TEXT_MUTED,
        )
        doc_badge_style = ParagraphStyle(
            f"DocBadgeTitle_{company_key}_{scale}",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=sz(13.0, 10.0),
            leading=sz(15.0, 11.5),
            alignment=2,
            textColor=TITLE_TEXT_COLOR,
        )
        doc_meta_right_style = ParagraphStyle(
            f"DocMetaRight_{company_key}_{scale}",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=sz(8.0, 6.6),
            leading=sz(10.0, 8.2),
            alignment=2,
            textColor=TEXT_DARK,
        )
        info_label_style = ParagraphStyle(
            f"DocInfoLabel_{company_key}_{scale}",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=sz(7.6, 6.3),
            leading=sz(9.5, 7.8),
            textColor=TEXT_MUTED,
        )
        info_val_style = ParagraphStyle(
            f"DocInfoVal_{company_key}_{scale}",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=sz(8.0, 6.5),
            leading=sz(9.8, 8.0),
            textColor=TEXT_DARK,
        )

        th_left = ParagraphStyle(
            f"BoqThLeft_{company_key}_{scale}",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=sz(7.6, 6.0),
            leading=sz(9.2, 7.2),
            textColor=HEADER_TEXT_COLOR,
        )
        th_center = ParagraphStyle(
            f"BoqThCenter_{company_key}_{scale}", parent=th_left, alignment=1
        )
        th_right = ParagraphStyle(
            f"BoqThRight_{company_key}_{scale}", parent=th_left, alignment=2
        )

        cell_left = ParagraphStyle(
            f"BoqCellLeft_{company_key}_{scale}",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=sz(7.5, 5.7),
            leading=sz(9.1, 6.8),
            textColor=TEXT_DARK,
        )
        cell_center = ParagraphStyle(
            f"BoqCellCenter_{company_key}_{scale}", parent=cell_left, alignment=1
        )
        cell_right = ParagraphStyle(
            f"BoqCellRight_{company_key}_{scale}", parent=cell_left, alignment=2
        )

        sec_left = ParagraphStyle(
            f"BoqSecLeft_{company_key}_{scale}",
            parent=cell_left,
            fontName="Helvetica-Bold",
            fontSize=sz(7.9, 6.0),
            leading=sz(9.6, 7.1),
            textColor=SECTION_TEXT_COLOR,
        )
        sec_center = ParagraphStyle(
            f"BoqSecCenter_{company_key}_{scale}", parent=sec_left, alignment=1
        )
        sec_right = ParagraphStyle(
            f"BoqSecRight_{company_key}_{scale}", parent=sec_left, alignment=2
        )

        subsec_left = ParagraphStyle(
            f"BoqSubSecLeft_{company_key}_{scale}",
            parent=cell_left,
            fontName="Helvetica-Bold",
            fontSize=sz(7.6, 5.8),
            leading=sz(9.2, 6.9),
            textColor=TEXT_DARK,
        )
        subsec_center = ParagraphStyle(
            f"BoqSubSecCenter_{company_key}_{scale}", parent=subsec_left, alignment=1
        )
        subsec_right = ParagraphStyle(
            f"BoqSubSecRight_{company_key}_{scale}", parent=subsec_left, alignment=2
        )

        logo_block = []
        if logo_path and os.path.exists(logo_path):
            try:
                from reportlab.lib.utils import ImageReader
                ir = ImageReader(logo_path)
                iw, ih = ir.getSize()
                if iw and ih:
                    target_h = max(18.0, 26.0 * scale)
                    target_w = min(115.0, target_h * (float(iw) / float(ih)))
                    img = Image(logo_path, width=target_w, height=target_h)
                    img.hAlign = "LEFT"
                    logo_block.append(img)
                    logo_block.append(Spacer(1, pad(2, 1)))
            except Exception:
                pass

        logo_block.append(Paragraph(_esc_pdf_text(cfg["name"]), company_style))
        logo_block.append(
            Paragraph(_esc_pdf_text(cfg["subtitle"]), company_sub_style)
        )

        right_header_block = [
            Paragraph("BILL OF QUANTITY", doc_badge_style),
            Spacer(1, pad(1.5, 0.5)),
            Paragraph(f"<b>No. Dokumen:</b> {_esc_pdf_text(boq_no or '-')}", doc_meta_right_style),
            Paragraph(f"<b>Tanggal:</b> {_esc_pdf_text(date_str)}", doc_meta_right_style),
            Paragraph(
                f"<font color='{cfg['badge_text_hex']}'><b>CHARGING WORK {_esc_pdf_text(charger_capacity)}</b></font>",
                doc_meta_right_style,
            ),
        ]

        header_table = Table(
            [[logo_block, right_header_block]],
            colWidths=[295, 256],
        )
        header_table.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), pad(5, 2.5)),
                    ("LINEBELOW", (0, 0), (-1, 0), 2.2, PRIMARY_BG),
                ]
            )
        )

        meta_rows = [
            [
                Paragraph("Site Name", info_label_style),
                Paragraph(f": {_esc_pdf_text(site_name)}", info_val_style),
                Paragraph("Charger Capacity", info_label_style),
                Paragraph(f": {_esc_pdf_text(charger_capacity)}", info_val_style),
            ],
            [
                Paragraph("Site Location", info_label_style),
                Paragraph(f": {_esc_pdf_text(site_location)}", info_val_style),
                Paragraph("Region / Area", info_label_style),
                Paragraph(f": {_esc_pdf_text(region)} ISLAND", info_val_style),
            ],
            [
                Paragraph("EPC / Partner", info_label_style),
                Paragraph(f": {_esc_pdf_text(epc_name or '-')}", info_val_style),
                Paragraph("Plan Reference", info_label_style),
                Paragraph(": NEW PLAN BOQ VGREEN", info_val_style),
            ],
        ]

        meta_table = Table(meta_rows, colWidths=[70, 228, 82, 171])
        meta_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), CARD_BG),
                    ("LINEBEFORE", (0, 0), (0, -1), 3.4, ACCENT_COLOR),
                    ("BOX", (0, 0), (-1, -1), 0.5, BORDER_LIGHT),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("TOPPADDING", (0, 0), (-1, -1), pad(2.5, 1.2)),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), pad(2.5, 1.2)),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )

        table_data = [
            [
                Paragraph("NO", th_center),
                Paragraph("ITEM DESCRIPTION", th_left),
                Paragraph("QTY", th_right),
                Paragraph("UNIT", th_center),
                Paragraph("BRAND / MERK", th_left),
                Paragraph("UNIT PRICE", th_right),
                Paragraph("TOTAL PRICE", th_right),
            ]
        ]

        dynamic_styles = [
            ("BACKGROUND", (0, 0), (-1, 0), PRIMARY_BG),
            ("TOPPADDING", (0, 0), (-1, 0), pad(4.0, 2.0)),
            ("BOTTOMPADDING", (0, 0), (-1, 0), pad(4.0, 2.0)),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]

        detail_zebra_counter = 0

        for _, row in pdf_df.iterrows():
            no_raw = clean_text(row.get("NO", "")).upper()
            item_raw = clean_text(row.get("Item", ""))
            qty_raw = row.get("Unit/Volume", "")
            unit_raw = clean_text(row.get("Satuan/Uom", ""))
            merk_raw = clean_text(row.get("MERK", ""))

            up = parse_price(row.get("UNIT PRICE", 0))
            tp = parse_price(row.get("TOTAL PRICE", 0))

            is_parent = (no_raw in alpha_parents) or (no_raw == "I" and not has_roman_ii)
            is_sub = (no_raw in roman_subsections) and not is_parent

            row_idx = len(table_data)

            if is_parent:
                detail_zebra_counter = 0
                up_str = format_currency(up) if up != 0 else ""
                tp_str = format_currency(tp) if tp != 0 else "-"
                qty_str = _format_qty_clean(qty_raw) if parse_qty_num(qty_raw) != 0 else ""

                table_data.append(
                    [
                        Paragraph(_esc_pdf_text(no_raw), sec_center),
                        Paragraph(_esc_pdf_text(item_raw), sec_left),
                        Paragraph(_esc_pdf_text(qty_str), sec_right),
                        Paragraph(_esc_pdf_text(unit_raw), sec_center),
                        Paragraph(_esc_pdf_text(merk_raw), sec_left),
                        Paragraph(up_str, sec_right),
                        Paragraph(tp_str, sec_right),
                    ]
                )
                dynamic_styles.extend(
                    [
                        ("BACKGROUND", (0, row_idx), (-1, row_idx), SECTION_BG),
                        ("TOPPADDING", (0, row_idx), (-1, row_idx), pad(3.0, 1.3)),
                        ("BOTTOMPADDING", (0, row_idx), (-1, row_idx), pad(3.0, 1.3)),
                        ("LINEABOVE", (0, row_idx), (-1, row_idx), 0.65, BORDER_PRIMARY),
                        ("LINEBELOW", (0, row_idx), (-1, row_idx), 0.45, BORDER_LIGHT),
                    ]
                )

            elif is_sub:
                detail_zebra_counter = 0
                tp_str = format_currency(tp) if tp != 0 else "-"
                table_data.append(
                    [
                        Paragraph(_esc_pdf_text(no_raw), subsec_center),
                        Paragraph(_esc_pdf_text(item_raw), subsec_left),
                        Paragraph("", subsec_right),
                        Paragraph("", subsec_center),
                        Paragraph("", subsec_left),
                        Paragraph("", subsec_right),
                        Paragraph(tp_str, subsec_right),
                    ]
                )
                dynamic_styles.extend(
                    [
                        ("BACKGROUND", (0, row_idx), (-1, row_idx), SUBSECTION_BG),
                        ("TOPPADDING", (0, row_idx), (-1, row_idx), pad(2.5, 1.1)),
                        ("BOTTOMPADDING", (0, row_idx), (-1, row_idx), pad(2.5, 1.1)),
                    ]
                )

            else:
                up_str = format_currency(up) if up != 0 else "-"
                tp_str = format_currency(tp) if tp != 0 else "-"
                qty_str = _format_qty_clean(qty_raw)

                table_data.append(
                    [
                        Paragraph(_esc_pdf_text(no_raw), cell_center),
                        Paragraph(_esc_pdf_text(item_raw), cell_left),
                        Paragraph(_esc_pdf_text(qty_str), cell_right),
                        Paragraph(_esc_pdf_text(unit_raw), cell_center),
                        Paragraph(_esc_pdf_text(merk_raw), cell_left),
                        Paragraph(up_str, cell_right),
                        Paragraph(tp_str, cell_right),
                    ]
                )
                bg = ZEBRA_BG if (detail_zebra_counter % 2 == 1) else colors.white
                detail_zebra_counter += 1
                dynamic_styles.extend(
                    [
                        ("BACKGROUND", (0, row_idx), (-1, row_idx), bg),
                        ("TOPPADDING", (0, row_idx), (-1, row_idx), pad(2.2, 0.85)),
                        ("BOTTOMPADDING", (0, row_idx), (-1, row_idx), pad(2.2, 0.85)),
                    ]
                )

        last_item_row = len(table_data) - 1

        sum_label_style = ParagraphStyle(
            f"BoqSumLabel_{company_key}_{scale}",
            parent=cell_right,
            fontName="Helvetica-Bold",
            fontSize=sz(8.0, 6.5),
            leading=sz(9.6, 7.6),
            textColor=TEXT_DARK,
        )
        grand_label_style = ParagraphStyle(
            f"BoqGrandLabel_{company_key}_{scale}",
            parent=cell_right,
            fontName="Helvetica-Bold",
            fontSize=sz(8.8, 7.2),
            leading=sz(10.5, 8.4),
            textColor=GRAND_TEXT_COLOR,
        )

        table_data.extend(
            [
                [
                    Paragraph("Sub Total (Exc. VAT)", sum_label_style),
                    "", "", "", "", "",
                    Paragraph(format_currency(sub_total), sum_label_style),
                ],
                [
                    Paragraph("VAT / PPN 11%", sum_label_style),
                    "", "", "", "", "",
                    Paragraph(format_currency(vat), sum_label_style),
                ],
                [
                    Paragraph("TOTAL CONTRACTOR PRICE (INC. VAT 11%)", grand_label_style),
                    "", "", "", "", "",
                    Paragraph(format_currency(grand_total), grand_label_style),
                ],
            ]
        )

        dynamic_styles.extend(
            [
                ("GRID", (0, 0), (-1, last_item_row), 0.35, BORDER_LIGHT),
                ("BOX", (0, 0), (-1, -1), 0.85, BORDER_PRIMARY),
                ("SPAN", (0, -3), (5, -3)),
                ("BACKGROUND", (0, -3), (-1, -3), CARD_BG),
                ("LINEABOVE", (0, -3), (-1, -3), 0.9, BORDER_PRIMARY),
                ("TOPPADDING", (0, -3), (-1, -2), pad(3.0, 1.5)),
                ("BOTTOMPADDING", (0, -3), (-1, -2), pad(3.0, 1.5)),
                ("SPAN", (0, -2), (5, -2)),
                ("BACKGROUND", (0, -2), (-1, -2), CARD_BG),
                ("LINEBELOW", (0, -2), (-1, -2), 0.45, BORDER_LIGHT),
                ("SPAN", (0, -1), (5, -1)),
                ("BACKGROUND", (0, -1), (-1, -1), PRIMARY_BG),
                ("TOPPADDING", (0, -1), (-1, -1), pad(4.5, 2.5)),
                ("BOTTOMPADDING", (0, -1), (-1, -1), pad(4.5, 2.5)),
            ]
        )

        t = Table(
            table_data,
            colWidths=col_widths,
            repeatRows=1,
        )
        t.setStyle(TableStyle(dynamic_styles))

        sp1 = Spacer(1, pad(6.0, 3.0))
        sp2 = Spacer(1, pad(6.0, 3.0))
        built = [header_table, sp1, meta_table, sp2, t]

        total_h = 0.0
        for item in built:
            _, ih = item.wrap(avail_w, avail_h)
            total_h += ih

        return built, total_h

    chosen_elements = None
    chosen_h = 0.0
    for sc in (1.06, 1.00, 0.94, 0.88, 0.82, 0.76, 0.70, 0.65, 0.60):
        candidate_elements, measured_h = _build_scaled_elements(sc)
        chosen_elements = candidate_elements
        chosen_h = measured_h
        if measured_h <= safe_target_h:
            break

    if chosen_h > safe_target_h:
        elements = [
            KeepInFrame(
                avail_w,
                safe_target_h,
                chosen_elements,
                mode="shrink",
                hAlign="CENTER",
                vAlign="TOP",
            )
        ]
    else:
        elements = chosen_elements

    footer_line = cfg["border_primary_hex"] if company_key == "Sunrise" else "#CBD5E1"

    doc.build(
        elements,
        canvasmaker=lambda *args, **kwargs: _BoqNumberedCanvas(
            *args,
            company_name=cfg["name"],
            doc_ref=clean_text(boq_no),
            footer_line_hex=footer_line,
            **kwargs,
        ),
    )
    buffer.seek(0)

    return buffer.getvalue()


# ==============================================================================
# PERIOD & BOQ MS MODULES
# ==============================================================================
def _period_text(date_obj):
    months = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sept", "Oct", "Nov", "Dec",
    ]

    d = (
        date_obj.date()
        if isinstance(date_obj, datetime.datetime)
        else date_obj
    )

    if d.day >= 16:
        m1 = d.month
        y1 = d.year
        m2 = m1 + 1
        y2 = y1

        if m2 == 13:
            m2 = 1
            y2 = y1 + 1
    else:
        m2 = d.month
        y2 = d.year
        m1 = m2 - 1
        y1 = y2

        if m1 == 0:
            m1 = 12
            y1 = y2 - 1

    return f"{months[m1-1]}-{months[m2-1]} {y2}"


def _ensure_boq_ms_headers(ws):
    headers = BOQ_MS_HEADERS + BOQ_MS_HELPER_HEADERS
    try:
        vals = ws.get_all_values()
        first = vals[0] if vals else []
        if first[:9] != BOQ_MS_HEADERS:
            ws.update("A1:I1", [BOQ_MS_HEADERS])
        vals = ws.get_all_values()
        first = vals[0] if vals else []
        full = BOQ_MS_HEADERS + BOQ_MS_HELPER_HEADERS
        for col_idx, header in enumerate(full[9:], start=10):
            if len(first) < col_idx or str(first[col_idx - 1]).strip() != header:
                ws.update_cell(1, col_idx, header)
    except Exception:
        ws.update("A1:O1", [headers])
    return headers


@st.cache_data(
    ttl=120,
    show_spinner=False,
)
def _read_boq_ms_records():
    df = read_sheet_df(DB_BOQ_MS_SHEET)

    if df.empty:
        return []

    cols = list(df.columns)

    def get(row, idx, default=""):
        c = cols[idx] if idx < len(cols) else None
        return row.get(c, default) if c else default

    out = []
    for i, row in df.iterrows():
        if not any(clean_text(v) for v in row.tolist()):
            continue

        out.append(
            {
                "row_idx": i + 2,
                "No": clean_text(get(row, 0)),
                "Date": clean_text(get(row, 1)),
                "No. BOQ": clean_text(get(row, 2)),
                "SN Machine": clean_text(get(row, 3)),
                "Site Name": clean_text(get(row, 4)),
                "Charging Type": normalize_charger_type(get(row, 5)),
                "Item Name": clean_text(get(row, 6)),
                "Price": parse_price(get(row, 7)),
                "Periode": clean_text(get(row, 8)),
                "Qty": parse_qty_num(get(row, 9)),
                "Template": clean_text(get(row, 10)),
                "Region": clean_text(get(row, 11)),
                "Issue Note": clean_text(get(row, 12)),
                "Photo Evident 1": clean_text(get(row, 13)),
                "Photo Evident 2": clean_text(get(row, 14)),
            }
        )

    return out


def _next_boq_ms_sequence():
    max_no = 0
    for r in _read_boq_ms_records():
        m = re.match(r"^(\d+)", r["No"])
        if m:
            max_no = max(max_no, int(m.group(1)))

        m2 = re.match(
            r"^(\d+)/(?:CLX|SI|SUNRISE|SR)/BOQ/MS/",
            r["No. BOQ"].upper(),
        )
        if m2:
            max_no = max(max_no, int(m2.group(1)))

    return max_no + 1


def _material_rows_for_ms(df):
    if df is None or df.empty:
        return pd.DataFrame(
            columns=["Source Row", "NO", "Item Name", "Satuan", "Merk", "Unit Price"]
        )

    parents = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"}

    rows = []
    for i, r in df.iterrows():
        no = clean_text(r.get("NO", ""))
        item = clean_text(r.get("Item", ""))

        if not item or no.upper() in parents:
            continue

        rows.append(
            {
                "Source Row": i,
                "NO": no,
                "Item Name": item,
                "Satuan": clean_text(r.get("Satuan/Uom", "")),
                "Merk": clean_text(r.get("MERK", "")),
                "Unit Price": parse_price(r.get("UNIT PRICE", 0)),
            }
        )

    return pd.DataFrame(rows)


def _boq_ms_picker(material_df, selected_keys=None, key="boq_ms_picker"):
    work = material_df.copy()
    work.insert(0, "Pilih", False)
    work.insert(7, "Qty", 1.0)
    work["Total DPP"] = work["Unit Price"] * work["Qty"]

    if selected_keys is not None:
        selected_keys = set(selected_keys)
        work["Pilih"] = work["Source Row"].isin(selected_keys)

    view = work[
        [
            "Pilih",
            "NO",
            "Item Name",
            "Satuan",
            "Merk",
            "Unit Price",
            "Qty",
            "Total DPP",
        ]
    ].copy()
    view["Unit Price"] = view["Unit Price"].astype(float)
    view["Total DPP"] = view["Total DPP"].astype(float)

    edited = st.data_editor(
        view,
        key=key,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Pilih": st.column_config.CheckboxColumn("Pilih", help="Pilih material pengganti"),
            "NO": st.column_config.TextColumn("NO", disabled=True),
            "Item Name": st.column_config.TextColumn("Item Name", disabled=True),
            "Satuan": st.column_config.TextColumn("Satuan", disabled=True),
            "Merk": st.column_config.TextColumn("Merk", disabled=True),
            "Unit Price": st.column_config.NumberColumn("Unit Price", format="Rp %.0f", disabled=True),
            "Qty": st.column_config.NumberColumn("Qty", min_value=0, step=1, format="%.2f"),
            "Total DPP": st.column_config.NumberColumn("Total DPP", format="Rp %.0f", disabled=True),
        },
        disabled=[
            "NO", "Item Name", "Satuan", "Merk", "Unit Price", "Total DPP"
        ],
    )

    edited = edited.copy()
    edited["Qty"] = edited["Qty"].apply(parse_qty_num)
    edited["Total DPP"] = edited["Unit Price"] * edited["Qty"]

    selected = []
    for i, r in edited.iterrows():
        if bool(r["Pilih"]):
            src = material_df.iloc[i]
            selected.append(
                {
                    "Source Row": int(src["Source Row"]),
                    "NO": clean_text(src["NO"]),
                    "Item Name": clean_text(src["Item Name"]),
                    "Satuan": clean_text(src["Satuan"]),
                    "Merk": clean_text(src["Merk"]),
                    "Unit Price": parse_price(src["Unit Price"]),
                    "Qty": parse_qty_num(r["Qty"]),
                    "Total DPP": calculate_detail_total(r["Qty"], src["Unit Price"]),
                }
            )

    return pd.DataFrame(selected)


def _get_pdf_asset_path(filename):
    for base_dir in _candidate_asset_dirs():
        candidate = os.path.join(base_dir, filename)
        if os.path.isfile(candidate):
            return candidate
    return None


def _draw_pdf_asset(canvas, path, x, y_top, max_width, max_height):
    if not path or not os.path.exists(path):
        return

    try:
        from reportlab.lib.utils import ImageReader
        reader = ImageReader(path)
        iw, ih = reader.getSize()

        if not iw or not ih:
            return

        scale = min(float(max_width) / float(iw), float(max_height) / float(ih))
        width = iw * scale
        height = ih * scale
        x_draw = x + (max_width - width) / 2.0
        y_draw = y_top - height

        canvas.drawImage(
            reader,
            x_draw,
            y_draw,
            width=width,
            height=height,
            preserveAspectRatio=True,
            mask="auto",
        )
    except Exception:
        pass


def _boq_ms_pdf_header_footer(canvas, doc, company="CLX"):
    canvas.saveState()
    page_width, page_height = portrait(A5)
    left = 10
    usable_width = page_width - 20
    company_key = normalize_company_key(company)
    cfg = get_company_config(company_key)

    header_path = _get_pdf_asset_path("header.png") if company_key == "CLX" else None
    footer_path = _get_pdf_asset_path("Footer.png") if company_key == "CLX" else None

    if header_path:
        _draw_pdf_asset(canvas, header_path, left, page_height - 4, usable_width, 42)
    else:
        logo_path = get_logo_path(company_key)
        if logo_path:
            _draw_pdf_asset(canvas, logo_path, left, page_height - 7, 95, 34)

        canvas.setFont("Helvetica-Bold", 8)
        canvas.drawRightString(page_width - 10, page_height - 18, "BOQ MS / REPLACEMENT MATERIAL")
        canvas.setFont("Helvetica", 6)
        canvas.drawRightString(page_width - 10, page_height - 28, cfg["name"])
        canvas.setStrokeColor(colors.HexColor(cfg["border_primary_hex"]))
        canvas.setLineWidth(0.8 if company_key == "Sunrise" else 0.5)
        canvas.line(left, page_height - 40, page_width - 10, page_height - 40)

    if footer_path:
        _draw_pdf_asset(canvas, footer_path, left, 34, usable_width, 30)
    else:
        canvas.setStrokeColor(colors.HexColor(cfg["border_primary_hex"]))
        canvas.setLineWidth(0.5)
        canvas.line(left, 34, page_width - 10, 34)
        canvas.setFont("Helvetica", 6)
        canvas.drawString(left, 22, cfg["name"])
        canvas.drawRightString(page_width - 10, 22, f"Page {doc.page}")

    canvas.restoreState()


def generate_boq_ms_pdf(
    records,
    boq_no,
    sn_machine,
    site_name,
    charging_type,
    period,
    template_name,
    region,
    issue_note="",
    photo_bytes=None,
    date_obj=None,
    company=None,
):
    buf = io.BytesIO()
    resolved_company = company or detect_company_from_boq_no(boq_no)
    company_key = normalize_company_key(resolved_company)
    cfg = get_company_config(company_key)

    ms_header_bg = colors.HexColor(cfg["ms_header_hex"])
    ms_header_fg = colors.HexColor(cfg["ms_header_text_hex"])

    doc = SimpleDocTemplate(
        buf,
        pagesize=portrait(A5),
        rightMargin=10,
        leftMargin=10,
        topMargin=52,
        bottomMargin=42,
    )

    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "ms_title",
        parent=styles["Title"],
        fontSize=12,
        leading=14,
        alignment=1,
        spaceAfter=4,
    )

    small = ParagraphStyle("ms_small", parent=styles["Normal"], fontSize=6.5, leading=8)
    bold = ParagraphStyle("ms_bold", parent=small, fontName="Helvetica-Bold")
    cell = ParagraphStyle("ms_cell", parent=small, fontSize=6, leading=7)
    table_header = ParagraphStyle(
        "ms_table_header",
        parent=small,
        fontName="Helvetica-Bold",
        fontSize=6,
        leading=7,
        alignment=1,
        textColor=ms_header_fg,
    )

    story = []
    dt = date_obj or datetime.datetime.now()

    story.append(Paragraph("BOQ MS / REPLACEMENT MATERIAL", title))

    info = [
        [Paragraph("No. BOQ", bold), Paragraph(clean_text(boq_no), small)],
        [Paragraph("Date", bold), Paragraph(dt.strftime("%d-%m-%Y"), small)],
        [Paragraph("SN Machine", bold), Paragraph(clean_text(sn_machine), small)],
        [Paragraph("Site Name", bold), Paragraph(clean_text(site_name), small)],
        [Paragraph("Charging Type", bold), Paragraph(clean_text(charging_type), small)],
        [Paragraph("Periode", bold), Paragraph(clean_text(period), small)],
    ]

    ti = Table(info, colWidths=[65, 315])
    ti.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ]
        )
    )

    story += [ti, Spacer(1, 4)]

    data = [
        [
            Paragraph("No", table_header),
            Paragraph("Item Name", table_header),
            Paragraph("Qty", table_header),
            Paragraph("Unit Price", table_header),
            Paragraph("DPP", table_header),
        ]
    ]

    total = 0.0

    for i, r in enumerate(records, 1):
        qty = parse_qty_num(r.get("Qty", 0))
        up = parse_price(r.get("Unit Price", 0))
        dpp = calculate_detail_total(qty, up)
        total += dpp

        data.append(
            [
                Paragraph(str(i), cell),
                Paragraph(clean_text(r.get("Item Name", "")), cell),
                Paragraph(f"{qty:g}", cell),
                Paragraph(format_currency(up), cell),
                Paragraph(format_currency(dpp), cell),
            ]
        )

    table = Table(data, colWidths=[22, 171, 35, 76, 76], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), ms_header_bg),
                ("TEXTCOLOR", (0, 0), (-1, 0), ms_header_fg),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("ALIGN", (0, 1), (0, -1), "CENTER"),
                ("ALIGN", (2, 1), (2, -1), "CENTER"),
                ("ALIGN", (3, 1), (-1, -1), "RIGHT"),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )

    story += [table, Spacer(1, 4)]

    vat = total * 0.11
    grand = total + vat

    sm_styles = [
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]
    if company_key == "Sunrise":
        sm_styles.append(("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#FDE047")))

    sm = Table(
        [
            [Paragraph("DPP", bold), Paragraph(format_currency(total), bold)],
            [Paragraph("PPN 11%", small), Paragraph(format_currency(vat), small)],
            [Paragraph("Grand Total", bold), Paragraph(format_currency(grand), bold)],
        ],
        colWidths=[285, 95],
    )
    sm.setStyle(TableStyle(sm_styles))

    story.append(sm)
    story.append(Spacer(1, 5))

    issue_text = clean_text(issue_note).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    issue_text = issue_text.replace("\n", "<br/>") or "-"

    issue_table = Table(
        [[Paragraph("ISSUE NOTE", table_header)], [Paragraph(issue_text, cell)]],
        colWidths=[380],
    )
    issue_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), ms_header_bg),
                ("TEXTCOLOR", (0, 0), (-1, 0), ms_header_fg),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(issue_table)

    photos = [x for x in (photo_bytes or []) if x]
    if photos:
        story.append(Spacer(1, 5))
        photo_title = Table(
            [[Paragraph("FOTO EVIDENT", table_header)]],
            colWidths=[380],
        )
        photo_title.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                    ("BACKGROUND", (0, 0), (-1, 0), ms_header_bg),
                    ("TEXTCOLOR", (0, 0), (-1, 0), ms_header_fg),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        story.append(photo_title)

        photo_cells = []
        for b in photos[:2]:
            try:
                from reportlab.lib.utils import ImageReader
                reader = ImageReader(io.BytesIO(b))
                iw, ih = reader.getSize()
                max_w, max_h = 174, 115
                scale = min(max_w / float(iw), max_h / float(ih)) if iw and ih else 1
                w = iw * scale
                h = ih * scale
                photo_cells.append(Image(io.BytesIO(b), width=w, height=h))
            except Exception:
                photo_cells.append(Paragraph("Foto tidak dapat ditampilkan", cell))

        while len(photo_cells) < 2:
            photo_cells.append(Paragraph("", cell))

        photo_table = Table(
            [photo_cells[:2]],
            colWidths=[190, 190],
        )
        photo_table.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )
        story.append(photo_table)

    doc.build(
        story,
        onFirstPage=lambda c, d: _boq_ms_pdf_header_footer(c, d, company=company_key),
        onLaterPages=lambda c, d: _boq_ms_pdf_header_footer(c, d, company=company_key),
    )

    buf.seek(0)
    return buf.getvalue()


def _upload_boq_ms_photo(uploaded_file, boq_no):
    if uploaded_file is None:
        return ""
    try:
        sh = get_google_sheet_connection()
        creds = getattr(sh, "auth", None) if sh else None
        if creds is None:
            return ""
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload

        drive = build("drive", "v3", credentials=creds, cache_discovery=False)
        q = "name='BOQ MS Evidence' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res = drive.files().list(q=q, spaces="drive", fields="files(id,name)", pageSize=10).execute()
        files = res.get("files", [])
        if files:
            folder_id = files[0]["id"]
        else:
            folder = drive.files().create(
                body={"name": "BOQ MS Evidence", "mimeType": "application/vnd.google-apps.folder"},
                fields="id",
            ).execute()
            folder_id = folder["id"]
        safe = re.sub(r"[^A-Za-z0-9._-]+", "_", uploaded_file.name)
        filename = f"{boq_no.replace('/', '-')}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{safe}"
        media = MediaIoBaseUpload(
            io.BytesIO(uploaded_file.getvalue()),
            mimetype=uploaded_file.type or "image/jpeg",
            resumable=False,
        )
        created = drive.files().create(
            body={"name": filename, "parents": [folder_id]},
            media_body=media,
            fields="id,webViewLink",
        ).execute()
        file_id = created["id"]
        try:
            drive.permissions().create(fileId=file_id, body={"type": "anyone", "role": "reader"}).execute()
        except Exception:
            pass
        return created.get("webViewLink") or f"https://drive.google.com/file/d/{file_id}/view"
    except Exception as e:
        st.warning(f"⚠️ Foto berhasil dibaca, tetapi upload ke Google Drive gagal: {e}")
        return ""


def _download_boq_ms_photo(url):
    if not url:
        return None
    try:
        sh = get_google_sheet_connection()
        creds = getattr(sh, "auth", None) if sh else None
        if creds is None:
            return None
        m = re.search(r"/d/([A-Za-z0-9_-]+)", url) or re.search(r"[?&]id=([A-Za-z0-9_-]+)", url)
        if not m:
            return None
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload

        drive = build("drive", "v3", credentials=creds, cache_discovery=False)
        req = drive.files().get_media(fileId=m.group(1))
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        return fh.getvalue()
    except Exception:
        return None


def _save_boq_ms(
    records,
    sn_machine,
    site_name,
    charging_type,
    template_name,
    region,
    issue_note="",
    photo_files=None,
    date_obj=None,
    company="CLX",
):
    ws = _get_ws(DB_BOQ_MS_SHEET)
    if ws is None:
        raise RuntimeError(f"Worksheet '{DB_BOQ_MS_SHEET}' tidak ditemukan. Buat worksheet tersebut terlebih dahulu.")
    _ensure_boq_ms_headers(ws)
    dt = date_obj or datetime.datetime.now()
    seq = _next_boq_ms_sequence()
    boq_no = generate_boq_ms_number(seq, dt, company=company)
    period = _period_text(dt)
    photo_files = photo_files or []
    photo_urls = []
    for f in photo_files[:2]:
        u = _upload_boq_ms_photo(f, boq_no)
        photo_urls.append(u)
    while len(photo_urls) < 2:
        photo_urls.append("")
    rows = []
    for item in records:
        rows.append(
            [
                seq,
                dt.strftime("%d-%m-%Y"),
                boq_no,
                clean_text(sn_machine),
                clean_text(site_name),
                normalize_charger_type(charging_type),
                clean_text(item["Item Name"]),
                parse_price(item["Total DPP"]),
                period,
                parse_qty_num(item["Qty"]),
                template_db_value(template_name),
                clean_text(region),
                clean_text(issue_note),
                photo_urls[0],
                photo_urls[1],
            ]
        )
    if rows:
        ws.append_rows(rows, value_input_option="USER_ENTERED")
    return boq_no, period, seq, photo_urls


def _update_boq_ms(
    boq_no,
    records,
    sn_machine,
    site_name,
    charging_type,
    template_name,
    region,
    issue_note="",
    photo_files=None,
    date_text=None,
    existing_photo_urls=None,
):
    ws = _get_ws(DB_BOQ_MS_SHEET)
    if ws is None:
        raise RuntimeError(f"Worksheet '{DB_BOQ_MS_SHEET}' tidak ditemukan.")
    old = [r for r in _read_boq_ms_records() if r["No. BOQ"] == boq_no]
    if not old:
        raise RuntimeError("BOQ MS tidak ditemukan.")
    seq = parse_qty_num(old[0]["No"])
    dt = datetime.datetime.now()
    try:
        if date_text:
            dt = datetime.datetime.strptime(date_text, "%d-%m-%Y")
    except Exception:
        pass
    period = _period_text(dt)
    existing_photo_urls = existing_photo_urls or [
        old[0].get("Photo Evident 1", ""),
        old[0].get("Photo Evident 2", ""),
    ]
    uploaded_urls = []
    for f in (photo_files or [])[:2]:
        u = _upload_boq_ms_photo(f, boq_no)
        if u:
            uploaded_urls.append(u)
    photos = (uploaded_urls + existing_photo_urls)[:2]
    while len(photos) < 2:
        photos.append("")
    for r in old:
        ws.update(f"A{r['row_idx']}:O{r['row_idx']}", [[""] * 15])
    new_rows = []
    for item in records:
        new_rows.append(
            [
                int(seq),
                dt.strftime("%d-%m-%Y"),
                boq_no,
                clean_text(sn_machine),
                clean_text(site_name),
                normalize_charger_type(charging_type),
                clean_text(item["Item Name"]),
                parse_price(item["Total DPP"]),
                period,
                parse_qty_num(item["Qty"]),
                template_db_value(template_name),
                clean_text(region),
                clean_text(issue_note),
                photos[0],
                photos[1],
            ]
        )
    if not new_rows:
        return period
    ws.update(f"A{old[0]['row_idx']}:O{old[0]['row_idx']}", [new_rows[0]], value_input_option="USER_ENTERED")
    if len(new_rows) > 1:
        ws.insert_rows(new_rows[1:], row=old[0]["row_idx"] + 1, value_input_option="USER_ENTERED")
    return period


def _get_boq_ms_grouped():
    records = _read_boq_ms_records()
    groups = {}
    for r in records:
        groups.setdefault(r["No. BOQ"], []).append(r)
    return groups


def _records_to_pdf_items(records):
    return [
        {
            "Item Name": r["Item Name"],
            "Qty": r["Qty"],
            "Unit Price": r["Price"] / r["Qty"] if r["Qty"] else 0,
            "Total DPP": r["Price"],
        }
        for r in records
    ]


# ==============================================================================
# BOQ MS UI (CREATE & EDIT)
# ==============================================================================
def render_boq_ms_create():
    reset_no = int(st.session_state.get("boq_ms_create_reset_no", 0))
    key_suffix = f"_{reset_no}"

    st.subheader("➕ Create BOQ MS / Replacement Material")
    st.info("Pilih template, charging type, lalu pilih material replacement. Qty hanya berlaku untuk material yang dipilih.")

    col_comp, col_tmpl = st.columns([1, 2])
    with col_comp:
        selected_company = st.radio(
            "Output Perusahaan / Kop BOQ MS",
            COMPANY_OPTIONS,
            index=0,
            horizontal=True,
            key=f"ms_create_company{key_suffix}",
        )
    with col_tmpl:
        template_name = st.selectbox(
            "Template",
            list(TEMPLATE_OPTIONS.keys()),
            index=1,
            key=f"ms_create_template{key_suffix}",
        )

    charging_type = st.selectbox(
        "Charging Type",
        ["6S1P", "12S1P", "12S3P", "DC20", "DC30", "DC60"],
        key=f"ms_create_charger{key_suffix}",
    )
    region = st.selectbox(
        "Region / Island",
        ["JAVA", "SUMATERA", "BALI NUSATENGGARA", "KALIMANTAN", "SULAWESI"],
        key=f"ms_create_region{key_suffix}",
    )

    c1, c2 = st.columns(2)
    sn_machine = c1.text_input("SN Machine", key=f"ms_create_sn{key_suffix}")
    site_name = c2.text_input("Site Name", key=f"ms_create_site{key_suffix}")

    issue_note = st.text_area(
        "Issue Note",
        placeholder="Tuliskan issue / alasan replacement material...",
        key=f"ms_create_issue_note{key_suffix}",
        height=90,
    )
    photos = st.file_uploader(
        "📷 Foto Evident (Maks. 2 Foto)",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        key=f"ms_create_photos{key_suffix}",
        help="Upload maksimal 2 foto evidence.",
    )

    if len(photos) > 2:
        st.error("❌ Maksimal 2 foto evident. Silakan hapus foto yang berlebih.")
        return

    if not site_name:
        st.caption("Isi Site Name terlebih dahulu untuk melanjutkan.")
        return

    try:
        df_template, _ms_sheet, _ms_region = load_boq_dataframe(charging_type, region, template_name)
    except ValueError as e:
        st.error(f"❌ Regional template mismatch: {e}")
        return

    if df_template is None or df_template.empty:
        return

    materials = _material_rows_for_ms(df_template)
    if materials.empty:
        st.warning("Tidak ada material detail yang ditemukan pada template/sheet tersebut.")
        return

    st.markdown("#### Pilih Replacement Material")
    selected_df = _boq_ms_picker(materials, key=f"boq_ms_create_picker{key_suffix}")
    if selected_df.empty:
        st.warning("Belum ada material yang dipilih.")
        return

    total_dpp = float(selected_df["Total DPP"].sum())
    vat = total_dpp * 0.11
    grand = total_dpp + vat
    today = datetime.datetime.now()
    period = _period_text(today)

    c1, c2, c3 = st.columns(3)
    c1.metric("Item", len(selected_df))
    c2.metric("DPP", format_currency(total_dpp))
    c3.metric("Grand Total", format_currency(grand))
    st.write(f"**Periode:** {period}")

    preview_boq_no = generate_boq_ms_number(_next_boq_ms_sequence(), today, company=selected_company)
    pdf = generate_boq_ms_pdf(
        selected_df.to_dict("records"),
        preview_boq_no,
        sn_machine,
        site_name,
        charging_type,
        period,
        template_name,
        region,
        issue_note,
        [f.getvalue() for f in photos],
        today,
        company=selected_company,
    )

    st.download_button(
        "📄 Preview / Download PDF",
        pdf,
        file_name="BOQ_MS_PREVIEW.pdf",
        mime="application/pdf",
        key=f"ms_preview_pdf{key_suffix}",
    )

    if st.button("💾 Save BOQ MS", type="primary", key=f"ms_save{key_suffix}"):
        try:
            boq_no, period, seq, urls = _save_boq_ms(
                selected_df.to_dict("records"),
                sn_machine,
                site_name,
                charging_type,
                template_name,
                region,
                issue_note,
                photos,
                today,
                company=selected_company,
            )
            pdf = generate_boq_ms_pdf(
                selected_df.to_dict("records"),
                boq_no,
                sn_machine,
                site_name,
                charging_type,
                period,
                template_name,
                region,
                issue_note,
                [f.getvalue() for f in photos],
                today,
                company=selected_company,
            )
            st.cache_data.clear()
            st.session_state["boq_ms_last_pdf"] = pdf
            st.session_state["boq_ms_last_no"] = boq_no
            st.session_state["boq_ms_create_reset_no"] = reset_no + 1
            st.success(f"BOQ MS {boq_no} berhasil disimpan. Form Create BOQ MS dikosongkan kembali.")
            st.rerun()
        except Exception as e:
            st.error(f"Gagal menyimpan BOQ MS: {e}")


def render_boq_ms_edit():
    st.subheader("✏️ Edit / Re-Download BOQ MS")
    groups = _get_boq_ms_grouped()
    if not groups:
        st.info("Belum ada data pada DB BOQ MS.")
        return

    keys = list(groups.keys())
    labels = [f"{k} — {groups[k][0]['Site Name']} — {groups[k][0]['Charging Type']}" for k in keys]
    label = st.selectbox("Pilih BOQ MS", labels, key="ms_edit_select")
    boq_no = keys[labels.index(label)]
    records = groups[boq_no]
    first = records[0]

    default_company = detect_company_from_boq_no(boq_no)
    template_name = template_name_from_db_value(first.get("Template", "BOQ BARU"))
    region = first.get("Region") or "JAVA"
    charger = normalize_charger_type(first["Charging Type"])

    col_comp, col_tmpl = st.columns([1, 2])
    with col_comp:
        selected_company = st.radio(
            "Output Perusahaan / Kop BOQ MS",
            COMPANY_OPTIONS,
            index=COMPANY_OPTIONS.index(default_company),
            horizontal=True,
            key=f"ms_edit_company_{boq_no}",
        )
    with col_tmpl:
        template_name = st.selectbox(
            "Template",
            list(TEMPLATE_OPTIONS.keys()),
            index=list(TEMPLATE_OPTIONS.keys()).index(template_name),
            key="ms_edit_template",
        )

    sn = st.text_input("SN Machine", first.get("SN Machine", ""), key="ms_edit_sn")
    site = st.text_input("Site Name", first.get("Site Name", ""), key="ms_edit_site")

    charger_opts = ["6S1P", "12S1P", "12S3P", "DC20", "DC30", "DC60"]
    charger = st.selectbox(
        "Charging Type",
        charger_opts,
        index=charger_opts.index(charger) if charger in charger_opts else 0,
        key="ms_edit_charger",
    )

    ropts = ["JAVA", "SUMATERA", "BALI NUSATENGGARA", "KALIMANTAN", "SULAWESI"]
    region = st.selectbox(
        "Region / Island",
        ropts,
        index=ropts.index(region) if region in ropts else 0,
        key="ms_edit_region",
    )

    issue_note = st.text_area("Issue Note", first.get("Issue Note", ""), key="ms_edit_issue_note", height=90)
    existing_urls = [first.get("Photo Evident 1", ""), first.get("Photo Evident 2", "")]
    active_existing = [u for u in existing_urls if u]
    if active_existing:
        st.caption("📎 Existing Foto Evident: " + " | ".join(active_existing))

    new_photos = st.file_uploader(
        "📷 Ganti/Tambah Foto Evident (Maks. 2 Foto)",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        key="ms_edit_photos",
    )
    if len(new_photos) > 2:
        st.error("❌ Maksimal 2 foto evident.")
        return

    try:
        df_template, _ms_sheet, _ms_region = load_boq_dataframe(charger, region, template_name)
    except ValueError as e:
        st.error(f"❌ Regional template mismatch: {e}")
        return

    if df_template is None or df_template.empty:
        return

    materials = _material_rows_for_ms(df_template)
    existing_by_name = {}
    for r in records:
        existing_by_name.setdefault(r["Item Name"], []).append(r)

    selected_keys = [int(m["Source Row"]) for _, m in materials.iterrows() if m["Item Name"] in existing_by_name]
    selected_df = _boq_ms_picker(materials, selected_keys=selected_keys, key=f"boq_ms_edit_picker_{boq_no}")

    for i, r in selected_df.iterrows():
        matches = existing_by_name.get(r["Item Name"], [])
        if matches:
            q = matches[0].get("Qty", 0)
            selected_df.at[i, "Qty"] = q
            selected_df.at[i, "Total DPP"] = calculate_detail_total(q, r["Unit Price"])

    if selected_df.empty:
        st.warning("Belum ada material yang dipilih.")
        return

    total_dpp = float(selected_df["Total DPP"].sum())
    vat = total_dpp * 0.11
    grand = total_dpp + vat

    c1, c2, c3 = st.columns(3)
    c1.metric("Item", len(selected_df))
    c2.metric("DPP", format_currency(total_dpp))
    c3.metric("Grand Total", format_currency(grand))

    try:
        dt = datetime.datetime.strptime(first.get("Date", ""), "%d-%m-%Y")
    except Exception:
        dt = datetime.datetime.now()
    period = _period_text(dt)

    photo_bytes = [f.getvalue() for f in new_photos[:2]]
    if not photo_bytes:
        for u in active_existing[:2]:
            b = _download_boq_ms_photo(u)
            if b:
                photo_bytes.append(b)

    pdf = generate_boq_ms_pdf(
        selected_df.to_dict("records"),
        boq_no,
        sn,
        site,
        charger,
        period,
        template_name,
        region,
        issue_note,
        photo_bytes,
        dt,
        company=selected_company,
    )

    a, b = st.columns(2)
    with a:
        if st.button("💾 Update BOQ MS", type="primary", key=f"ms_update_{boq_no}"):
            try:
                _update_boq_ms(
                    boq_no,
                    selected_df.to_dict("records"),
                    sn,
                    site,
                    charger,
                    template_name,
                    region,
                    issue_note,
                    new_photos,
                    first.get("Date", ""),
                    existing_urls,
                )
                st.cache_data.clear()
                st.success(f"BOQ MS {boq_no} berhasil di-update.")
                st.rerun()
            except Exception as e:
                st.error(f"Gagal update BOQ MS: {e}")
    with b:
        st.download_button(
            "⬇️ Re-Download PDF",
            pdf,
            file_name=f"{boq_no.replace('/', '-')}.pdf",
            mime="application/pdf",
            key=f"ms_redownload_{boq_no}",
        )


def _render_boq_ms():
    st.header("🧰 BOQ MS — Replacement Material")
    tab1, tab2 = st.tabs(["➕ Create BOQ MS", "✏️ Edit / Re-Download BOQ MS"])
    with tab1:
        render_boq_ms_create()
    with tab2:
        render_boq_ms_edit()


# ==============================================================================
# MAIN RENDER & UI CONTROLLERS
# ==============================================================================
def _render_boq_create():
    st.subheader("➕ Buat BOQ Baru")

    if st.session_state.pop("boq_saved_success", False):
        saved_no = st.session_state.pop("last_saved_boq_no", "")
        st.success(f"BOQ berhasil disimpan{f' ({saved_no})' if saved_no else ''}.")

    site_options, site_data_map = fetch_query_site_options(exclude_saved=True)
    PLACEHOLDER_OPTION = "-- Pilih Site Name & Charger --"

    selected_label = st.selectbox(
        "Pilih Site Name & Charger",
        [PLACEHOLDER_OPTION] + sorted(site_options),
        index=0,
        key="create_site_selector",
    )

    if selected_label == PLACEHOLDER_OPTION:
        st.info("💡 Silakan pilih **Site Name & Charger** dari dropdown di atas.")
        return

    meta = site_data_map[selected_label]
    selected_site = meta.get("Site Name", selected_label)
    default_company = detect_company_from_context(epc_name=meta.get("EPC Name", ""))

    col_comp, col_tmpl = st.columns([1, 2])
    with col_comp:
        selected_company = st.radio(
            "Output Perusahaan / Kop BOQ",
            COMPANY_OPTIONS,
            index=COMPANY_OPTIONS.index(default_company),
            horizontal=True,
            key=f"create_boq_company_{selected_label}",
        )
    with col_tmpl:
        template_name = st.selectbox(
            "Template yang digunakan",
            list(TEMPLATE_OPTIONS.keys()),
            index=1,
            key="create_boq_template",
        )

    c1, c2 = st.columns(2)
    site_address = c2.text_input(
        "Address",
        value=meta.get("Address", "-"),
        key=f"address_{selected_label}",
    )

    c3, c4, c5 = st.columns([1.5, 1.5, 1])
    charging_type = c3.text_input(
        "Charging Type",
        value=meta.get("Charging Type", "DC20"),
        key=f"charger_{selected_label}",
    )

    charging_type = normalize_charger_type(charging_type)
    raw_province = clean_text(meta.get("Province", ""))
    province = map_to_standard_province(raw_province)

    c4.text_input("BOQ Region (AUTO)", value=province, disabled=True, key=f"province_mapped_{selected_label}")
    c5.caption(f"`{raw_province or '-'} → {province}`")

    epc_name = st.text_input(
        "EPC Name",
        value=meta.get("EPC Name", "-"),
        key=f"epc_{selected_label}",
    )

    saved_pairs = get_existing_saved_site_charger_pairs()
    pair = (selected_site.strip().lower(), normalize_charger_type(charging_type))
    is_already_saved = pair in saved_pairs and selected_site != "-"

    if is_already_saved and selected_site != "-":
        st.warning(f"⚠️ Kombinasi Site `{selected_site}` dengan Charger `{charging_type}` sudah pernah dibuatkan BOQ.")

    try:
        df_boq, target_sheet, region_normalized = load_boq_dataframe(
            charging_type,
            province,
            template_name,
        )
    except ValueError as e:
        st.error(f"❌ Regional template mismatch: {e}")
        return

    if df_boq is None or df_boq.empty:
        return

    st.subheader(f"2. Table BOQ ({target_sheet} - {region_normalized})")

    if template_name == "Template BOQ Vgreen Lama":
        df_boq = edit_boq_sections(df_boq, charging_type, f"create_old_template_{selected_label}")
    else:
        df_boq = edit_boq_sections(df_boq, charging_type, f"create_new_template_{selected_label}")

    df_boq, sub_total, vat_amount, grand_total = recalculate_boq_totals(df_boq)

    display = df_boq.copy()
    for c in ["UNIT PRICE", "TOTAL PRICE"]:
        display[c] = display[c].apply(lambda x: format_currency(x) if parse_price(x) != 0 else "-")

    st.dataframe(display, use_container_width=True, hide_index=True)

    preview_boq_no = generate_boq_number(
        sequence_num=_next_boq_sequence(),
        company=selected_company,
    )
    cfg_active = get_company_config(selected_company)

    c1, c2 = st.columns([2, 1])
    c1.success(
        f"✅ Tabel BOQ Aktif: **{selected_site}** | Tipe: **{charging_type}** | "
        f"Wilayah: **{region_normalized}** | Output: **{cfg_active['name']}** (`{preview_boq_no}`)"
    )
    c2.metric("Grand Total (Inc. VAT 11%)", format_currency(grand_total))

    if get_logo_path(selected_company) is None:
        st.warning(
            f"⚠️ File logo `{cfg_active['logo_candidates'][0]}` belum ditemukan di folder `assets/templates/`."
        )

    pdf_bytes = generate_boq_pdf(
        selected_site,
        site_address,
        charging_type,
        region_normalized,
        df_boq,
        sub_total,
        vat_amount,
        grand_total,
        boq_no=preview_boq_no,
        epc_name=epc_name,
        company=selected_company,
    )

    company_code = cfg_active["code"]
    safe_site = re.sub(r'[\\/*?:"<>|]', "_", str(selected_site))
    safe_charger = re.sub(r'[\\/*?:"<>|]', "_", str(charging_type))
    filename_pdf = f"BOQ_{company_code}_{safe_charger}_{region_normalized}_{safe_site}.pdf"

    b1, b2 = st.columns(2)
    with b1:
        if st.button("🚀 Simpan ke Database (DB BOQ)", type="primary", disabled=is_already_saved):
            boq_no = save_to_db_boq(
                selected_site,
                charging_type,
                sub_total,
                grand_total,
                epc_name,
                template_name,
                company=selected_company,
                df_boq=df_boq,
            )
            if boq_no:
                update_google_sheet_summary(selected_site, selected_site, sub_total, grand_total)
                st.session_state["boq_saved_success"] = True
                st.session_state["last_saved_boq_no"] = boq_no
                st.cache_data.clear()
                st.rerun()

    with b2:
        st.download_button("📥 Download PDF BOQ", pdf_bytes, file_name=filename_pdf, mime="application/pdf")


def _render_boq_edit():
    st.subheader("✏️ Edit, Reuse, & Re-Download BOQ Tersimpan")

    if st.session_state.pop("boq_updated_success", False):
        updated_no = st.session_state.pop("last_updated_boq_no", "")
        st.success(f"Perubahan BOQ berhasil disimpan ke database{f' ({updated_no})' if updated_no else ''}.")

    saved_boq_list = get_all_saved_boq()

    if not saved_boq_list:
        st.info("ℹ️ Belum ada data BOQ yang tersimpan di `DB BOQ`.")
        return

    options = {
        f"{x.get('BOQ No.', '-')}: {x.get('Site Name', '-')} [{x.get('Charger Type', '-')}]": x
        for x in saved_boq_list
    }

    key = st.selectbox("Pilih Nomor BOQ", sorted(options), key="edit_boq_selector")
    data = options[key]

    old_boq_no = data.get("BOQ No.", "-")
    old_site = data.get("Site Name", "")
    old_charger = data.get("Charger Type", "")
    old_epc = data.get("EPC Name", "-")
    saved_items_json = data.get("Items JSON", "")
    default_template_name = template_name_from_db_value(data.get("Template", ""))
    default_company = detect_company_from_context(boq_no=old_boq_no, epc_name=old_epc)

    col_comp, col_tmpl = st.columns([1, 2])
    with col_comp:
        selected_company = st.radio(
            "Output Perusahaan / Kop BOQ",
            COMPANY_OPTIONS,
            index=COMPANY_OPTIONS.index(default_company),
            horizontal=True,
            key=f"edit_boq_company_{key}",
        )
    with col_tmpl:
        template_name = st.selectbox(
            "Pilih Template",
            list(TEMPLATE_OPTIONS.keys()),
            index=list(TEMPLATE_OPTIONS.keys()).index(default_template_name),
            key=f"edit_boq_template_{key}",
        )

    active_boq_no = replace_boq_company_code(old_boq_no, selected_company)
    cfg_active = get_company_config(selected_company)
    st.caption(f"Output: **{cfg_active['name']}** | Nomor BOQ Aktif: `{active_boq_no}`")

    if get_logo_path(selected_company) is None:
        st.warning(
            f"⚠️ File logo `{cfg_active['logo_candidates'][0]}` belum ditemukan di folder `assets/templates/`."
        )

    _, qmap = fetch_query_site_options(exclude_saved=False)
    meta = next(
        (
            v for v in qmap.values()
            if v.get("Site Name", "").strip().lower() == old_site.strip().lower()
            and normalize_charger_type(v.get("Charging Type", "")) == normalize_charger_type(old_charger)
        ),
        {"Address": "-", "Province": "JAVA", "EPC Name": "-"},
    )

    site = st.text_input("Site Name", old_site, key=f"edit_site_{key}")
    charger = normalize_charger_type(st.text_input("Charger Type", old_charger, key=f"edit_charger_{key}"))
    epc = st.text_input("EPC Name", old_epc, key=f"edit_epc_{key}")
    address = st.text_input("Address", meta.get("Address", "-"), key=f"edit_address_{key}")
    province = map_to_standard_province(meta.get("Province", ""))

    try:
        df, target, region = load_boq_dataframe(charger, province, template_name)
    except ValueError as e:
        st.error(f"❌ Regional template mismatch: {e}")
        return

    if df is None or df.empty:
        return

    # Pulihkan rincian Qty hasil edit sebelumnya jika tersimpan di DB BOQ (Items JSON)
    has_saved_snapshot = bool(clean_text(saved_items_json) and clean_text(saved_items_json) != "[]")
    if has_saved_snapshot:
        df = apply_saved_boq_items_json(df, saved_items_json)
        st.info("📂 Data Qty & rincian BOQ sebelumnya berhasil dimuat dari database.")
    else:
        st.caption("ℹ️ BOQ ini disimpan sebelum fitur penyimpanan detail aktif, sehingga memuat nilai awal template.")

    df = edit_boq_sections(df, charger, f"edit_detail_{key}_{template_name}_{charger}")
    df, sub, vat, grand = recalculate_boq_totals(df)

    display = df.copy()
    for c in ["UNIT PRICE", "TOTAL PRICE"]:
        display[c] = display[c].apply(lambda x: format_currency(x) if parse_price(x) != 0 else "-")
    st.dataframe(display, use_container_width=True, hide_index=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Sub Total (Exc. PPN)", format_currency(sub))
    c2.metric("PPN 11%", format_currency(vat))
    c3.metric("Grand Total (Inc. PPN)", format_currency(grand))

    pdf = generate_boq_pdf(
        site,
        address,
        charger,
        region,
        df,
        sub,
        vat,
        grand,
        boq_no=active_boq_no,
        epc_name=epc,
        company=selected_company,
    )

    company_code = cfg_active["code"]
    safe_site = re.sub(r'[\\/*?:"<>|]', "_", str(site))
    safe_charger = re.sub(r'[\\/*?:"<>|]', "_", str(charger))
    filename_updated = f"BOQ_{company_code}_{safe_charger}_{region}_{safe_site}_Updated.pdf"

    a, b = st.columns(2)
    with a:
        if st.button("💾 Save & Update BOQ Database", type="primary", key=f"btn_update_boq_{key}"):
            if update_db_boq_row(
                data["row_idx"],
                old_site,
                site,
                charger,
                sub,
                grand,
                epc,
                template_name,
                new_boq_no=active_boq_no,
                df_boq=df,
            ):
                st.session_state["boq_updated_success"] = True
                st.session_state["last_updated_boq_no"] = active_boq_no
                st.cache_data.clear()
                st.rerun()

    with b:
        st.download_button("📥 Re-Download PDF BOQ", pdf, file_name=filename_updated, mime="application/pdf", key=f"download_edit_{key}")


def render():
    st.title("📝 Quotation & BOQ Manager")
    tab_create, tab_edit, tab_ms = st.tabs([
        "➕ Buat BOQ Baru",
        "✏️ Edit / Reuse BOQ",
        "🧰 Create BOQ MS",
    ])

    with tab_create:
        _render_boq_create()
    with tab_edit:
        _render_boq_edit()
    with tab_ms:
        _render_boq_ms()


if __name__ == "__main__":
    render()
