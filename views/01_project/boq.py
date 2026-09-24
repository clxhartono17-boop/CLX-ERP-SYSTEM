import datetime
import io
import os
import re
import sys

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, A5, portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, Image
import streamlit as st


# ==============================================================================
# SAFE IMPORT / ROOT CONFIGURATION
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
# TEMPLATE CONSTANTS
# ==============================================================================
OLD_TEMPLATE_FILENAME = "Template BOQ Vgreen.xlsx"
NEW_TEMPLATE_FILENAME = "New Template BOQ Sept 2026 All Charger.xlsx"

# Increment whenever template parsing/edit-state behavior changes.
# This also invalidates Streamlit cache after deployment.
BOQ_TEMPLATE_PARSER_VERSION = "2026-09-24-V3"
TEMPLATE_OPTIONS = {
    "Template BOQ Vgreen Lama": OLD_TEMPLATE_FILENAME,
    "New Template BOQ Sept 2026": NEW_TEMPLATE_FILENAME,
}
TEMPLATE_DB_VALUES = {
    "Template BOQ Vgreen Lama": "BOQ LAMA",
    "New Template BOQ Sept 2026": "BOQ BARU",
}
DB_TEMPLATE_VALUES = ["BOQ LAMA", "BOQ BARU"]

DB_BOQ_SHEET = "DB BOQ"
DB_BOQ_MS_SHEET = "DB BOQ MS"
QUERY_SHEET = "Query"
SUM_PROJECT_SHEET = "Sum Project"

BOQ_MS_HEADERS = [
    "No", "Date", "No. BOQ", "SN Machine", "Site Name",
    "Charging Type", "Item Name", "Price", "Periode"
]
BOQ_MS_HELPER_HEADERS = ["Qty", "Template", "Region", "Issue Note", "Photo Evident 1", "Photo Evident 2"]


# ==============================================================================
# GENERIC HELPERS
# ==============================================================================
def parse_price(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return 0.0
    s = s.replace("Rp", "").replace("rp", "").replace(" ", "")
    if "," in s and "." in s:
        # Indonesian convention: 1.234.567,89
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
        # A single dot followed by 1-2 digits is usually decimal; otherwise
        # treat dots as thousands separators.
        if s.count(".") > 1 or (s.count(".") == 1 and len(s.rsplit(".", 1)[-1]) == 3):
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
    if not s:
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
        return f"Rp. {float(value):,.0f}".replace(",", ".")
    except Exception:
        return "Rp. 0"


def clean_text(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def map_to_standard_province(raw_province):
    s = clean_text(raw_province).upper()
    if not s:
        return "JAVA"
    if any(x in s for x in ["JAKARTA", "BANTEN", "JAWA", "YOGYAKARTA", "DI YOGYAKARTA"]):
        return "JAVA"
    if any(x in s for x in ["SUMATERA", "ACEH", "RIAU", "JAMBI", "BENGKULU", "LAMPUNG", "BANGKA", "BELITUNG"]):
        return "SUMATERA"
    if any(x in s for x in ["BALI", "NUSA TENGGARA", "NTB", "NTT"]):
        return "BALI NUSATENGGARA"
    if "KALIMANTAN" in s:
        return "KALIMANTAN"
    if any(x in s for x in ["SULAWESI", "GORONTALO"]):
        return "SULAWESI"
    if any(x in s for x in ["PAPUA", "MALUKU"]):
        return "SULAWESI"
    return s


def normalize_charger_type(charging_type):
    s = clean_text(charging_type).upper().replace(" ", "")
    aliases = {
        "20KW": "DC20", "DC20KW": "DC20", "DC20": "DC20",
        "30KW": "DC30", "DC30KW": "DC30", "DC30": "DC30",
        "60KW": "DC60", "DC60KW": "DC60", "DC60": "DC60",
        "6S1P": "6S1P", "BSS6S1P": "6S1P",
        "12S1P": "12S1P", "BSS12S1P": "12S1P",
        "12S3P": "12S3P", "BSS12S3P": "12S3P",
        "7KW": "7KW", "AC7": "7KW", "22KW": "22KW", "AC22": "22KW",
        "DC120": "DC120", "120KW": "DC120",
    }
    return aliases.get(s, clean_text(charging_type).strip())


def generate_boq_number(sequence_num=1):
    now = datetime.datetime.now()
    roman = {1:"I",2:"II",3:"III",4:"IV",5:"V",6:"VI",7:"VII",8:"VIII",9:"IX",10:"X",11:"XI",12:"XII"}[now.month]
    return f"{int(sequence_num):04d}/CLX/BOQ/{roman}/{now.year}"


def generate_boq_ms_number(sequence_num=1, date_obj=None):
    now = date_obj or datetime.datetime.now()
    roman = {1:"I",2:"II",3:"III",4:"IV",5:"V",6:"VI",7:"VII",8:"VIII",9:"IX",10:"X",11:"XI",12:"XII"}[now.month]
    return f"{int(sequence_num):04d}/CLX/BOQ/MS/{roman}/{now.year}"


def get_template_path(filename):
    candidates = [
        os.path.join(CURRENT_DIR, "assets", "templates", filename),
        os.path.join(ROOT_DIR, "assets", "templates", filename),
        os.path.join(CWD, "assets", "templates", filename),
        os.path.join(CWD, "assets", "templates", "boq", filename),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return candidates[0]


def get_logo_path():
    candidates = [
        os.path.join(CURRENT_DIR, "assets", "templates", "CLX.png"),
        os.path.join(ROOT_DIR, "assets", "templates", "CLX.png"),
        os.path.join(CWD, "assets", "templates", "CLX.png"),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def template_db_value(template_name):
    return TEMPLATE_DB_VALUES.get(template_name, template_name)


def template_name_from_db_value(value):
    v = clean_text(value).upper()
    return "Template BOQ Vgreen Lama" if v == "BOQ LAMA" else "New Template BOQ Sept 2026"


def get_new_template_sheet_name(charging_type, province_str):
    c = normalize_charger_type(charging_type)
    region = map_to_standard_province(province_str)
    if c in {"6S1P", "12S1P", "12S3P"}:
        return c
    if c in {"DC20", "DC30", "DC60"}:
        return c if region == "JAVA" else f"{c} (NON JAVA)"
    return c


def detect_template_header(raw, keywords=("NO", "ITEM")):
    for i in range(min(30, len(raw))):
        row = [clean_text(x).upper() for x in raw.iloc[i].tolist()]
        if all(any(k in cell for cell in row) for k in keywords):
            return i
    return 0


def find_column_by_keywords(columns, keywords):
    for c in columns:
        s = clean_text(c).upper()
        if any(k in s for k in keywords):
            return c
    return None



def build_new_template_column_mapping(columns):
    """
    Map the real Excel template columns without guessing from column position.

    The September-2026 workbook uses headers such as:
        NO | Item | Spec | Qty | Lot | MERK | UNIT PRICE 20kw | TOTAL PRICE 20kw

    The old workbook uses:
        NO | Item | VOLUME | [Lot] | MERK | UNIT PRICE | TOTAL PRICE

    Exact header matches are always preferred.  Keyword matching is only a
    fallback and every source column can be used only once.
    """
    cols = list(columns)

    def norm(value):
        return re.sub(r"[^A-Z0-9]+", "", clean_text(value).upper())

    normalized = [(col, norm(col)) for col in cols]
    used = set()
    mapping = {}

    def pick_exact(aliases):
        aliases_norm = {norm(x) for x in aliases}
        for col, ncol in normalized:
            if col in used:
                continue
            if ncol in aliases_norm:
                used.add(col)
                return col
        return None

    def pick_keyword(keywords):
        keywords_norm = [norm(x) for x in keywords]
        for col, ncol in normalized:
            if col in used:
                continue
            if any(k and k in ncol for k in keywords_norm):
                used.add(col)
                return col
        return None

    mapping["NO"] = pick_exact(["NO", "NO."])
    mapping["ITEM"] = pick_exact([
        "ITEM", "ITEM NAME", "DESCRIPTION", "DESCRIPTION / ITEM",
        "MATERIAL", "MATERIAL NAME",
    ])

    # Qty / Volume is deliberately matched before Satuan/Lot.
    mapping["UNIT_VOLUME"] = pick_exact([
        "UNIT/VOLUME", "UNIT / VOLUME", "UNIT/VOL", "UNIT / VOL",
        "VOLUME", "QTY", "QTY.", "QUANTITY",
    ])

    # New template calls this column "Lot"; old template can have a blank
    # header in this position, so Lot is an explicit alias.
    mapping["SATUAN"] = pick_exact([
        "SATUAN", "SATUAN/UOM", "SATUAN / UOM", "UOM",
        "LOT", "UNIT",
    ])

    mapping["MERK"] = pick_exact(["MERK", "MEREK", "BRAND"])

    mapping["UNIT_PRICE"] = pick_exact([
        "UNIT PRICE", "UNIT PRICE (IDR)", "UNIT PRICE (RP)", "PRICE",
    ])
    mapping["TOTAL_PRICE"] = pick_exact([
        "TOTAL PRICE", "TOTAL PRICE (IDR)", "TOTAL PRICE (RP)", "TOTAL",
    ])

    # The EVCS template has "UNIT PRICE 20kw", "UNIT PRICE 30kw", etc.
    # Therefore exact matching above is followed by a safe keyword fallback.
    if mapping["NO"] is None:
        mapping["NO"] = pick_keyword(["NO"])
    if mapping["ITEM"] is None:
        mapping["ITEM"] = pick_keyword(["ITEM", "DESCRIPTION", "MATERIAL"])
    if mapping["UNIT_VOLUME"] is None:
        mapping["UNIT_VOLUME"] = pick_keyword([
            "UNIT/VOL", "UNITVOL", "VOLUME", "QTY", "QUANTITY"
        ])
    if mapping["SATUAN"] is None:
        # Never use generic "UNIT" here; it can capture Unit/Volume.
        mapping["SATUAN"] = pick_keyword(["SATUAN", "UOM", "LOT"])
    if mapping["MERK"] is None:
        mapping["MERK"] = pick_keyword(["MERK", "MEREK", "BRAND"])
    if mapping["UNIT_PRICE"] is None:
        mapping["UNIT_PRICE"] = pick_keyword(["UNITPRICE", "PRICE"])
    if mapping["TOTAL_PRICE"] is None:
        mapping["TOTAL_PRICE"] = pick_keyword(["TOTALPRICE", "TOTAL"])

    return mapping

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
        rr = list(r) + [""] * max(0, width - len(r))
        rows.append(rr[:width])
    return pd.DataFrame(rows, columns=header)


@st.cache_data(ttl=120, show_spinner=False)
def read_sheet_df(sheet_name):
    ws = _get_ws(sheet_name)
    if ws is None:
        return pd.DataFrame()
    try:
        return _values_to_df(ws.get_all_values())
    except Exception:
        return pd.DataFrame()


# ==============================================================================
# DB BOQ (EXISTING)
# ==============================================================================
def ensure_db_boq_headers(worksheet):
    headers = ["No", "BOQ No.", "Site Name", "Charger Type", "BOQ Amount Exc. PPN", "BOQ Amount inc. PPN", "EPC Name", "Template"]
    try:
        vals = worksheet.get_all_values()
        first = vals[0] if vals else []
        if first[:len(headers)] != headers:
            worksheet.update("A1:H1", [headers])
    except Exception:
        pass
    return headers


@st.cache_data(ttl=120, show_spinner=False)
def get_all_saved_boq():
    df = read_sheet_df(DB_BOQ_SHEET)
    if df.empty:
        return []
    cols = list(df.columns)
    def col(letter, fallback):
        idx = ord(letter) - 65
        return cols[idx] if idx < len(cols) else fallback
    no_c, boq_c, site_c, charger_c, dpp_c, inc_c, epc_c, temp_c = [col(chr(65+i), "") for i in range(8)]
    out = []
    for idx, row in df.iterrows():
        if not any(clean_text(v) for v in row.tolist()):
            continue
        out.append({
            "row_idx": idx + 2,
            "No": clean_text(row.get(no_c, "")),
            "BOQ No.": clean_text(row.get(boq_c, "")),
            "Site Name": clean_text(row.get(site_c, "")),
            "Charger Type": clean_text(row.get(charger_c, "")),
            "BOQ Amount Exc. PPN": parse_price(row.get(dpp_c, 0)),
            "BOQ Amount inc. PPN": parse_price(row.get(inc_c, 0)),
            "EPC Name": clean_text(row.get(epc_c, "")),
            "Template": clean_text(row.get(temp_c, "")),
        })
    return out


@st.cache_data(ttl=120, show_spinner=False)
def get_existing_saved_site_charger_pairs():
    return {(x["Site Name"].strip().lower(), normalize_charger_type(x["Charger Type"])) for x in get_all_saved_boq() if x["Site Name"]}


@st.cache_data(ttl=120, show_spinner=False)
def fetch_query_site_options(exclude_saved=True):
    df = read_sheet_df(QUERY_SHEET)
    if df.empty:
        return [], {}
    # Original Query structure: B EPC, C Charger, D Status, F Site, G Address, I Province.
    def at(pos):
        return df.columns[pos] if pos < len(df.columns) else None
    epc_c, charger_c, status_c, site_c, addr_c, province_c = at(1), at(2), at(3), at(5), at(6), at(8)
    if site_c is None:
        return [], {}
    saved = get_existing_saved_site_charger_pairs() if exclude_saved else set()
    options, data = [], {}
    for _, r in df.iterrows():
        site = clean_text(r.get(site_c, ""))
        charger = normalize_charger_type(r.get(charger_c, ""))
        status = clean_text(r.get(status_c, "")).upper() if status_c else ""
        if not site or status in {"DROP", "CANCEL"}:
            continue
        key = (site.lower(), charger)
        if exclude_saved and key in saved:
            continue
        display = f"{site} ({charger})"
        if display not in data:
            data[display] = {
                "Site Name": site,
                "Charging Type": charger,
                "Address": clean_text(r.get(addr_c, "")) if addr_c else "",
                "Province": clean_text(r.get(province_c, "")) if province_c else "",
                "EPC Name": clean_text(r.get(epc_c, "")) if epc_c else "",
            }
            options.append(display)
    return options, data


def save_to_db_boq(
    site_name, charger_capacity, sub_total, grand_total, epc_name="-", template_name="Template BOQ Vgreen Lama"
):
    try:
        sh = get_google_sheet_connection()
        if not sh:
            return None
        try:
            worksheet = sh.worksheet("DB BOQ")
        except Exception:
            worksheet = sh.add_worksheet(title="DB BOQ", rows=1000, cols=10)
        ensure_db_boq_headers(worksheet)
        existing_rows = worksheet.get_all_values()
        no_urut = len([r for r in existing_rows[1:] if any(r)]) + 1 if len(existing_rows) > 1 else 1
        boq_no = generate_boq_number(sequence_num=no_urut)
        new_row = [no_urut, boq_no, str(site_name), str(charger_capacity), sub_total, grand_total, str(epc_name), template_db_value(template_name)]
        worksheet.append_row(new_row)
        return boq_no
    except Exception as e:
        st.error(f"❌ Gagal menyimpan ke DB BOQ: {e}")
        return None

def update_db_boq_row(
    row_idx, old_site_name, new_site_name, charger_capacity, sub_total, grand_total, epc_name, template_name=None
):
    try:
        sh = get_google_sheet_connection()
        if not sh:
            return False
        worksheet = sh.worksheet("DB BOQ")
        ensure_db_boq_headers(worksheet)
        worksheet.update_cell(row_idx,3,new_site_name)
        worksheet.update_cell(row_idx,4,charger_capacity)
        worksheet.update_cell(row_idx,5,sub_total)
        worksheet.update_cell(row_idx,6,grand_total)
        worksheet.update_cell(row_idx,7,epc_name)
        if template_name:
            worksheet.update_cell(row_idx,8,template_db_value(template_name))
        update_google_sheet_summary(old_site_name,new_site_name,sub_total,grand_total)
        return True
    except Exception as e:
        st.error(f"❌ Gagal memperbarui DB BOQ: {e}")
        return False

def update_google_sheet_summary(old_site_name,new_site_name,sub_total,grand_total):
    try:
        sh=get_google_sheet_connection()
        if not sh: return False
        try:
            worksheet=sh.worksheet("Sum Project")
        except Exception:
            return False
        data_sum=worksheet.get_all_values()
        for idx,row in enumerate(data_sum[1:],start=2):
            if len(row)>2 and str(row[2]).strip().lower()==str(old_site_name).strip().lower():
                worksheet.update_cell(idx,3,new_site_name)
                if len(row)>=19:
                    worksheet.update_cell(idx,18,sub_total)
                    worksheet.update_cell(idx,19,grand_total)
                return True
    except Exception as e:
        st.caption(f"ℹ️ Info: Sheet 'Sum Project' belum ter-update ({e})")
    return False

def calculate_detail_total(quantity, unit_price):
    return parse_qty_num(quantity) * parse_price(unit_price)





def recalculate_boq_totals(df_boq):
    """
    Recalculate BOQ totals while preserving the template's displayed values.

    The template is the source of truth for every existing detail TOTAL PRICE.
    User edits are already applied by edit_boq_sections() / the old-template
    editor, which updates the affected detail TOTAL PRICE. This function then
    rolls those values up into section totals and the DPP.

    This avoids changing a template value such as 37,000 into 36,972 merely
    because Excel displayed a rounded/cached value.
    """
    if df_boq is None or df_boq.empty:
        return df_boq, 0.0, 0.0, 0.0

    df = df_boq.copy()

    required = [
        "NO", "Item", "Unit/Volume", "Satuan/Uom",
        "MERK", "UNIT PRICE", "TOTAL PRICE"
    ]
    for c in required:
        if c not in df.columns:
            df[c] = ""

    parent_labels = {"A", "B", "C", "D", "E", "F", "G", "H"}

    def item_upper(row):
        return clean_text(row.get("Item", "")).upper()

    def is_dpp_summary(row):
        item = item_upper(row)
        return item in {"TOTAL", "SUB TOTAL", "SUBTOTAL"}

    def is_vat_summary(row):
        item = item_upper(row)
        return bool(re.match(r"^(VAT|PPN)\b", item))

    def is_grand_summary(row):
        item = item_upper(row)
        return (
            "TOTAL CONTRACTOR PRICE" in item
            or item in {"GRAND TOTAL", "TOTAL INC. VAT", "TOTAL INCLUDING VAT"}
        )

    def is_roman_subtotal(row):
        no = clean_text(row.get("NO", "")).upper()
        item = clean_text(row.get("Item", ""))
        if not no or not item:
            return False
        if not re.fullmatch(r"[IVXLCDM]+", no):
            return False

        # Roman rows such as I / II in EVCS are cached subsection totals.
        return (
            parse_price(row.get("UNIT PRICE", 0)) == 0
            and parse_price(row.get("TOTAL PRICE", 0)) != 0
        )

    def template_row_total(row):
        """
        Preserve the actual displayed/cached TOTAL PRICE from Excel.

        If the template has no total but the row is directly calculable,
        calculate it as a safe fallback.
        """
        cached_total = parse_price(row.get("TOTAL PRICE", 0))
        if cached_total != 0:
            return cached_total

        qty = parse_qty_num(row.get("Unit/Volume", 0))
        unit_price = parse_price(row.get("UNIT PRICE", 0))
        if qty != 0 and unit_price != 0:
            return qty * unit_price

        return 0.0

    parent_positions = [
        pos for pos, (_, row) in enumerate(df.iterrows())
        if clean_text(row.get("NO", "")).upper() in parent_labels
    ]

    totals = [0.0] * len(df)

    # --------------------------------------------------------------------------
    # Main section roll-up
    # --------------------------------------------------------------------------
    for parent_index, parent_pos in enumerate(parent_positions):
        next_parent_pos = (
            parent_positions[parent_index + 1]
            if parent_index + 1 < len(parent_positions)
            else len(df)
        )

        parent_row = df.iloc[parent_pos]
        parent_qty = parse_qty_num(parent_row.get("Unit/Volume", 0))
        parent_price = parse_price(parent_row.get("UNIT PRICE", 0))
        parent_cached_total = parse_price(parent_row.get("TOTAL PRICE", 0))

        # A directly priced section (for example transport) remains exactly
        # the template value / Qty x Unit Price.
        if parent_qty != 0 and parent_price != 0:
            totals[parent_pos] = (
                parent_qty * parent_price
                if parent_cached_total == 0
                else parent_cached_total
            )
            continue

        section_total = 0.0
        roman_active = False

        for pos in range(parent_pos + 1, next_parent_pos):
            row = df.iloc[pos]

            if is_dpp_summary(row) or is_vat_summary(row) or is_grand_summary(row):
                continue

            no = clean_text(row.get("NO", "")).upper()

            # Roman subsection: count its subtotal once.
            if is_roman_subtotal(row):
                section_total += parse_price(row.get("TOTAL PRICE", 0))
                roman_active = True
                continue

            if roman_active:
                # All ordinary rows following a Roman subtotal belong to that
                # subsection and are already included in its subtotal.
                if re.fullmatch(r"[IVXLCDM]+", no):
                    roman_active = True
                else:
                    continue

            section_total += template_row_total(row)

        # If the section had no child detail but has a valid cached total,
        # preserve that value.
        if section_total == 0 and parent_cached_total != 0:
            section_total = parent_cached_total

        totals[parent_pos] = section_total

    # --------------------------------------------------------------------------
    # No A-H section fallback
    # --------------------------------------------------------------------------
    if parent_positions:
        subtotal = sum(totals[p] for p in parent_positions)
    else:
        subtotal = 0.0
        for _, row in df.iterrows():
            if is_dpp_summary(row) or is_vat_summary(row) or is_grand_summary(row):
                continue
            subtotal += template_row_total(row)

    vat = subtotal * 0.11
    grand = subtotal + vat

    # --------------------------------------------------------------------------
    # Write rolled-up values
    # --------------------------------------------------------------------------
    for pos, (_, row) in enumerate(df.iterrows()):
        if pos in parent_positions:
            df.iloc[pos, df.columns.get_loc("TOTAL PRICE")] = totals[pos]
        elif is_dpp_summary(row):
            df.iloc[pos, df.columns.get_loc("TOTAL PRICE")] = subtotal
        elif is_vat_summary(row):
            df.iloc[pos, df.columns.get_loc("TOTAL PRICE")] = vat
        elif is_grand_summary(row):
            df.iloc[pos, df.columns.get_loc("TOTAL PRICE")] = grand
        elif is_roman_subtotal(row):
            # Preserve template subsection subtotal.
            df.iloc[pos, df.columns.get_loc("TOTAL PRICE")] = parse_price(
                row.get("TOTAL PRICE", 0)
            )
        else:
            # Keep the template detail total; user-edited rows already carry
            # their newly calculated TOTAL PRICE.
            df.iloc[pos, df.columns.get_loc("TOTAL PRICE")] = template_row_total(row)

    return df, subtotal, vat, grand

def _old_template_offsets(region):
    return {"JAVA":0, "SUMATERA":8, "BALI NUSATENGGARA":16, "KALIMANTAN":25, "SULAWESI":33}.get(region, 0)


def _first_series_by_header(data, source):
    """Return exactly one Series even when Excel contains duplicate headers."""
    if source is None or data is None or data.empty:
        return None

    # IMPORTANT: data[source] returns a DataFrame when Excel has duplicate
    # headers.  Select by position so assignment to one output column is safe.
    try:
        positions = [i for i, col in enumerate(data.columns) if col == source]
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



def _standardize_template_df(raw, header_row=None):
    if raw is None or raw.empty:
        return pd.DataFrame(
            columns=["NO", "Item", "Unit/Volume", "Satuan/Uom",
                     "MERK", "UNIT PRICE", "TOTAL PRICE"]
        )

    raw = raw.copy()
    if header_row is None:
        header_row = detect_template_header(raw)

    if header_row < 0 or header_row >= len(raw):
        header_row = 0

    headers = [clean_text(x) for x in raw.iloc[header_row].tolist()]
    data = raw.iloc[header_row + 1:].copy()
    data.columns = headers

    # Remove only columns whose header is genuinely blank.
    # Do not remove blank data rows: they are part of the visual template.
    keep_positions = [i for i, c in enumerate(data.columns) if clean_text(c) != ""]
    if not keep_positions:
        return pd.DataFrame(
            columns=["NO", "Item", "Unit/Volume", "Satuan/Uom",
                     "MERK", "UNIT PRICE", "TOTAL PRICE"]
        )

    data = data.iloc[:, keep_positions]

    mapping = build_new_template_column_mapping(data.columns)
    out = pd.DataFrame(index=data.index)

    for target, source in mapping.items():
        if source is None:
            continue

        # Select by physical column position. This is important when an Excel
        # template contains duplicate/merged headers.
        positions = [i for i, col in enumerate(data.columns) if col == source]
        if not positions:
            continue

        series = data.iloc[:, positions[0]]
        out[target] = series.to_numpy()

    out = out.rename(columns={
        "NO": "NO",
        "ITEM": "Item",
        "UNIT_VOLUME": "Unit/Volume",
        "SATUAN": "Satuan/Uom",
        "MERK": "MERK",
        "UNIT_PRICE": "UNIT PRICE",
        "TOTAL_PRICE": "TOTAL PRICE",
    })

    required = [
        "NO", "Item", "Unit/Volume", "Satuan/Uom",
        "MERK", "UNIT PRICE", "TOTAL PRICE"
    ]
    for c in required:
        if c not in out.columns:
            out[c] = ""

    # Defensive fallback for September-2026 templates: the UOM column is
    # physically the Excel "Lot" column. If a header variation prevented the
    # normal mapping, recover it by position relative to Qty/MERK.
    if out["Satuan/Uom"].apply(clean_text).eq("").all():
        try:
            qty_src = mapping.get("UNIT_VOLUME")
            lot_src = mapping.get("SATUAN")
            if lot_src is not None:
                lot_positions = [i for i, col in enumerate(data.columns) if col == lot_src]
                if lot_positions:
                    out["Satuan/Uom"] = data.iloc[:, lot_positions[0]].to_numpy()
        except Exception:
            pass


    # Normalize summary rows that use a different physical layout in the
    # legacy workbook.  Example:
    #   F54 = "Sub Total :" and G54 = 47,100,000
    #   A56 = 41,500 and F56 = "Total Contractor Price"
    #
    # The visible BOQ meaning is preserved in the canonical Item column so
    # the table/PDF and calculation engine can recognize it correctly.
    for i in out.index:
        no = clean_text(out.at[i, "NO"])
        item = clean_text(out.at[i, "Item"])
        unit_price_raw = clean_text(out.at[i, "UNIT PRICE"])

        if not item:
            label = ""

            if re.match(r"^(SUB\s*TOTAL|SUBTOTAL|VAT|TOTAL\s+CONTRACTOR\s+PRICE|GRAND\s+TOTAL|TOTAL)\b",
                        no, flags=re.IGNORECASE):
                label = no

            if re.match(r"^(SUB\s*TOTAL|SUBTOTAL|VAT|TOTAL\s+CONTRACTOR\s+PRICE|GRAND\s+TOTAL)\b",
                        unit_price_raw, flags=re.IGNORECASE):
                label = unit_price_raw

            if label:
                out.at[i, "Item"] = label.strip()
                out.at[i, "NO"] = ""
                out.at[i, "UNIT PRICE"] = ""

    return out[required].reset_index(drop=True)



def _normalize_boq_summary_rows(df):
    """
    Normalize legacy summary rows into the canonical BOQ columns.

    Legacy Excel places labels such as "Sub Total :" and "VAT 11%" in the
    Unit Price column and may contain a numeric artifact in NO. The displayed
    meaning is moved to Item so the rest of the existing UI/calculation logic
    can recognize the row correctly.
    """
    if df is None or df.empty:
        return df

    out = df.copy()

    for i in out.index:
        no = clean_text(out.at[i, "NO"])
        item = clean_text(out.at[i, "Item"])
        unit_price_raw = clean_text(out.at[i, "UNIT PRICE"])

        if item:
            continue

        label = ""

        if re.match(
            r"^(SUB\s*TOTAL|SUBTOTAL|VAT|TOTAL\s+CONTRACTOR\s+PRICE|GRAND\s+TOTAL|TOTAL)\b",
            no,
            flags=re.IGNORECASE,
        ):
            label = no

        if re.match(
            r"^(SUB\s*TOTAL|SUBTOTAL|VAT|TOTAL\s+CONTRACTOR\s+PRICE|GRAND\s+TOTAL|TOTAL)\b",
            unit_price_raw,
            flags=re.IGNORECASE,
        ):
            label = unit_price_raw

        if label:
            out.at[i, "Item"] = label.strip()
            out.at[i, "NO"] = ""
            out.at[i, "UNIT PRICE"] = ""

    return out

@st.cache_data(ttl=600, show_spinner=False)
def load_boq_dataframe(charging_type, province_str, template_name="Template BOQ Vgreen Lama", _parser_version=BOQ_TEMPLATE_PARSER_VERSION):
    """
    Load the selected BOQ template exactly as displayed in Excel.

    Important fixes:
    1. Excel formulas are read with their cached/displayed values
       (openpyxl data_only=True), so formula text such as ``=D9*G9`` is
       never accidentally parsed as a price.
    2. OLD template regional blocks are horizontal (JAVA, SUMATERA,
       BALI NUSATENGGARA, KALIMANTAN, SULAWESI).  The previous code shifted
       rows, which mixed unrelated rows.  We now select the correct block.
    3. NEW template uses Qty/Lot and charger-specific Unit Price/Total Price
       headers.  Header mapping is exact-first and never maps Unit/Volume to
       Satuan/Lot.
    4. The loader does NOT recalculate the template.  It mirrors the template
       first.  Existing calculation logic runs only later when the user edits
       Qty/Volume or when totals are requested.
    """
    filename = TEMPLATE_OPTIONS.get(template_name, template_name)
    path = get_template_path(filename)

    if not os.path.exists(path):
        raise FileNotFoundError(f"Template tidak ditemukan: {path}")

    c = normalize_charger_type(charging_type)
    region = map_to_standard_province(province_str)

    # --------------------------------------------------------------------------
    # Read Excel using cached/displayed values.
    # --------------------------------------------------------------------------
    try:
        from openpyxl import load_workbook

        wb = load_workbook(
            filename=path,
            data_only=True,
            read_only=False,
        )
    except Exception as e:
        raise RuntimeError(f"Gagal membaca template Excel '{filename}': {e}")

    try:
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

            if sheet not in wb.sheetnames:
                raise ValueError(
                    f"Sheet '{sheet}' tidak ditemukan di template lama. "
                    f"Sheet tersedia: {', '.join(wb.sheetnames)}"
                )

            ws = wb[sheet]

            # OLD EVCS sheets contain five BOQ blocks side-by-side:
            #   A:G = JAVA
            #   I:O = SUMATERA
            #   Q:W = BALI NUSATENGGARA
            #   Y:AE = KALIMANTAN
            #   AG:AM = SULAWESI
            #
            # Each block has 7 real columns + 1 spacer column.
            block_start = {
                "JAVA": 1,
                "SUMATERA": 9,
                "BALI NUSATENGGARA": 17,
                "KALIMANTAN": 25,
                "SULAWESI": 33,
            }.get(region)

            # BSS old template contains only one 7-column BOQ block.
            if block_start is None:
                block_start = 1

            # Find the header row from the selected block, not from unrelated
            # regional blocks.
            header_row = None
            for r in range(1, min(30, ws.max_row) + 1):
                no_val = clean_text(ws.cell(r, block_start).value).upper()
                item_val = clean_text(ws.cell(r, block_start + 1).value).upper()
                if no_val in {"NO", "NO."} and "ITEM" in item_val:
                    header_row = r
                    break

            if header_row is None:
                # Known old-template layouts.
                header_row = 6 if c in {"DC20", "DC30", "DC60", "DC120"} else 7

            rows = []
            for r in range(header_row, ws.max_row + 1):
                values = [
                    ws.cell(r, block_start + offset).value
                    for offset in range(7)
                ]
                rows.append(values)

            raw = pd.DataFrame(
                rows,
                columns=[
                    "NO",
                    "Item",
                    "VOLUME",
                    "Satuan/Uom",
                    "MERK",
                    "UNIT PRICE",
                    "TOTAL PRICE",
                ],
            )

            # Drop the actual header row from the data.
            if not raw.empty:
                raw = raw.iloc[1:].reset_index(drop=True)

            df = raw.copy()
            df = df.rename(columns={
                "VOLUME": "Unit/Volume",
                "Satuan/Uom": "Satuan/Uom",
            })
            # Keep the canonical seven columns used by the rest of the module.
            df = df[[
                "NO", "Item", "Unit/Volume", "Satuan/Uom",
                "MERK", "UNIT PRICE", "TOTAL PRICE"
            ]]
            df = _normalize_boq_summary_rows(df)
            sheet_return = sheet

        else:
            sheet = get_new_template_sheet_name(c, province_str)

            if sheet not in wb.sheetnames:
                alternatives = [
                    c,
                    c.upper(),
                    c.replace("DC", ""),
                ]
                sheet = next(
                    (x for x in alternatives if x in wb.sheetnames),
                    None,
                )

            if not sheet:
                raise ValueError(
                    f"Sheet template baru untuk charger '{c}' tidak ditemukan."
                )

            ws = wb[sheet]

            # Read complete sheet as displayed/cached values.
            values = list(ws.iter_rows(values_only=True))
            raw = pd.DataFrame(values)

            header_row = detect_template_header(raw)
            df = _standardize_template_df(raw, header_row)
            sheet_return = sheet

    finally:
        try:
            wb.close()
        except Exception:
            pass

    # --------------------------------------------------------------------------
    # Preserve the template values.
    # --------------------------------------------------------------------------
    for col in ["NO", "Item", "Unit/Volume", "Satuan/Uom", "MERK"]:
        df[col] = df[col].apply(clean_text)

    # Cached Excel values are already numeric for normal price cells.
    # Formula cells without a cached result remain blank/zero rather than
    # incorrectly extracting digits from the formula text.
    df["UNIT PRICE"] = df["UNIT PRICE"].apply(parse_price)
    df["TOTAL PRICE"] = df["TOTAL PRICE"].apply(parse_price)

    # IMPORTANT:
    # Do NOT call recalculate_boq_totals() here.
    #
    # The purpose of this function is to mirror the selected Excel template.
    # Existing UI flows call recalculate_boq_totals() only after user edits
    # or when totals are explicitly calculated.
    return df.reset_index(drop=True), sheet_return, region

# ==============================================================================
# EXISTING SECTION EDITOR
# ==============================================================================
def get_section_indices(df_boq, section_no):
    """
    Ambil seluruh baris dalam satu Point/section, bukan hanya header Point.

    Contoh Point A:
        A
        1
        2
        3
        -

    Akan dikembalikan seluruh baris tersebut sampai sebelum Point berikutnya.
    Ini penting supaya nilai Qty/Volume dari template tetap terlihat pada
    editor BOQ biasa.
    """
    if df_boq is None or df_boq.empty or "NO" not in df_boq.columns:
        return []

    parent_labels = {"A", "B", "C", "D", "E", "F", "G", "H"}
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


def edit_boq_sections(
    df_boq,
    charger_type,
    widget_prefix,
):
    """
    Editor BOQ TEMPLATE BARU.

    PENTING:
    - Hanya Point tertentu yang boleh mengubah Qty/Volume:
        EVCS (DC20/DC30/DC60) -> Point A dan G
        BSS  (6S1P/12S1P/12S3P) -> Point A dan D
    - Item, Satuan, MERK, dan UNIT PRICE tetap mengikuti template.
    - TOTAL PRICE dihitung otomatis dari Qty x Unit Price.
    - BOQ MS TIDAK menggunakan function ini.
    """
    if df_boq is None or df_boq.empty:
        return df_boq

    editable_sections = get_editable_sections(charger_type)
    if not editable_sections:
        return df_boq

    df = df_boq.copy()

    # Pandas 3 / ArrowStringArray dapat membaca kolom template sebagai dtype
    # string. Kolom berikut memang akan menerima hasil kalkulasi numeric saat
    # editor dikembalikan, sehingga gunakan object dtype agar aman menerima
    # string maupun float tanpa mengubah nilai/logika BOQ.
    for _col in ["Unit/Volume", "UNIT PRICE", "TOTAL PRICE"]:
        if _col in df.columns:
            df[_col] = df[_col].astype(object)

    st.markdown("### ✏️ Edit Qty / Volume BOQ")
    st.info(
        "💡 Hanya **Qty / Volume** pada Point yang ditentukan yang dapat diubah. "
        "Item, Satuan, MERK, dan UNIT PRICE tetap mengikuti template. "
        "TOTAL PRICE otomatis dihitung = Qty × UNIT PRICE."
    )

    for section_no in editable_sections:
        section_indices = get_section_indices(df, section_no)
        if not section_indices:
            continue

        section_df = df.loc[section_indices].copy()
        st.markdown(f"#### Point {section_no}")

        # Hanya Qty/Volume yang editable. Kolom lain readonly agar nilai
        # bawaan template tidak berubah ketika membuat BOQ biasa.
        editor_columns = [
            "NO",
            "Item",
            "Unit/Volume",
            "Satuan/Uom",
            "MERK",
            "UNIT PRICE",
            "TOTAL PRICE",
        ]
        editor_df = section_df[editor_columns].copy()

        # Pastikan dtype cocok dengan NumberColumn Streamlit.
        # Pertahankan nilai Qty/Volume persis seperti yang ada di template.
        # Jangan cast ke float karena template dapat mempunyai value seperti
        # "30kw dc". parse_qty_num() hanya digunakan oleh calculation engine.
        editor_df["Unit/Volume"] = (
            editor_df["Unit/Volume"]
            .apply(clean_text)
        )
        editor_df["UNIT PRICE"] = (
            editor_df["UNIT PRICE"]
            .apply(parse_price)
            .astype(float)
        )
        editor_df["TOTAL PRICE"] = (
            editor_df["TOTAL PRICE"]
            .apply(parse_price)
            .astype(float)
        )

        column_config = {
            "NO": st.column_config.TextColumn(
                "NO",
                disabled=True,
            ),
            "Item": st.column_config.TextColumn(
                "Item",
                disabled=True,
            ),
            "Unit/Volume": st.column_config.TextColumn(
                "Qty / Volume",
                disabled=False,
                help="Nilai Qty/Volume mengikuti template. Hanya kolom ini yang dapat diubah pada Point yang ditentukan.",
            ),
            "Satuan/Uom": st.column_config.TextColumn(
                "Satuan",
                disabled=True,
            ),
            "MERK": st.column_config.TextColumn(
                "MERK",
                disabled=True,
            ),
            "UNIT PRICE": st.column_config.NumberColumn(
                "UNIT PRICE",
                format="%.0f",
                disabled=True,
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
            key=f"{widget_prefix}_section_{section_no}",
        )

        for position, idx in enumerate(section_indices):
            if position >= len(edited_section):
                continue

            edited_row = edited_section.iloc[position]
            original_qty = df.loc[idx, "Unit/Volume"]

            # IMPORTANT: Streamlit data_editor can return its persisted widget
            # state instead of the freshly loaded Excel value. In older widget
            # states this may be blank. A blank editor value is NOT a user
            # request to erase Qty, so keep the template value.
            edited_qty_raw = edited_row.get("Unit/Volume", "")
            edited_qty_text = clean_text(edited_qty_raw)
            original_qty_text = clean_text(original_qty)

            if edited_qty_text == "" and original_qty_text != "":
                new_qty = original_qty
            else:
                new_qty = parse_qty_num(edited_qty_raw)

            # Section/header rows such as A, G, etc. do not represent an
            # editable material quantity. Never turn their blank Qty into 0.
            no_value = clean_text(df.loc[idx, "NO"]).upper()
            item_value = clean_text(df.loc[idx, "Item"])
            if no_value in {"A", "B", "C", "D", "E", "F", "G", "H"} and original_qty_text == "":
                new_qty = original_qty

            df.loc[idx, "Unit/Volume"] = new_qty

            # Only rows with a real Qty + Unit Price get a newly calculated
            # detail total. Otherwise preserve the template cached total.
            if clean_text(new_qty) != "" and parse_qty_num(new_qty) != 0 and parse_price(df.loc[idx, "UNIT PRICE"]) != 0:
                df.loc[idx, "TOTAL PRICE"] = calculate_detail_total(
                    new_qty,
                    df.loc[idx, "UNIT PRICE"],
                )

            # Never let the editor overwrite readonly template fields.
            # This explicitly preserves Lot/UOM, MERK and UNIT PRICE.
            # (They are also disabled in the data_editor.)
            # No assignment from edited_section is made for these fields.

    # Final calculation tetap menggunakan engine BOQ existing.
    df, _, _, _ = recalculate_boq_totals(df)
    return df

def generate_boq_pdf(site_name,site_location,charger_capacity,region,df_boq,sub_total,vat,grand_total):
    buffer=io.BytesIO()
    doc=SimpleDocTemplate(buffer,pagesize=A4,rightMargin=10,leftMargin=10,topMargin=10,bottomMargin=10)
    elements=[]; styles=getSampleStyleSheet()
    title_style=ParagraphStyle("DocTitle",parent=styles["Heading1"],fontSize=9,leading=10,fontName="Helvetica-Bold",spaceAfter=1)
    sub_style=ParagraphStyle("DocSub",parent=styles["Normal"],fontSize=7,leading=8,fontName="Helvetica",spaceAfter=0)
    table_text=ParagraphStyle("TableText",parent=styles["Normal"],fontSize=5.5,leading=6.5,fontName="Helvetica")
    table_text_bold=ParagraphStyle("TableTextBold",parent=table_text,fontName="Helvetica-Bold")
    table_header=ParagraphStyle("TableHeader",parent=styles["Normal"],fontSize=6,leading=7.5,fontName="Helvetica-Bold",textColor=colors.white)
    elements.append(Paragraph(f"CHARGING WORK {charger_capacity}",title_style))
    elements.append(Paragraph(f"<b>Site Name:</b> {site_name}",sub_style))
    elements.append(Paragraph(f"<b>Site Location:</b> {site_location}",sub_style))
    elements.append(Paragraph(f"<b>NEW PLAN BOQ VGREEN - {region} ISLAND</b>",ParagraphStyle("SubHeader",parent=title_style,fontSize=7.5,leading=8.5,spaceAfter=2,spaceBefore=1)))
    table_data=[[Paragraph(x,table_header) for x in ["NO","Item","Unit/Vol","Satuan","MERK","UNIT PRICE","TOTAL PRICE"]]]
    parents={"A","B","C","D","E","F","G","H"}; pdf_df,_,_,_=recalculate_boq_totals(df_boq)
    for _,row in pdf_df.iterrows():
        no=str(row.get("NO","")).strip().upper(); parent=no in parents; sty=table_text_bold if parent else table_text
        up=parse_price(row.get("UNIT PRICE",0)); tp=parse_price(row.get("TOTAL PRICE",0))
        up_str=format_currency(up) if up!=0 else ("-" if not parent else "")
        tp_str=format_currency(tp) if tp!=0 else "-"
        table_data.append([Paragraph(no,sty),Paragraph(str(row.get("Item","")),sty),Paragraph(str(row.get("Unit/Volume","")),sty),Paragraph(str(row.get("Satuan/Uom","")),sty),Paragraph(str(row.get("MERK","")),sty),Paragraph(up_str,sty),Paragraph(tp_str,sty)])
    table_data += [["","",Paragraph("<b>Sub Total:</b>",table_text_bold),"","","",Paragraph(f"<b>{format_currency(sub_total)}</b>",table_text_bold)], ["","",Paragraph("<b>VAT 11%</b>",table_text_bold),"","","",Paragraph(f"<b>{format_currency(vat)}</b>",table_text_bold)], ["","",Paragraph("<b>Total Contractor Price</b>",table_text_bold),"","","",Paragraph(f"<b>{format_currency(grand_total)}</b>",table_text_bold)]]
    t=Table(table_data,colWidths=[18,260,45,45,67,70,70])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#2C3E50")),("ALIGN",(0,0),(-1,-1),"LEFT"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("GRID",(0,0),(-1,-4),0.3,colors.HexColor("#CCCCCC")),("BACKGROUND",(0,-3),(-1,-1),colors.HexColor("#F8F9F9")),("LINEABOVE",(0,-3),(-1,-3),0.8,colors.HexColor("#2C3E50")),("TOPPADDING",(0,0),(-1,-1),0.5),("BOTTOMPADDING",(0,0),(-1,-1),0.5),("LEFTPADDING",(0,0),(-1,-1),2),("RIGHTPADDING",(0,0),(-1,-1),2)]))
    elements.append(t); doc.build(elements); buffer.seek(0); return buffer.getvalue()

def _period_text(date_obj):
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sept", "Oct", "Nov", "Dec"]
    d = date_obj.date() if isinstance(date_obj, datetime.datetime) else date_obj
    if d.day >= 16:
        m1, y1 = d.month, d.year
        m2 = m1 + 1
        y2 = y1
        if m2 == 13:
            m2, y2 = 1, y1 + 1
    else:
        m2, y2 = d.month, d.year
        m1 = m2 - 1
        y1 = y2
        if m1 == 0:
            m1, y1 = 12, y2 - 1
    return f"{months[m1-1]}-{months[m2-1]} {y2}"


def _ensure_boq_ms_headers(ws):
    headers = BOQ_MS_HEADERS + BOQ_MS_HELPER_HEADERS
    try:
        vals = ws.get_all_values()
        first = vals[0] if vals else []
        if first[:9] != BOQ_MS_HEADERS:
            ws.update("A1:I1", [BOQ_MS_HEADERS])
        # Add helper columns without disturbing A:I.
        vals = ws.get_all_values()
        first = vals[0] if vals else []
        full = BOQ_MS_HEADERS + BOQ_MS_HELPER_HEADERS
        for col_idx, header in enumerate(full[9:], start=10):
            if len(first) < col_idx or str(first[col_idx-1]).strip() != header:
                ws.update_cell(1, col_idx, header)
    except Exception:
        ws.update("A1:O1", [headers])
    return headers

@st.cache_data(ttl=120, show_spinner=False)
def _read_boq_ms_records():
    df = read_sheet_df(DB_BOQ_MS_SHEET)
    if df.empty:
        return []
    cols=list(df.columns)
    def get(row,idx,default=""):
        c=cols[idx] if idx<len(cols) else None
        return row.get(c,default) if c else default
    out=[]
    for i,row in df.iterrows():
        if not any(clean_text(v) for v in row.tolist()): continue
        out.append({
            "row_idx":i+2,"No":clean_text(get(row,0)),"Date":clean_text(get(row,1)),"No. BOQ":clean_text(get(row,2)),
            "SN Machine":clean_text(get(row,3)),"Site Name":clean_text(get(row,4)),"Charging Type":normalize_charger_type(get(row,5)),
            "Item Name":clean_text(get(row,6)),"Price":parse_price(get(row,7)),"Periode":clean_text(get(row,8)),
            "Qty":parse_qty_num(get(row,9)),"Template":clean_text(get(row,10)),"Region":clean_text(get(row,11)),
            "Issue Note":clean_text(get(row,12)),"Photo Evident 1":clean_text(get(row,13)),"Photo Evident 2":clean_text(get(row,14)),
        })
    return out

def _next_boq_ms_sequence():
    max_no = 0
    for r in _read_boq_ms_records():
        m = re.match(r"^(\d+)", r["No"])
        if m:
            max_no = max(max_no, int(m.group(1)))
        m2 = re.match(r"^(\d+)/CLX/BOQ/MS/", r["No. BOQ"].upper())
        if m2:
            max_no = max(max_no, int(m2.group(1)))
    return max_no + 1


def _material_rows_for_ms(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=["Source Row","NO","Item Name","Satuan","Merk","Unit Price"])
    parents = {"A","B","C","D","E","F","G","H"}
    rows = []
    for i, r in df.iterrows():
        no = clean_text(r.get("NO", ""))
        item = clean_text(r.get("Item", ""))
        # A replacement-material picker must not offer section headers or blank rows.
        if not item or no.upper() in parents:
            continue
        rows.append({
            "Source Row": i,
            "NO": no,
            "Item Name": item,
            "Satuan": clean_text(r.get("Satuan/Uom", "")),
            "Merk": clean_text(r.get("MERK", "")),
            "Unit Price": parse_price(r.get("UNIT PRICE", 0)),
        })
    return pd.DataFrame(rows)


def _boq_ms_picker(material_df, selected_keys=None, key="boq_ms_picker"):
    work = material_df.copy()
    work.insert(0, "Pilih", False)
    work.insert(7, "Qty", 1.0)
    work["Total DPP"] = work["Unit Price"] * work["Qty"]
    if selected_keys is not None:
        selected_keys = set(selected_keys)
        work["Pilih"] = work["Source Row"].isin(selected_keys)
    view = work[["Pilih","NO","Item Name","Satuan","Merk","Unit Price","Qty","Total DPP"]].copy()
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
        disabled=["NO","Item Name","Satuan","Merk","Unit Price","Total DPP"],
    )
    edited = edited.copy()
    edited["Qty"] = edited["Qty"].apply(parse_qty_num)
    edited["Total DPP"] = edited["Unit Price"] * edited["Qty"]
    # Map edited rows back to source rows using positional alignment.
    selected = []
    for i, r in edited.iterrows():
        if bool(r["Pilih"]):
            src = material_df.iloc[i]
            selected.append({
                "Source Row": int(src["Source Row"]),
                "NO": clean_text(src["NO"]),
                "Item Name": clean_text(src["Item Name"]),
                "Satuan": clean_text(src["Satuan"]),
                "Merk": clean_text(src["Merk"]),
                "Unit Price": parse_price(src["Unit Price"]),
                "Qty": parse_qty_num(r["Qty"]),
                "Total DPP": calculate_detail_total(r["Qty"], src["Unit Price"]),
            })
    return pd.DataFrame(selected)


def _get_pdf_asset_path(filename):
    """Cari asset PDF pada assets/templates di beberapa root yang digunakan app."""
    candidates = [
        os.path.join(CURRENT_DIR, "assets", "templates", filename),
        os.path.join(ROOT_DIR, "assets", "templates", filename),
        os.path.join(CWD, "assets", "templates", filename),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def _draw_pdf_asset(canvas, path, x, y_top, max_width, max_height):
    """Draw image proportionally without stretching."""
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
        canvas.drawImage(reader, x_draw, y_draw, width=width, height=height, preserveAspectRatio=True, mask="auto")
    except Exception:
        pass


def _boq_ms_pdf_header_footer(canvas, doc):
    """Header/Footer khusus PDF BOQ MS. Tidak mempengaruhi PDF BOQ biasa."""
    canvas.saveState()
    page_width, page_height = portrait(A5)
    left = 10
    usable_width = page_width - 20

    # Prioritas 1: header/footer image yang sudah disiapkan.
    header_path = _get_pdf_asset_path("header.png")
    footer_path = _get_pdf_asset_path("Footer.png")

    if header_path:
        _draw_pdf_asset(canvas, header_path, left, page_height - 4, usable_width, 42)
    else:
        # Fallback agar header TETAP muncul walaupun header.png tidak ditemukan.
        logo_path = _get_pdf_asset_path("CLX.png") or get_logo_path()
        if logo_path:
            _draw_pdf_asset(canvas, logo_path, left, page_height - 7, 95, 34)
        canvas.setFont("Helvetica-Bold", 8)
        canvas.drawRightString(page_width - 10, page_height - 18, "BOQ MS / REPLACEMENT MATERIAL")
        canvas.setFont("Helvetica", 6)
        canvas.drawRightString(page_width - 10, page_height - 28, "PT. Connectivity Leads eXcellence")
        canvas.setLineWidth(0.5)
        canvas.line(left, page_height - 40, page_width - 10, page_height - 40)

    if footer_path:
        _draw_pdf_asset(canvas, footer_path, left, 34, usable_width, 30)
    else:
        # Fallback footer agar footer TETAP muncul walaupun Footer.png tidak ada.
        canvas.setLineWidth(0.5)
        canvas.line(left, 34, page_width - 10, 34)
        canvas.setFont("Helvetica", 6)
        canvas.drawString(left, 22, "PT. Connectivity Leads eXcellence")
        canvas.drawRightString(page_width - 10, 22, f"Page {doc.page}")

    canvas.restoreState()

def generate_boq_ms_pdf(records, boq_no, sn_machine, site_name, charging_type, period, template_name, region, issue_note="", photo_bytes=None, date_obj=None):
    # Margin atas/bawah dibuat cukup agar header/footer tidak bertabrakan dengan isi.
    buf = io.BytesIO()
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
    small = ParagraphStyle(
        "ms_small",
        parent=styles["Normal"],
        fontSize=6.5,
        leading=8,
    )
    bold = ParagraphStyle(
        "ms_bold",
        parent=small,
        fontName="Helvetica-Bold",
    )
    cell = ParagraphStyle(
        "ms_cell",
        parent=small,
        fontSize=6,
        leading=7,
    )
    table_header = ParagraphStyle(
        "ms_table_header",
        parent=small,
        fontName="Helvetica-Bold",
        fontSize=6,
        leading=7,
        alignment=1,
        textColor=colors.white,
    )

    story = []
    dt = date_obj or datetime.datetime.now()

    # Judul tetap dipertahankan, sedangkan logo CLX lama tidak lagi diperlukan
    # karena header.png menjadi header utama dokumen.
    story.append(Paragraph("BOQ MS / REPLACEMENT MATERIAL", title))

    # --------------------------------------------------------------------------
    # INFORMASI DOKUMEN
    # Template dan Region sengaja TIDAK ditampilkan di PDF sesuai permintaan.
    # --------------------------------------------------------------------------
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
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ])
    )
    story += [ti, Spacer(1, 4)]

    # --------------------------------------------------------------------------
    # DETAIL MATERIAL
    # Semua judul kolom dibuat center + putih.
    # --------------------------------------------------------------------------
    data = [[
        Paragraph("No", table_header),
        Paragraph("Item Name", table_header),
        Paragraph("Qty", table_header),
        Paragraph("Unit Price", table_header),
        Paragraph("DPP", table_header),
    ]]
    total = 0.0

    for i, r in enumerate(records, 1):
        qty = parse_qty_num(r.get("Qty", 0))
        up = parse_price(r.get("Unit Price", 0))
        dpp = calculate_detail_total(qty, up)
        total += dpp
        data.append([
            Paragraph(str(i), cell),
            Paragraph(clean_text(r.get("Item Name", "")), cell),
            Paragraph(f"{qty:g}", cell),
            Paragraph(format_currency(up), cell),
            Paragraph(format_currency(dpp), cell),
        ])

    table = Table(
        data,
        colWidths=[22, 171, 35, 76, 76],
        repeatRows=1,
    )
    table.setStyle(
        TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            # Header: seluruh kolom center.
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            # Body: No dan Qty center, angka nominal kanan.
            ("ALIGN", (0, 1), (0, -1), "CENTER"),
            ("ALIGN", (2, 1), (2, -1), "CENTER"),
            ("ALIGN", (3, 1), (-1, -1), "RIGHT"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ])
    )
    story += [table, Spacer(1, 4)]

    # --------------------------------------------------------------------------
    # TOTAL
    # --------------------------------------------------------------------------
    vat = total * 0.11
    grand = total + vat
    sm = Table(
        [
            [Paragraph("DPP", bold), Paragraph(format_currency(total), bold)],
            [Paragraph("PPN 11%", small), Paragraph(format_currency(vat), small)],
            [Paragraph("Grand Total", bold), Paragraph(format_currency(grand), bold)],
        ],
        colWidths=[285, 95],
    )
    sm.setStyle(
        TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ])
    )
    story.append(sm)

    # --------------------------------------------------------------------------
    # ISSUE NOTE - dibuat table agar lebih rapi.
    # --------------------------------------------------------------------------
    story.append(Spacer(1, 5))
    issue_text = clean_text(issue_note).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    issue_text = issue_text.replace("\n", "<br/>") or "-"
    issue_table = Table(
        [
            [Paragraph("ISSUE NOTE", table_header)],
            [Paragraph(issue_text, cell)],
        ],
        colWidths=[380],
    )
    issue_table.setStyle(
        TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ])
    )
    story.append(issue_table)

    # --------------------------------------------------------------------------
    # FOTO EVIDENT - table 2 kolom, foto proporsional.
    # --------------------------------------------------------------------------
    photos = [x for x in (photo_bytes or []) if x]
    if photos:
        story.append(Spacer(1, 5))
        photo_title = Table(
            [[Paragraph("FOTO EVIDENT", table_header)]],
            colWidths=[380],
        )
        photo_title.setStyle(
            TableStyle([
                ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ])
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
            TableStyle([
                ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ])
        )
        story.append(photo_table)

    doc.build(
        story,
        onFirstPage=_boq_ms_pdf_header_footer,
        onLaterPages=_boq_ms_pdf_header_footer,
    )
    buf.seek(0)
    return buf.getvalue()


def _upload_boq_ms_photo(uploaded_file, boq_no):
    """Upload evidence photo to Google Drive and return a shareable URL.
    Falls back to an empty string if Drive API is not available.
    """
    if uploaded_file is None:
        return ""
    try:
        sh=get_google_sheet_connection()
        creds=getattr(sh,"auth",None) if sh else None
        if creds is None:
            return ""
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload
        drive=build("drive","v3",credentials=creds,cache_discovery=False)
        q="name='BOQ MS Evidence' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res=drive.files().list(q=q,spaces="drive",fields="files(id,name)",pageSize=10).execute()
        files=res.get("files",[])
        if files:
            folder_id=files[0]["id"]
        else:
            folder=drive.files().create(body={"name":"BOQ MS Evidence","mimeType":"application/vnd.google-apps.folder"},fields="id").execute()
            folder_id=folder["id"]
        safe=re.sub(r'[^A-Za-z0-9._-]+','_',uploaded_file.name)
        filename=f"{boq_no.replace('/','-')}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{safe}"
        media=MediaIoBaseUpload(io.BytesIO(uploaded_file.getvalue()),mimetype=uploaded_file.type or "image/jpeg",resumable=False)
        created=drive.files().create(body={"name":filename,"parents":[folder_id]},media_body=media,fields="id,webViewLink").execute()
        file_id=created["id"]
        try:
            drive.permissions().create(fileId=file_id,body={"type":"anyone","role":"reader"}).execute()
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
        sh=get_google_sheet_connection(); creds=getattr(sh,"auth",None) if sh else None
        if creds is None: return None
        m=re.search(r"/d/([A-Za-z0-9_-]+)",url) or re.search(r"[?&]id=([A-Za-z0-9_-]+)",url)
        if not m: return None
        from googleapiclient.discovery import build
        drive=build("drive","v3",credentials=creds,cache_discovery=False)
        from googleapiclient.http import MediaIoBaseDownload
        req=drive.files().get_media(fileId=m.group(1)); fh=io.BytesIO(); downloader=MediaIoBaseDownload(fh,req)
        done=False
        while not done: _,done=downloader.next_chunk()
        fh.seek(0); return fh.getvalue()
    except Exception:
        return None

def _save_boq_ms(records, sn_machine, site_name, charging_type, template_name, region, issue_note="", photo_files=None, date_obj=None):
    ws=_get_ws(DB_BOQ_MS_SHEET)
    if ws is None:
        raise RuntimeError(f"Worksheet '{DB_BOQ_MS_SHEET}' tidak ditemukan. Buat worksheet tersebut terlebih dahulu.")
    _ensure_boq_ms_headers(ws)
    dt=date_obj or datetime.datetime.now(); seq=_next_boq_ms_sequence(); boq_no=generate_boq_ms_number(seq,dt); period=_period_text(dt)
    photo_files=photo_files or []
    photo_urls=[]
    for f in photo_files[:2]:
        u=_upload_boq_ms_photo(f,boq_no)
        photo_urls.append(u)
    while len(photo_urls)<2: photo_urls.append("")
    rows=[]
    for item in records:
        rows.append([seq,dt.strftime("%d-%m-%Y"),boq_no,clean_text(sn_machine),clean_text(site_name),normalize_charger_type(charging_type),clean_text(item["Item Name"]),parse_price(item["Total DPP"]),period,parse_qty_num(item["Qty"]),template_db_value(template_name),clean_text(region),clean_text(issue_note),photo_urls[0],photo_urls[1]])
    if rows: ws.append_rows(rows,value_input_option="USER_ENTERED")
    return boq_no,period,seq,photo_urls

def _update_boq_ms(boq_no, records, sn_machine, site_name, charging_type, template_name, region, issue_note="", photo_files=None, date_text=None, existing_photo_urls=None):
    ws=_get_ws(DB_BOQ_MS_SHEET)
    if ws is None: raise RuntimeError(f"Worksheet '{DB_BOQ_MS_SHEET}' tidak ditemukan.")
    old=[r for r in _read_boq_ms_records() if r["No. BOQ"]==boq_no]
    if not old: raise RuntimeError("BOQ MS tidak ditemukan.")
    seq=parse_qty_num(old[0]["No"]); dt=datetime.datetime.now()
    try:
        if date_text: dt=datetime.datetime.strptime(date_text,"%d-%m-%Y")
    except Exception: pass
    period=_period_text(dt); existing_photo_urls=existing_photo_urls or [old[0].get("Photo Evident 1", ""),old[0].get("Photo Evident 2", "")]
    uploaded_urls=[]
    for f in (photo_files or [])[:2]:
        u=_upload_boq_ms_photo(f,boq_no)
        if u: uploaded_urls.append(u)
    photos=(uploaded_urls + existing_photo_urls)[:2]
    while len(photos)<2: photos.append("")
    # Clear only the old BOQ rows; unrelated DB BOQ MS records are untouched.
    for r in old:
        ws.update(f"A{r['row_idx']}:O{r['row_idx']}", [[""]*15])
    new_rows=[]
    for item in records:
        new_rows.append([int(seq),dt.strftime("%d-%m-%Y"),boq_no,clean_text(sn_machine),clean_text(site_name),normalize_charger_type(charging_type),clean_text(item["Item Name"]),parse_price(item["Total DPP"]),period,parse_qty_num(item["Qty"]),template_db_value(template_name),clean_text(region),clean_text(issue_note),photos[0],photos[1]])
    if not new_rows: return period
    ws.update(f"A{old[0]['row_idx']}:O{old[0]['row_idx']}",[new_rows[0]],value_input_option="USER_ENTERED")
    if len(new_rows)>1: ws.insert_rows(new_rows[1:],row=old[0]['row_idx']+1,value_input_option="USER_ENTERED")
    return period

def _get_boq_ms_grouped():
    records = _read_boq_ms_records()
    groups = {}
    for r in records:
        groups.setdefault(r["No. BOQ"], []).append(r)
    return groups


def _records_to_pdf_items(records):
    return [{
        "Item Name": r["Item Name"], "Qty": r["Qty"], "Unit Price": r["Price"] / r["Qty"] if r["Qty"] else 0,
        "Total DPP": r["Price"]
    } for r in records]


# ==============================================================================
# BOQ MS UI
# ==============================================================================
def render_boq_ms_create():
    # Reset counter khusus CREATE BOQ MS.
    # Setelah berhasil save, counter dinaikkan lalu st.rerun() sehingga
    # seluruh widget Create BOQ MS dibuat ulang dalam kondisi default/blank.
    reset_no = int(st.session_state.get("boq_ms_create_reset_no", 0))
    key_suffix = f"_{reset_no}"

    st.subheader("➕ Create BOQ MS / Replacement Material")
    st.info("Pilih template, charging type, lalu pilih material replacement. Qty hanya berlaku untuk material yang dipilih.")
    template_name=st.selectbox("Template",list(TEMPLATE_OPTIONS.keys()),index=1,key=f"ms_create_template{key_suffix}")
    charging_type=st.selectbox("Charging Type",["6S1P","12S1P","DC20","DC30","DC60"],key=f"ms_create_charger{key_suffix}")
    region=st.selectbox("Region / Island",["JAVA","SUMATERA","BALI NUSATENGGARA","KALIMANTAN","SULAWESI"],key=f"ms_create_region{key_suffix}")
    c1,c2=st.columns(2); sn_machine=c1.text_input("SN Machine",key=f"ms_create_sn{key_suffix}"); site_name=c2.text_input("Site Name",key=f"ms_create_site{key_suffix}")
    issue_note=st.text_area("Issue Note",placeholder="Tuliskan issue / alasan replacement material...",key=f"ms_create_issue_note{key_suffix}",height=90)
    photos=st.file_uploader("📷 Foto Evident (Maks. 2 Foto)",type=["jpg","jpeg","png","webp"],accept_multiple_files=True,key=f"ms_create_photos{key_suffix}",help="Upload maksimal 2 foto evidence.")
    if len(photos)>2:
        st.error("❌ Maksimal 2 foto evident. Silakan hapus foto yang berlebih."); return
    if not site_name: st.caption("Isi Site Name terlebih dahulu untuk melanjutkan."); return
    df_template, _ms_sheet, _ms_region=load_boq_dataframe(charging_type,region,template_name)
    if df_template is None or df_template.empty: return
    materials=_material_rows_for_ms(df_template)
    if materials.empty: st.warning("Tidak ada material detail yang ditemukan pada template/sheet tersebut."); return
    st.markdown("#### Pilih Replacement Material")
    selected_df=_boq_ms_picker(materials,key=f"boq_ms_create_picker{key_suffix}")
    if selected_df.empty: st.warning("Belum ada material yang dipilih."); return
    total_dpp=float(selected_df["Total DPP"].sum()); vat=total_dpp*0.11; grand=total_dpp+vat; today=datetime.datetime.now(); period=_period_text(today)
    c1,c2,c3=st.columns(3); c1.metric("Item",len(selected_df)); c2.metric("DPP",format_currency(total_dpp)); c3.metric("Grand Total",format_currency(grand)); st.write(f"**Periode:** {period}")
    preview_boq_no=generate_boq_ms_number(_next_boq_ms_sequence(),today)
    pdf=generate_boq_ms_pdf(selected_df.to_dict("records"),preview_boq_no,sn_machine,site_name,charging_type,period,template_name,region,issue_note,[f.getvalue() for f in photos],today)
    st.download_button("📄 Preview / Download PDF",pdf,file_name="BOQ_MS_PREVIEW.pdf",mime="application/pdf",key=f"ms_preview_pdf{key_suffix}")
    if st.button("💾 Save BOQ MS",type="primary",key=f"ms_save{key_suffix}"):
        try:
            boq_no,period,seq,urls=_save_boq_ms(selected_df.to_dict("records"),sn_machine,site_name,charging_type,template_name,region,issue_note,photos,today)
            pdf=generate_boq_ms_pdf(selected_df.to_dict("records"),boq_no,sn_machine,site_name,charging_type,period,template_name,region,issue_note,[f.getvalue() for f in photos],today)
            st.cache_data.clear()
            st.session_state["boq_ms_last_pdf"]=pdf; st.session_state["boq_ms_last_no"]=boq_no
            st.session_state["boq_ms_create_reset_no"] = reset_no + 1
            st.success(f"BOQ MS {boq_no} berhasil disimpan. Form Create BOQ MS dikosongkan kembali.")
            st.rerun()
        except Exception as e: st.error(f"Gagal menyimpan BOQ MS: {e}")

def render_boq_ms_edit():
    st.subheader("✏️ Edit / Re-Download BOQ MS")
    groups=_get_boq_ms_grouped()
    if not groups: st.info("Belum ada data pada DB BOQ MS."); return
    keys=list(groups.keys()); labels=[f"{k} — {groups[k][0]['Site Name']} — {groups[k][0]['Charging Type']}" for k in keys]; label=st.selectbox("Pilih BOQ MS",labels,key="ms_edit_select"); boq_no=keys[labels.index(label)]; records=groups[boq_no]; first=records[0]
    template_name=template_name_from_db_value(first.get("Template","BOQ BARU")); region=first.get("Region") or "JAVA"; charger=normalize_charger_type(first["Charging Type"])
    sn=st.text_input("SN Machine",first.get("SN Machine",""),key="ms_edit_sn"); site=st.text_input("Site Name",first.get("Site Name",""),key="ms_edit_site")
    template_name=st.selectbox("Template",list(TEMPLATE_OPTIONS.keys()),index=list(TEMPLATE_OPTIONS.keys()).index(template_name),key="ms_edit_template")
    charger=st.selectbox("Charging Type",["6S1P","12S1P","DC20","DC30","DC60"],index=["6S1P","12S1P","DC20","DC30","DC60"].index(charger) if charger in ["6S1P","12S1P","DC20","DC30","DC60"] else 0,key="ms_edit_charger")
    ropts=["JAVA","SUMATERA","BALI NUSATENGGARA","KALIMANTAN","SULAWESI"]; region=st.selectbox("Region / Island",ropts,index=ropts.index(region) if region in ropts else 0,key="ms_edit_region")
    issue_note=st.text_area("Issue Note",first.get("Issue Note",""),key="ms_edit_issue_note",height=90)
    existing_urls=[first.get("Photo Evident 1",""),first.get("Photo Evident 2","")]
    active_existing=[u for u in existing_urls if u]
    if active_existing: st.caption("📎 Existing Foto Evident: " + " | ".join(active_existing))
    new_photos=st.file_uploader("📷 Ganti/Tambah Foto Evident (Maks. 2 Foto)",type=["jpg","jpeg","png","webp"],accept_multiple_files=True,key="ms_edit_photos")
    if len(new_photos)>2: st.error("❌ Maksimal 2 foto evident."); return
    df_template, _ms_sheet, _ms_region=load_boq_dataframe(charger,region,template_name)
    if df_template is None or df_template.empty: return
    materials=_material_rows_for_ms(df_template); existing_by_name={}
    for r in records: existing_by_name.setdefault(r["Item Name"],[]).append(r)
    selected_keys=[int(m["Source Row"]) for _,m in materials.iterrows() if m["Item Name"] in existing_by_name]
    selected_df=_boq_ms_picker(materials,selected_keys=selected_keys,key=f"boq_ms_edit_picker_{boq_no}")
    # Restore saved quantities after the picker while preserving the selected rows.
    for i,r in selected_df.iterrows():
        matches=existing_by_name.get(r["Item Name"],[])
        if matches:
            q=matches[0].get("Qty",0); selected_df.at[i,"Qty"]=q; selected_df.at[i,"Total DPP"]=calculate_detail_total(q,r["Unit Price"])
    if selected_df.empty: st.warning("Belum ada material yang dipilih."); return
    total_dpp=float(selected_df["Total DPP"].sum()); vat=total_dpp*0.11; grand=total_dpp+vat
    c1,c2,c3=st.columns(3); c1.metric("Item",len(selected_df)); c2.metric("DPP",format_currency(total_dpp)); c3.metric("Grand Total",format_currency(grand))
    try: dt=datetime.datetime.strptime(first.get("Date",""),"%d-%m-%Y")
    except Exception: dt=datetime.datetime.now()
    period=_period_text(dt)
    # Existing evidence is loaded for PDF preview if available. New uploads replace/add to it.
    photo_bytes=[f.getvalue() for f in new_photos[:2]]
    if not photo_bytes:
        for u in active_existing[:2]:
            b=_download_boq_ms_photo(u)
            if b: photo_bytes.append(b)
    pdf=generate_boq_ms_pdf(selected_df.to_dict("records"),boq_no,sn,site,charger,period,template_name,region,issue_note,photo_bytes,dt)
    a,b=st.columns(2)
    with a:
        if st.button("💾 Update BOQ MS",type="primary",key=f"ms_update_{boq_no}"):
            try:
                _update_boq_ms(boq_no,selected_df.to_dict("records"),sn,site,charger,template_name,region,issue_note,new_photos,first.get("Date",""),existing_urls)
                st.cache_data.clear()
                st.success(f"BOQ MS {boq_no} berhasil di-update."); st.rerun()
            except Exception as e: st.error(f"Gagal update BOQ MS: {e}")
    with b: st.download_button("⬇️ Re-Download PDF",pdf,file_name=f"{boq_no.replace('/','-')}.pdf",mime="application/pdf",key=f"ms_redownload_{boq_no}")

def render_boq_ms():
    st.header("🧰 BOQ MS — Replacement Material")
    tab1, tab2 = st.tabs(["➕ Create BOQ MS", "✏️ Edit / Re-Download BOQ MS"])
    with tab1:
        render_boq_ms_create()
    with tab2:
        render_boq_ms_edit()


# ==============================================================================
# MAIN BOQ MANAGER (EXISTING FEATURES)
# ==============================================================================
def _render_boq_create():
    st.subheader("➕ Buat BOQ Baru")
    if st.session_state.pop("boq_saved_success", False): st.success("BOQ berhasil disimpan.")
    site_options, site_data_map = fetch_query_site_options(exclude_saved=True)
    PLACEHOLDER_OPTION="-- Pilih Site Name & Charger --"
    template_name=st.selectbox("Template yang digunakan",list(TEMPLATE_OPTIONS.keys()),index=1,key="create_boq_template")
    selected_label=st.selectbox("Pilih Site Name & Charger (Filtered: Drop/Cancel & Saved BOQ Pair Excluded)",[PLACEHOLDER_OPTION]+sorted(site_options),index=0,key="create_site_selector")
    if selected_label==PLACEHOLDER_OPTION:
        st.info("💡 Silakan pilih **Site Name & Charger** dari dropdown di atas."); return
    meta=site_data_map[selected_label]; selected_site=meta.get("Site Name",selected_label)
    c1,c2=st.columns(2)
    site_address=c2.text_input("Address (Kolom G)",value=meta.get("Address","-"),key=f"address_{selected_label}")
    c3,c4,c5=st.columns([1.5,1.5,1])
    charging_type=c3.text_input("Charging Type (Kolom C)",value=meta.get("Charging Type","DC20"),key=f"charger_{selected_label}")
    province=c4.text_input("Province Standar (Kolom I Mapped)",value=meta.get("Province","JAVA"),key=f"province_{selected_label}")
    c5.caption(f"Raw Province Sheet: `{meta.get('Province','-')}`")
    epc_name=st.text_input("EPC Name (Kolom B)",value=meta.get("EPC Name","-"),key=f"epc_{selected_label}")
    saved_pairs=get_existing_saved_site_charger_pairs(); pair=(selected_site.strip().lower(),charging_type.strip().lower()); is_already_saved=pair in saved_pairs or selected_site=="-"
    if is_already_saved and selected_site!="-": st.warning(f"⚠️ Kombinasi Site `{selected_site}` dengan Charger `{charging_type}` sudah pernah dibuatkan BOQ.")
    df_boq,target_sheet,region_normalized=load_boq_dataframe(charging_type,province,template_name)
    if df_boq is None or df_boq.empty: return
    st.subheader(f"2. Table BOQ ({target_sheet} - {region_normalized})")
    st.caption(f"📑 Template aktif: **{template_name}**")
    if template_name=="Template BOQ Vgreen Lama":
        st.info("💡 Khusus **CABLING AND ACCESSORIES INSTALLATION** (Poin 1-3), Qty/Volume dapat diubah. TOTAL PRICE akan otomatis mengikuti Qty × UNIT PRICE.")
        e_start=False; editable_indices=[]
        for idx,row in df_boq.iterrows():
            no=str(row.get("NO","")).strip().upper(); item=str(row.get("Item","")).upper()
            if "CABLING AND ACCESSORIES" in item or no=="E": e_start=True; continue
            elif no in ["A","B","C","D","F"]: e_start=False
            if e_start and no in ["1","2","3"]: editable_indices.append(idx)
        cols=st.columns(3)
        for idx in editable_indices:
            no=str(df_boq.loc[idx,"NO"]); item=str(df_boq.loc[idx,"Item"]); q=parse_qty_num(df_boq.loc[idx,"Unit/Volume"]); target=cols[0 if no=="1" else 1 if no=="2" else 2]
            with target:
                nq=st.number_input(f"Qty Poin {no}: {item[:25]}...",min_value=0.0,value=float(q),step=1.0,key=f"qty_e_{no}_{selected_label}_{template_name}")
                df_boq.loc[idx,"Unit/Volume"]=str(int(nq)) if nq.is_integer() else nq
                df_boq.loc[idx,"TOTAL PRICE"]=calculate_detail_total(df_boq.loc[idx,"Unit/Volume"],df_boq.loc[idx,"UNIT PRICE"])
    else:
        df_boq=edit_boq_sections(df_boq,charging_type,f"create_new_template_{selected_label}_{normalize_charger_type(charging_type)}")
    df_boq,sub_total,vat_amount,grand_total=recalculate_boq_totals(df_boq)
    display=df_boq.copy()
    for c in ["UNIT PRICE","TOTAL PRICE"]: display[c]=display[c].apply(lambda x:format_currency(x) if parse_price(x)!=0 else "-")
    st.dataframe(display,use_container_width=True,hide_index=True)
    c1,c2=st.columns([2,1]); c1.success(f"✅ Tabel BOQ Aktif: **{selected_site}** | Tipe Charger: **{charging_type}** | Wilayah: **{region_normalized}**"); c2.metric("Total Contractor Price (Inc. VAT 11%)",format_currency(grand_total))
    pdf_bytes=generate_boq_pdf(selected_site,site_address,charging_type,region_normalized,df_boq,sub_total,vat_amount,grand_total)
    safe_site=re.sub(r'[\\/*?:"<>|]','_',str(selected_site)); safe_charger=re.sub(r'[\\/*?:"<>|]','_',str(charging_type)); filename_pdf=f"BOQ_{safe_charger}_{region_normalized}_{safe_site}.pdf"
    b1,b2=st.columns(2)
    with b1:
        if st.button("🚀 Simpan ke Database (DB BOQ)",type="primary",disabled=is_already_saved,key=f"btn_save_boq_{selected_label}_{template_name}"):
            boq_no=save_to_db_boq(selected_site,charging_type,sub_total,grand_total,epc_name,template_name)
            if boq_no:
                update_google_sheet_summary(selected_site,selected_site,sub_total,grand_total); st.cache_data.clear(); st.session_state["last_saved_info"]={"boq_no":boq_no,"site_name":selected_site,"pdf_bytes":pdf_bytes,"filename":filename_pdf}; st.rerun()
    with b2: st.download_button("📥 Download PDF BOQ (Draft Preview)",pdf_bytes,file_name=filename_pdf,mime="application/pdf",key=f"dl_active_{selected_site}_{template_name}")

def _render_boq_edit():
    st.subheader("✏️ Edit, Reuse, & Re-Download BOQ Tersimpan")
    saved_boq_list=get_all_saved_boq()
    if not saved_boq_list: st.info("ℹ️ Belum ada data BOQ yang tersimpan di `DB BOQ`."); return
    options={f"{x.get('BOQ No.','-')}: {x.get('Site Name','-')} [{x.get('Charger Type','-')}]":x for x in saved_boq_list}
    key=st.selectbox("Pilih Nomor BOQ yang Ingin Di-edit / Re-assign / Re-Download",sorted(options),key="edit_boq_selector"); data=options[key]
    old_site=data.get("Site Name",""); old_charger=data.get("Charger Type",""); stored=data.get("Template",""); template_name=template_name_from_db_value(stored)
    template_name=st.selectbox("Pilih Template untuk Edit / Re-Download PDF",list(TEMPLATE_OPTIONS.keys()),index=list(TEMPLATE_OPTIONS.keys()).index(template_name),key=f"edit_boq_template_{key}")
    _,qmap=fetch_query_site_options(exclude_saved=False)
    meta=next((v for v in qmap.values() if v.get("Site Name","").strip().lower()==old_site.strip().lower() and normalize_charger_type(v.get("Charging Type",""))==normalize_charger_type(old_charger)),{"Address":"-","Province":"JAVA","EPC Name":"-"})
    site=st.text_input("Site Name",old_site,key=f"edit_site_{key}"); charger=st.text_input("Charger Type",old_charger,key=f"edit_charger_{key}"); epc=st.text_input("EPC Name",data.get("EPC Name","-"),key=f"edit_epc_{key}"); address=st.text_input("Address",meta.get("Address","-"),key=f"edit_address_{key}"); province=st.text_input("Province",meta.get("Province","JAVA"),key=f"edit_province_{key}")
    df,target,region=load_boq_dataframe(charger,province,template_name)
    if df is None or df.empty: return
    df=edit_boq_sections(df,charger,f"edit_detail_{key}_{template_name}_{charger}"); df,sub,vat,grand=recalculate_boq_totals(df)
    display=df.copy()
    for c in ["UNIT PRICE","TOTAL PRICE"]: display[c]=display[c].apply(lambda x:format_currency(x) if parse_price(x)!=0 else "-")
    st.dataframe(display,use_container_width=True,hide_index=True)
    c1,c2,c3=st.columns(3); c1.metric("Sub Total / Exc. PPN",format_currency(sub)); c2.metric("PPN 11%",format_currency(vat)); c3.metric("Grand Total / Inc. PPN",format_currency(grand))
    pdf=generate_boq_pdf(site,address,charger,region,df,sub,vat,grand)
    a,b=st.columns(2)
    with a:
        if st.button("💾 Save & Update BOQ Database",type="primary",key=f"btn_update_boq_{key}"):
            if update_db_boq_row(data["row_idx"],old_site,site,charger,sub,grand,epc,template_name): st.cache_data.clear(); st.rerun()
    with b: st.download_button(f"📥 Re-Download PDF BOQ ({data.get('BOQ No.','-')})",pdf,file_name=f"BOQ_{str(data.get('BOQ No.','-')).replace('/','_')}.pdf",mime="application/pdf",key=f"download_edit_{key}_{template_name}")


def render():
    st.title("📝 Quotation & BOQ Manager")
    tab_create, tab_edit, tab_ms = st.tabs([
        "➕ Buat BOQ Baru", "✏️ Edit / Reuse / Re-Download BOQ", "🧰 Create BOQ MS"
    ])
    with tab_create:
        _render_boq_create()
    with tab_edit:
        _render_boq_edit()
    with tab_ms:
        render_boq_ms()


if __name__ == "__main__":
    render()
