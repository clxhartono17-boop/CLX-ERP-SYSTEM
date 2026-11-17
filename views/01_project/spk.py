import datetime
import io
import os
import re
from typing import List
from xml.sax.saxutils import escape as xml_escape

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as pdf_canvas
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

CSO_PIN_SECRET = os.getenv("CLX_CSO_PIN", "clx0007")

SHEET_QUERY = "Query"
SHEET_DROPDOWN = "Master Dropdown"
SHEET_SOW = "Master SOW"

SHEET_SURVEY = "DB SPK Survey"
SHEET_CONS = "DB SPK Cons"
SHEET_MS = "DB SPK MS"

STATUS_PENDING = "Pending CSO Approval"
STATUS_APPROVED = "Approved by CSO"
STATUS_REJECTED = "Rejected by CSO"

# Backward compatibility agar data lama berstatus COO tetap terbaca
LEGACY_PENDING_STATUSES = {
    STATUS_PENDING.lower(),
    "pending coo approval",
    "pending approval",
    "pending",
}
LEGACY_APPROVED_STATUSES = {
    STATUS_APPROVED.lower(),
    "approved by coo",
    "approved",
}

CACHE_TTL_SHEET = 120

# Konfigurasi Penandatangan
VM_SIGNER_NAME = "Vendor Management"
VM_SIGNER_TITLE = "PT. Connectivity Leads eXcellence"

CSO_SIGNER_NAME = "Chief Services Officer"
CSO_SIGNER_TITLE = "PT. Connectivity Leads eXcellence"

# Path Asset Logo & Digital Signature
LOGO_FILENAME_CANDIDATES = ["CLX.png", "clx.png"]
VM_SIGNATURE_FILENAME_CANDIDATES = ["Sign VM.jpg", "sign vm.jpg", "Sign VM.png"]
CSO_SIGNATURE_FILENAME_CANDIDATES = ["Sign CSO.jpg", "sign cso.jpg", "Sign CSO.png"]


# ==============================================================================
# ASSET & SIGNATURE RESOLVERS
# ==============================================================================

def _candidate_asset_dirs() -> List[str]:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.abspath(os.path.join(base_dir, "../"))
    root_dir = os.path.abspath(os.path.join(base_dir, "../../"))
    cwd = os.getcwd()

    raw_dirs = [
        os.path.join(cwd, "assets"),
        os.path.join(base_dir, "assets"),
        os.path.join(parent_dir, "assets"),
        os.path.join(root_dir, "assets"),
        os.path.join(cwd, "assets", "templates"),
        os.path.join(base_dir, "assets", "templates"),
        os.path.join(parent_dir, "assets", "templates"),
        os.path.join(root_dir, "assets", "templates"),
    ]

    seen = set()
    unique_dirs = []
    for d in raw_dirs:
        norm = os.path.abspath(d)
        if norm not in seen:
            seen.add(norm)
            unique_dirs.append(norm)
    return unique_dirs


def resolve_asset_file(filename_candidates: List[str]):
    target_lower = {name.lower() for name in filename_candidates}
    asset_dirs = _candidate_asset_dirs()

    # 1. Exact filename match
    for directory in asset_dirs:
        for name in filename_candidates:
            candidate = os.path.join(directory, name)
            if os.path.isfile(candidate):
                return candidate

    # 2. Case-insensitive scan di folder assets
    for directory in asset_dirs:
        if os.path.isdir(directory):
            try:
                for entry in os.listdir(directory):
                    if entry.lower() in target_lower:
                        full_path = os.path.join(directory, entry)
                        if os.path.isfile(full_path):
                            return full_path
            except Exception:
                pass

    return None


def get_company_logo_path():
    return resolve_asset_file(LOGO_FILENAME_CANDIDATES)


def get_vm_signature_path():
    return resolve_asset_file(VM_SIGNATURE_FILENAME_CANDIDATES)


def get_cso_signature_path():
    return resolve_asset_file(CSO_SIGNATURE_FILENAME_CANDIDATES)


def _build_proportional_image(
    image_path,
    max_width=1.45 * inch,
    max_height=0.62 * inch,
    h_align="CENTER",
):
    if not image_path or not os.path.isfile(image_path):
        return None
    try:
        reader = ImageReader(image_path)
        iw, ih = reader.getSize()
        if iw <= 0 or ih <= 0:
            return None

        scale = min(float(max_width) / float(iw), float(max_height) / float(ih))
        width = iw * scale
        height = ih * scale

        img = Image(image_path, width=width, height=height)
        img.hAlign = h_align
        return img
    except Exception:
        return None


def build_vm_signature_image(max_width=1.40 * inch, max_height=0.60 * inch):
    return _build_proportional_image(
        get_vm_signature_path(),
        max_width=max_width,
        max_height=max_height,
        h_align="CENTER",
    )


def build_cso_signature_image(max_width=1.40 * inch, max_height=0.60 * inch):
    return _build_proportional_image(
        get_cso_signature_path(),
        max_width=max_width,
        max_height=max_height,
        h_align="CENTER",
    )


# ==============================================================================
# MAINTENANCE SOW & STANDARD HEADERS
# ==============================================================================

MAINTENANCE_SOW_EVCS = "Maintenance Service EVCS"
MAINTENANCE_SOW_BSS = "Maintenance Service BSS"

MAINTENANCE_SOW_LIST = [
    MAINTENANCE_SOW_EVCS,
    MAINTENANCE_SOW_BSS,
]

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

DEFAULT_MS_HEADERS = [
    "Date SPK",
    "No. SPK",
    "Pekerjaan",
    "Mitra",
    "Periode Start",
    "Periode End",
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
# CONNECTION & GOOGLE SHEETS CACHE
# ==============================================================================

@st.cache_resource(show_spinner=False)
def get_cached_connection():
    return get_google_sheet_connection()


@st.cache_data(ttl=CACHE_TTL_SHEET, show_spinner=False)
def load_sheet_values_cached(sheet_name: str) -> List[List[str]]:
    sh = get_cached_connection()
    if not sh:
        return []
    worksheet = sh.worksheet(sheet_name)
    return worksheet.get_all_values()


def clear_sheet_cache(sheet_name: str = None):
    try:
        load_sheet_values_cached.clear()
    except Exception:
        pass


def get_worksheet(sheet_name: str):
    sh = get_cached_connection()
    if not sh:
        return None
    return sh.worksheet(sheet_name)


# ==============================================================================
# GENERIC & STATUS HELPERS
# ==============================================================================

def normalize_header(value):
    if value is None:
        return ""
    text = str(value).replace("\n", " ").replace("\r", " ").strip()
    return " ".join(text.split())


def normalize_dataframe_headers(df):
    if df is None:
        return df

    df = df.copy()
    normalized = [normalize_header(col) for col in df.columns]
    seen = {}
    unique_headers = []

    for index, header in enumerate(normalized, start=1):
        base = header or f"Column_{index}"
        count = seen.get(base.lower(), 0) + 1
        seen[base.lower()] = count
        if count == 1:
            unique_headers.append(base)
        else:
            candidate = f"{base} ({count})"
            while candidate.lower() in seen:
                count += 1
                seen[base.lower()] = count
                candidate = f"{base} ({count})"
            seen[candidate.lower()] = 1
            unique_headers.append(candidate)

    df.columns = unique_headers
    return df


def find_column(df, candidates, fallback=None):
    if df is None or df.empty:
        return fallback

    normalized_columns = {
        normalize_header(col).lower(): col
        for col in df.columns
    }

    for candidate in candidates:
        candidate_norm = normalize_header(candidate).lower()
        if candidate_norm in normalized_columns:
            return normalized_columns[candidate_norm]

    for candidate in candidates:
        candidate_norm = normalize_header(candidate).lower()
        for normalized_col, original_col in normalized_columns.items():
            if candidate_norm in normalized_col or normalized_col in candidate_norm:
                return original_col

    return fallback


SHEET_ERROR_VALUES = {
    "#VALUE!",
    "#REF!",
    "#N/A",
    "#DIV/0!",
    "#NAME?",
    "#NUM!",
    "#NULL!",
    "#ERROR!",
}


def is_sheet_error_value(value):
    text = safe_str(value).strip().upper()
    return text in SHEET_ERROR_VALUES


def safe_str(value, default=""):
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except Exception:
        pass

    text = str(value).strip()
    if text.lower() in {"nan", "none"}:
        return default
    if text.upper() in SHEET_ERROR_VALUES:
        return default
    return text


def _normalize_status_text(value):
    return " ".join(safe_str(value).lower().split())


def _is_approved_status(status_val):
    s = _normalize_status_text(status_val)
    return s in LEGACY_APPROVED_STATUSES or s.startswith("approved")


def _is_rejected_status(status_val):
    s = _normalize_status_text(status_val)
    return s in {
        STATUS_REJECTED.lower(),
        "rejected by coo",
        "rejected by cso",
        "rejected",
        "reject",
    } or s.startswith("rejected")


def _is_pending_status(status_val, no_spk_val=""):
    if _is_approved_status(status_val) or _is_rejected_status(status_val):
        return False
    s = _normalize_status_text(status_val)
    if not s:
        spk_clean = safe_str(no_spk_val)
        return bool(spk_clean) and spk_clean != "-"
    return (
        s in LEGACY_PENDING_STATUSES
        or "pending" in s
        or "waiting" in s
        or "menunggu" in s
        or s == "draft"
    )


def esc_pdf(value, default="-"):
    """Escape XML special chars (&, <, >) agar ReportLab Paragraph tidak crash."""
    text = safe_str(value, default)
    if not text:
        text = default
    return xml_escape(text)


def dataframe_from_sheet_rows(rows):
    if not rows or len(rows) <= 1:
        return pd.DataFrame()

    headers = [normalize_header(h) for h in rows[0]]
    max_cols = max(
        len(headers),
        max([len(row) for row in rows[1:]], default=0),
    )

    while len(headers) < max_cols:
        headers.append(f"Column_{len(headers) + 1}")

    normalized_rows = []
    excel_rows = []
    for idx, row in enumerate(rows[1:], start=2):
        row_copy = list(row)
        while len(row_copy) < max_cols:
            row_copy.append("")
        normalized_rows.append(row_copy[:max_cols])
        excel_rows.append(idx)

    df = pd.DataFrame(normalized_rows, columns=headers)
    df = normalize_dataframe_headers(df)
    df["_excel_row"] = excel_rows
    return df


def ensure_dataframe_columns(df, required_columns):
    df = df.copy()
    for col in required_columns:
        if col not in df.columns:
            df[col] = ""
    return df


# ==============================================================================
# MAINTENANCE & SOW HELPERS
# ==============================================================================

def is_maintenance_sow(sow_type):
    sow = safe_str(sow_type).strip().lower()
    return sow in {
        MAINTENANCE_SOW_EVCS.lower(),
        MAINTENANCE_SOW_BSS.lower(),
    }


def is_maintenance_evcs(sow_type):
    return safe_str(sow_type).strip().lower() == MAINTENANCE_SOW_EVCS.lower()


def is_maintenance_bss(sow_type):
    return safe_str(sow_type).strip().lower() == MAINTENANCE_SOW_BSS.lower()


def get_sow_category(sow_type):
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
    category = get_sow_category(sow_type)
    if category == "SURVEY":
        return 11  # L
    if category == "CONS":
        return 22  # W
    return None


def _looks_like_wo(value):
    text = safe_str(value).strip()
    if not text or is_sheet_error_value(text):
        return False
    upper = text.upper()
    if upper in {"NAN", "NONE", "-"}:
        return False
    return "WO" in upper and "/" in text and len(text) >= 8


def _wo_column_score(series, header, physical_index, expected_index):
    try:
        values = series.tolist()
    except Exception:
        return -1

    valid_count = sum(1 for value in values if _looks_like_wo(value))
    if valid_count == 0:
        return -1

    header_text = normalize_header(header).lower()
    score = valid_count * 100

    if "wo" in header_text:
        score += 1000
        if "cons" in header_text or "construct" in header_text:
            score += 250
        if "survey" in header_text:
            score += 250

    if physical_index == expected_index:
        score += 500

    return score


def get_query_wo_column(df_query, sow_type):
    if df_query is None or df_query.empty:
        return None

    data_cols = [c for c in df_query.columns if c != "_excel_row"]
    category = get_sow_category(sow_type)

    if category == "SURVEY":
        header_candidates = [
            "WO Survey",
            "WO Number Survey",
            "No. WO Survey",
            "No WO Survey",
            "WO Survey Number",
            "Nomor WO Survey",
            "No. WO",
            "No WO",
            "WO Number",
            "Nomor WO",
        ]
        expected_index = 11  # L
    elif category == "CONS":
        header_candidates = [
            "WO Cons",
            "WO Construction",
            "WO Construct",
            "WO Number Cons",
            "No. WO Cons",
            "No WO Cons",
            "WO Cons Number",
            "Nomor WO Cons",
            "No. WO",
            "No WO",
            "WO Number",
            "Nomor WO",
        ]
        expected_index = 22  # W
    else:
        return None

    normalized_headers = {
        normalize_header(col).strip().lower(): idx
        for idx, col in enumerate(data_cols)
    }

    for candidate in header_candidates:
        idx = normalized_headers.get(normalize_header(candidate).strip().lower())
        if idx is not None:
            try:
                if any(_looks_like_wo(value) for value in df_query.iloc[:, idx].tolist()):
                    return idx
            except Exception:
                pass

    if len(data_cols) > expected_index:
        try:
            legacy_series = df_query.iloc[:, expected_index]
            if any(_looks_like_wo(value) for value in legacy_series.tolist()):
                return expected_index
        except Exception:
            pass

    best_idx = None
    best_score = -1

    for idx, col in enumerate(data_cols):
        try:
            score = _wo_column_score(
                df_query.iloc[:, idx],
                col,
                idx,
                expected_index,
            )
        except Exception:
            continue

        if score > best_score:
            best_score = score
            best_idx = idx

    return best_idx


def get_query_data_column(df_query, candidates, legacy_index=None):
    if df_query is None or df_query.empty:
        return None

    detected = find_column(df_query, candidates, fallback=None)
    if detected is not None and detected != "_excel_row":
        return detected

    data_cols = [c for c in df_query.columns if c != "_excel_row"]
    if legacy_index is not None and len(data_cols) > legacy_index:
        return data_cols[legacy_index]

    return None


# ==============================================================================
# BATCH WRITE & PIC
# ==============================================================================

def batch_update_cells(worksheet, updates):
    if not updates:
        return
    worksheet.batch_update(updates, value_input_option="USER_ENTERED")


def column_letter(column_number: int) -> str:
    result = ""
    while column_number > 0:
        column_number, remainder = divmod(column_number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def resolve_pic_name(selected_pic, manual_pic_name=""):
    selected_pic = safe_str(selected_pic)
    manual_pic_name = safe_str(manual_pic_name)
    if selected_pic.upper() == "IN HOUSE":
        return manual_pic_name
    return selected_pic


# ==============================================================================
# SPK NUMBERING & SEQUENCE
# ==============================================================================

def generate_spk_number(sow_type="GENERAL", sequence_num=1):
    now = datetime.datetime.now()
    roman_months = [
        "I", "II", "III", "IV", "V", "VI",
        "VII", "VIII", "IX", "X", "XI", "XII",
    ]
    month_roman = roman_months[now.month - 1]
    sow_lower = safe_str(sow_type).lower()
    sow_code = "SURVEY" if "survey" in sow_lower else "CONS"
    seq_str = f"{sequence_num:04d}"
    return f"{seq_str}/CLX/SPK/{sow_code}/{month_roman}/{now.year}"


def generate_ms_number(sequence_num=1):
    now = datetime.datetime.now()
    roman_months = [
        "I", "II", "III", "IV", "V", "VI",
        "VII", "VIII", "IX", "X", "XI", "XII",
    ]
    month_roman = roman_months[now.month - 1]
    seq_str = f"{sequence_num:04d}"
    return f"{seq_str}/CLX/MS/{month_roman}/{now.year}"


def _next_sequence_from_rows(rows, number_kind):
    now = datetime.datetime.now()
    roman = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
    period = f"/{roman[now.month - 1]}/{now.year}"
    prefix = "/CLX/MS/" if number_kind == "MS" else "/CLX/SPK/"
    max_seq = 0
    for row in (rows or [])[1:]:
        if len(row) < 2:
            continue
        number = safe_str(row[1]).upper()
        if prefix not in number or not number.endswith(period):
            continue
        match = re.match(r"^(\d+)/CLX/", number)
        if match:
            max_seq = max(max_seq, int(match.group(1)))
    return max_seq + 1


def get_next_spk_sequence_from_rows(rows):
    return _next_sequence_from_rows(rows, "PROJECT")


def get_next_ms_sequence_from_rows(rows):
    return _next_sequence_from_rows(rows, "MS")


# ==============================================================================
# LOAD MASTER & ENSURE SHEETS
# ==============================================================================

def load_master_dropdown():
    rows = load_sheet_values_cached(SHEET_DROPDOWN)
    if not rows or len(rows) <= 1:
        return pd.DataFrame()
    return dataframe_from_sheet_rows(rows)


def load_master_sow():
    rows = load_sheet_values_cached(SHEET_SOW)
    if not rows or len(rows) <= 1:
        return pd.DataFrame()
    return dataframe_from_sheet_rows(rows)


def ensure_spk_sheet(sh, sheet_name):
    try:
        return sh.worksheet(sheet_name)
    except Exception:
        target_sheet = sh.add_worksheet(title=sheet_name, rows=1000, cols=20)
        target_sheet.update("A1:N1", [DEFAULT_SPK_HEADERS], value_input_option="USER_ENTERED")
        clear_sheet_cache()
        return target_sheet


def ensure_ms_sheet(sh):
    try:
        return sh.worksheet(SHEET_MS)
    except Exception:
        target_sheet = sh.add_worksheet(title=SHEET_MS, rows=1000, cols=7)
        target_sheet.update("A1:G1", [DEFAULT_MS_HEADERS], value_input_option="USER_ENTERED")
        clear_sheet_cache()
        return target_sheet


def ensure_status_column(sheet, df):
    df = normalize_dataframe_headers(df.copy())
    status_col = find_column(
        df,
        ["Status Approval", "Approval Status", "Status"],
        fallback=None,
    )

    if status_col and status_col != "_excel_row":
        return df, status_col

    status_col = "Status Approval"
    try:
        headers = [normalize_header(h) for h in sheet.row_values(1)]
        if status_col not in headers:
            headers.append(status_col)
            header_end = column_letter(len(headers))
            sheet.update(f"A1:{header_end}1", [headers], value_input_option="USER_ENTERED")

            if len(df) > 0:
                status_col_index = len(headers)
                status_letter = column_letter(status_col_index)
                status_values = [[STATUS_PENDING] for _ in range(len(df))]
                sheet.update(
                    f"{status_letter}2:{status_letter}{len(df) + 1}",
                    status_values,
                    value_input_option="USER_ENTERED",
                )
            clear_sheet_cache()
    except Exception as e:
        st.warning(f"⚠️ Kolom Status Approval belum dapat ditambahkan ke Google Sheet: {e}")

    df[status_col] = STATUS_PENDING
    return df, status_col


# ==============================================================================
# MODERN CORPORATE PDF ENGINE (A4, HIGHLIGHTED SECTIONS, 3-SIGNER BLOCK)
# ==============================================================================

COLOR_NAVY = colors.HexColor("#1B365D")
COLOR_BLUE_ACCENT = colors.HexColor("#0284C7")
COLOR_HIGHLIGHT_BG = colors.HexColor("#EFF6FF")
COLOR_SECTION_BG = colors.HexColor("#DCE7F5")
COLOR_CARD_BG = colors.HexColor("#F8FAFC")
COLOR_ZEBRA_BG = colors.HexColor("#F8FAFC")
COLOR_BORDER_LIGHT = colors.HexColor("#CBD5E1")
COLOR_TEXT_DARK = colors.HexColor("#0F172A")
COLOR_TEXT_MUTED = colors.HexColor("#475569")

COLOR_APPROVED_BG = colors.HexColor("#DCFCE7")
COLOR_APPROVED_TEXT = colors.HexColor("#166534")
COLOR_APPROVED_BORDER = colors.HexColor("#16A34A")

COLOR_DRAFT_BG = colors.HexColor("#FEF3C7")
COLOR_DRAFT_TEXT = colors.HexColor("#92400E")
COLOR_DRAFT_BORDER = colors.HexColor("#D97706")


class _SpkNumberedCanvas(pdf_canvas.Canvas):
    def __init__(self, *args, **kwargs):
        self.doc_ref = kwargs.pop("doc_ref", "-")
        self.approved = kwargs.pop("approved", False)
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
        page_w, _ = A4
        margin_x = 0.42 * inch
        y_line = 28
        y_text = 16

        self.saveState()
        self.setStrokeColor(COLOR_BORDER_LIGHT)
        self.setLineWidth(0.8)
        self.line(margin_x, y_line, page_w - margin_x, y_line)

        self.setFont("Helvetica", 7.5)
        self.setFillColor(COLOR_TEXT_MUTED)
        status_label = "APPROVED BY CSO" if self.approved else "DRAFT / PENDING CSO APPROVAL"
        left_text = (
            f"PT. Connectivity Leads eXcellence  |  SPK Ref: {self.doc_ref}  |  Status: {status_label}"
        )
        self.drawString(margin_x, y_text, left_text)
        self.drawRightString(
            page_w - margin_x,
            y_text,
            f"Halaman {self._pageNumber} dari {total_pages}",
        )
        self.restoreState()


def get_pdf_styles():
    styles = getSampleStyleSheet()

    company_title = ParagraphStyle(
        "SpkCompanyTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=12.5,
        leading=15,
        textColor=COLOR_NAVY,
    )

    company_info = ParagraphStyle(
        "SpkCompanyInfo",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11.5,
        textColor=COLOR_TEXT_MUTED,
    )

    doc_title = ParagraphStyle(
        "SpkDocTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=14.5,
        leading=18,
        textColor=COLOR_NAVY,
        spaceAfter=2,
    )

    doc_subtitle = ParagraphStyle(
        "SpkDocSubtitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=12.5,
        textColor=COLOR_TEXT_DARK,
    )

    badge_approved = ParagraphStyle(
        "SpkBadgeApproved",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=11,
        alignment=1,
        textColor=COLOR_APPROVED_TEXT,
    )

    badge_draft = ParagraphStyle(
        "SpkBadgeDraft",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=11,
        alignment=1,
        textColor=COLOR_DRAFT_TEXT,
    )

    section_banner = ParagraphStyle(
        "SpkSectionBanner",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9.5,
        leading=12,
        textColor=COLOR_NAVY,
    )

    meta_label = ParagraphStyle(
        "SpkMetaLabel",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.8,
        leading=11.5,
        textColor=COLOR_TEXT_MUTED,
    )

    meta_val = ParagraphStyle(
        "SpkMetaVal",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.8,
        leading=11.5,
        textColor=COLOR_TEXT_DARK,
    )

    cell_head = ParagraphStyle(
        "SpkCellHead",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=11,
        alignment=1,
        textColor=colors.white,
    )

    cell_body = ParagraphStyle(
        "SpkCellBody",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11.5,
        textColor=COLOR_TEXT_DARK,
    )

    cell_center = ParagraphStyle(
        "SpkCellCenter",
        parent=cell_body,
        alignment=1,
    )

    return {
        "company_title": company_title,
        "company_info": company_info,
        "doc_title": doc_title,
        "doc_subtitle": doc_subtitle,
        "badge_approved": badge_approved,
        "badge_draft": badge_draft,
        "section_banner": section_banner,
        "meta_label": meta_label,
        "meta_val": meta_val,
        "cell_head": cell_head,
        "cell_body": cell_body,
        "cell_center": cell_center,
    }


def build_pdf_header(elements, st_styles):
    logo_path = get_company_logo_path()
    logo_cell = _build_proportional_image(
        logo_path,
        max_width=1.55 * inch,
        max_height=0.68 * inch,
        h_align="LEFT",
    )

    if logo_cell is None:
        logo_cell = Paragraph(
            "<b>PT. CLX</b>",
            ParagraphStyle(
                "LogoFallback",
                fontName="Helvetica-Bold",
                fontSize=16,
                textColor=COLOR_NAVY,
            ),
        )

    company_block = [
        Paragraph("PT. Connectivity Leads eXcellence", st_styles["company_title"]),
        Spacer(1, 2),
        Paragraph(
            "<b>Jakarta Office:</b> Jl. M. Ali 2 No. 19 RT 007 RW 004, "
            "Tanah Baru, Beji, Kota Depok, Jawa Barat<br/>"
            "<b>Email:</b> clx.central@gmail.com  |  "
            "<b>Services:</b> EV Charging &amp; Battery Swapping Infrastructure",
            st_styles["company_info"],
        ),
    ]

    t_header = Table(
        [[logo_cell, company_block]],
        colWidths=[134, 400],
    )
    t_header.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("LINEBELOW", (0, 0), (-1, -1), 2.2, COLOR_NAVY),
            ]
        )
    )

    elements.append(t_header)
    elements.append(Spacer(1, 10))


def _build_title_and_status_banner(elements, st_styles, no_spk, date_str, category_tag, approved=False):
    left_box = [
        Paragraph("SURAT PERINTAH KERJA (SPK)", st_styles["doc_title"]),
        Paragraph(
            f"<b>No. Dokumen:</b> <font color='#0284C7'><b>{esc_pdf(no_spk)}</b></font>"
            f"  &nbsp;|&nbsp;  <b>Tanggal:</b> Jakarta, {esc_pdf(date_str)}",
            st_styles["doc_subtitle"],
        ),
    ]

    if approved:
        badge_text = "STATUS DOKUMEN<br/>APPROVED BY CSO"
        badge_style = st_styles["badge_approved"]
        badge_bg = COLOR_APPROVED_BG
        badge_border = COLOR_APPROVED_BORDER
    else:
        badge_text = "STATUS DOKUMEN<br/>PENDING CSO APPROVAL"
        badge_style = st_styles["badge_draft"]
        badge_bg = COLOR_DRAFT_BG
        badge_border = COLOR_DRAFT_BORDER

    right_box = Table(
        [
            [Paragraph(esc_pdf(category_tag), st_styles["badge_approved"] if approved else ParagraphStyle(
                "CatTag",
                fontName="Helvetica-Bold",
                fontSize=8,
                leading=10,
                alignment=1,
                textColor=COLOR_NAVY,
            ))],
            [Paragraph(badge_text, badge_style)],
        ],
        colWidths=[152],
    )
    right_box.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, 0), COLOR_SECTION_BG),
                ("BACKGROUND", (0, 1), (0, 1), badge_bg),
                ("BOX", (0, 0), (-1, -1), 0.9, badge_border),
                ("LINEBELOW", (0, 0), (0, 0), 0.6, COLOR_BORDER_LIGHT),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )

    banner = Table([[left_box, right_box]], colWidths=[382, 152])
    banner.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, 0), COLOR_HIGHLIGHT_BG),
                ("LINEBEFORE", (0, 0), (0, 0), 3.5, COLOR_BLUE_ACCENT),
                ("BOX", (0, 0), (0, 0), 0.6, COLOR_BORDER_LIGHT),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (0, 0), 7),
                ("BOTTOMPADDING", (0, 0), (0, 0), 7),
                ("LEFTPADDING", (0, 0), (0, 0), 10),
                ("RIGHTPADDING", (0, 0), (0, 0), 8),
                ("LEFTPADDING", (1, 0), (1, 0), 6),
                ("RIGHTPADDING", (1, 0), (1, 0), 0),
                ("TOPPADDING", (1, 0), (1, 0), 0),
                ("BOTTOMPADDING", (1, 0), (1, 0), 0),
            ]
        )
    )

    elements.append(banner)
    elements.append(Spacer(1, 10))


def _build_section_highlight_header(title_text, st_styles):
    tbl = Table(
        [[Paragraph(esc_pdf(title_text), st_styles["section_banner"])]],
        colWidths=[534],
    )
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_SECTION_BG),
                ("LINEBEFORE", (0, 0), (0, -1), 3.5, COLOR_NAVY),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    return tbl


def _build_three_party_signature_block(
    pelaksana_name,
    pelaksana_sub,
    approved=False,
):
    sign_header = ParagraphStyle(
        "SignHeader3Col",
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=11,
        alignment=1,
        textColor=COLOR_NAVY,
    )
    sign_sub_header = ParagraphStyle(
        "SignSubHeader3Col",
        fontName="Helvetica",
        fontSize=7.8,
        leading=10,
        alignment=1,
        textColor=COLOR_TEXT_MUTED,
    )
    sign_name = ParagraphStyle(
        "SignName3Col",
        fontName="Helvetica-Bold",
        fontSize=8.8,
        leading=11.5,
        alignment=1,
        textColor=COLOR_TEXT_DARK,
    )
    sign_role = ParagraphStyle(
        "SignRole3Col",
        fontName="Helvetica",
        fontSize=8.0,
        leading=10.5,
        alignment=1,
        textColor=COLOR_TEXT_MUTED,
    )
    sign_placeholder = ParagraphStyle(
        "SignPlaceholder",
        fontName="Helvetica-Oblique",
        fontSize=7.5,
        leading=10,
        alignment=1,
        textColor=COLOR_TEXT_MUTED,
    )

    vm_sig_img = build_vm_signature_image(max_width=1.45 * inch, max_height=0.62 * inch)
    vm_sig_cell = (
        vm_sig_img
        if vm_sig_img is not None
        else Paragraph("(Verified by Vendor Management)", sign_placeholder)
    )

    cso_sig_img = (
        build_cso_signature_image(max_width=1.45 * inch, max_height=0.62 * inch)
        if approved
        else None
    )
    cso_sig_cell = (
        cso_sig_img
        if cso_sig_img is not None
        else Paragraph(
            "(Menunggu Approval CSO)" if not approved else "(Approved by CSO)",
            sign_placeholder,
        )
    )

    sign_data = [
        [
            Paragraph("PELAKSANA PEKERJAAN", sign_header),
            Paragraph("DIPERIKSA OLEH", sign_header),
            Paragraph("DISETUJUI OLEH", sign_header),
        ],
        [
            Paragraph("Mitra / Kontraktor", sign_sub_header),
            Paragraph(esc_pdf(VM_SIGNER_TITLE), sign_sub_header),
            Paragraph(esc_pdf(CSO_SIGNER_TITLE), sign_sub_header),
        ],
        [
            Spacer(1, 46),
            vm_sig_cell,
            cso_sig_cell,
        ],
        [
            Paragraph(
                f"<u><b>{esc_pdf(pelaksana_name)}</b></u>",
                sign_name,
            ),
            Paragraph(
                f"<u><b>{esc_pdf(VM_SIGNER_NAME)}</b></u>",
                sign_name,
            ),
            Paragraph(
                f"<u><b>{esc_pdf(CSO_SIGNER_NAME)}</b></u>",
                sign_name,
            ),
        ],
        [
            Paragraph(esc_pdf(pelaksana_sub, "Pelaksana Lapangan"), sign_role),
            Paragraph("Vendor Management", sign_role),
            Paragraph("Chief Services Officer (CSO)", sign_role),
        ],
    ]

    t_sign = Table(sign_data, colWidths=[178, 178, 178])
    t_sign.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), COLOR_CARD_BG),
                ("BOX", (0, 0), (0, -1), 0.5, COLOR_BORDER_LIGHT),
                ("BOX", (1, 0), (1, -1), 0.5, COLOR_BORDER_LIGHT),
                ("BOX", (2, 0), (2, -1), 0.5, COLOR_BORDER_LIGHT),
                ("LINEBELOW", (0, 1), (-1, 1), 0.5, COLOR_BORDER_LIGHT),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, 1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, 1), 3),
                ("TOPPADDING", (0, 2), (-1, 2), 6),
                ("BOTTOMPADDING", (0, 2), (-1, 2), 4),
                ("TOPPADDING", (0, 3), (-1, 4), 2),
                ("BOTTOMPADDING", (0, 3), (-1, 4), 4),
            ]
        )
    )
    return t_sign


# ==============================================================================
# PDF - PROJECT
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
        pagesize=portrait(A4),
        leftMargin=0.42 * inch,
        rightMargin=0.42 * inch,
        topMargin=0.40 * inch,
        bottomMargin=0.55 * inch,
    )

    st_styles = get_pdf_styles()
    elements = []

    build_pdf_header(elements, st_styles)

    no_spk = safe_str(spk_metadata.get("no_spk"), "-")
    date_str = (
        safe_str(spk_metadata.get("date_spk"))
        or datetime.datetime.now().strftime("%d/%m/%Y")
    )
    sow_type_label = safe_str(spk_metadata.get("sow_type"), "PROJECT")

    _build_title_and_status_banner(
        elements,
        st_styles,
        no_spk=no_spk,
        date_str=date_str,
        category_tag=f"SPK PROJECT • {sow_type_label.upper()}",
        approved=approved,
    )

    pic_name = safe_str(spk_metadata.get("pic_name"), "-")
    pic_phone = safe_str(spk_metadata.get("pic_phone"), "-")
    display_wo = safe_str(selected_wo, "-")

    meta_table_data = [
        [
            Paragraph("Proyek", st_styles["meta_label"]),
            Paragraph(f": <b>{esc_pdf(spk_metadata.get('proyek'), '-')}</b>", st_styles["meta_val"]),
            Paragraph("No. Work Order", st_styles["meta_label"]),
            Paragraph(f": <b>{esc_pdf(display_wo, '-')}</b>", st_styles["meta_val"]),
        ],
        [
            Paragraph("Pekerjaan", st_styles["meta_label"]),
            Paragraph(f": {esc_pdf(spk_metadata.get('pekerjaan'), '-')}", st_styles["meta_val"]),
            Paragraph("Mitra / PIC", st_styles["meta_label"]),
            Paragraph(f": <b>{esc_pdf(pic_name)}</b> ({esc_pdf(pic_phone)})", st_styles["meta_val"]),
        ],
        [
            Paragraph("Wilayah / Lokasi", st_styles["meta_label"]),
            Paragraph(f": {esc_pdf(spk_metadata.get('lokasi'), '-')}", st_styles["meta_val"]),
            Paragraph("Jenis SOW", st_styles["meta_label"]),
            Paragraph(f": {esc_pdf(sow_type_label, '-')}", st_styles["meta_val"]),
        ],
        [
            Paragraph("Jumlah Site", st_styles["meta_label"]),
            Paragraph(
                f": <b>{len(selected_sites) if selected_sites is not None else 0} Site</b> (Terlampir di bawah)",
                st_styles["meta_val"],
            ),
            Paragraph("Kontak CLX", st_styles["meta_label"]),
            Paragraph(
                ": Wikantiyoso S. (0878-8855-0300)<br/>&nbsp;&nbsp;Hartono (0818-0690-9317)",
                st_styles["meta_val"],
            ),
        ],
    ]

    t_meta = Table(meta_table_data, colWidths=[88, 179, 98, 169])
    t_meta.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 4.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_CARD_BG),
                ("LINEBEFORE", (0, 0), (0, -1), 3.5, COLOR_BLUE_ACCENT),
                ("BOX", (0, 0), (-1, -1), 0.6, COLOR_BORDER_LIGHT),
                ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#E2E8F0")),
            ]
        )
    )

    elements.append(t_meta)
    elements.append(Spacer(1, 12))

    # --------------------------------------------------------------------------
    # SECTION 1: SCOPE OF WORK (SOW)
    # --------------------------------------------------------------------------
    elements.append(_build_section_highlight_header("1. RUANG LINGKUP PEKERJAAN (SCOPE OF WORK)", st_styles))
    elements.append(Spacer(1, 4))

    sow_headers = [
        Paragraph("NO", st_styles["cell_head"]),
        Paragraph("URAIAN PEKERJAAN (SOW)", st_styles["cell_head"]),
        Paragraph("TARGET PENYELESAIAN", st_styles["cell_head"]),
    ]
    sow_rows = [sow_headers]

    if matched_sow_df is not None and not matched_sow_df.empty:
        for idx, (_, row) in enumerate(matched_sow_df.iterrows(), 1):
            deskripsi = safe_str(row.iloc[2], "-") if row.shape[0] > 2 else "-"
            target = safe_str(row.iloc[3], "1 Hari") if row.shape[0] > 3 else "1 Hari"
            sow_rows.append(
                [
                    Paragraph(str(idx), st_styles["cell_center"]),
                    Paragraph(esc_pdf(deskripsi), st_styles["cell_body"]),
                    Paragraph(f"<b>{esc_pdf(target)}</b>", st_styles["cell_center"]),
                ]
            )
    else:
        sow_rows.append(
            [
                Paragraph("1", st_styles["cell_center"]),
                Paragraph(
                    f"Pelaksanaan pekerjaan <b>{esc_pdf(sow_type_label)}</b> sesuai standar teknis dan K3 PT. Connectivity Leads eXcellence.",
                    st_styles["cell_body"],
                ),
                Paragraph("<b>1 Hari</b>", st_styles["cell_center"]),
            ]
        )

    sow_table_style = [
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_NAVY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.45, COLOR_BORDER_LIGHT),
        ("BOX", (0, 0), (-1, -1), 0.8, COLOR_NAVY),
        ("TOPPADDING", (0, 0), (-1, 0), 6),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("TOPPADDING", (0, 1), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]
    for r_i in range(1, len(sow_rows)):
        if r_i % 2 == 0:
            sow_table_style.append(("BACKGROUND", (0, r_i), (-1, r_i), COLOR_ZEBRA_BG))

    t_sow = Table(sow_rows, colWidths=[34, 385, 115], repeatRows=1)
    t_sow.setStyle(TableStyle(sow_table_style))

    elements.append(t_sow)
    elements.append(Spacer(1, 12))

    # --------------------------------------------------------------------------
    # SECTION 2: DAFTAR LOKASI / SITE LIST
    # --------------------------------------------------------------------------
    if selected_sites is not None and not selected_sites.empty:
        elements.append(_build_section_highlight_header("2. DAFTAR LOKASI PEKERJAAN (SITE LIST)", st_styles))
        elements.append(Spacer(1, 4))

        site_headers = [
            Paragraph("NO", st_styles["cell_head"]),
            Paragraph("SITE NAME", st_styles["cell_head"]),
            Paragraph("CHARGER", st_styles["cell_head"]),
            Paragraph("WO NUMBER", st_styles["cell_head"]),
            Paragraph("PROVINCE", st_styles["cell_head"]),
            Paragraph("PIC + CONTACT", st_styles["cell_head"]),
            Paragraph("MAPS", st_styles["cell_head"]),
        ]
        site_rows = [site_headers]

        for idx, (_, row) in enumerate(selected_sites.iterrows(), 1):
            gmaps_raw = safe_str(row.get("col_gmaps", "-"), "-")
            if gmaps_raw.lower().startswith(("http://", "https://")):
                safe_url = xml_escape(gmaps_raw, {'"': "&quot;"})
                gmaps_cell = Paragraph(
                    f'<a href="{safe_url}" color="#0284C7"><u><b>Buka Peta</b></u></a>',
                    st_styles["cell_center"],
                )
            else:
                gmaps_cell = Paragraph(esc_pdf(gmaps_raw), st_styles["cell_center"])

            site_rows.append(
                [
                    Paragraph(str(idx), st_styles["cell_center"]),
                    Paragraph(f"<b>{esc_pdf(row.get('col_site', '-'))}</b>", st_styles["cell_body"]),
                    Paragraph(esc_pdf(row.get("col_charge", "-")), st_styles["cell_center"]),
                    Paragraph(esc_pdf(display_wo), st_styles["cell_center"]),
                    Paragraph(esc_pdf(row.get("col_province", "-")), st_styles["cell_body"]),
                    Paragraph(esc_pdf(row.get("col_pic", "-")), st_styles["cell_body"]),
                    gmaps_cell,
                ]
            )

        site_table_style = [
            ("BACKGROUND", (0, 0), (-1, 0), COLOR_NAVY),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.45, COLOR_BORDER_LIGHT),
            ("BOX", (0, 0), (-1, -1), 0.8, COLOR_NAVY),
            ("TOPPADDING", (0, 0), (-1, 0), 6),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
            ("TOPPADDING", (0, 1), (-1, -1), 4.5),
            ("BOTTOMPADDING", (0, 1), (-1, -1), 4.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ]
        for r_i in range(1, len(site_rows)):
            if r_i % 2 == 0:
                site_table_style.append(("BACKGROUND", (0, r_i), (-1, r_i), COLOR_ZEBRA_BG))

        t_site = Table(
            site_rows,
            colWidths=[26, 134, 68, 96, 72, 86, 52],
            repeatRows=1,
        )
        t_site.setStyle(TableStyle(site_table_style))

        elements.append(t_site)
        elements.append(Spacer(1, 14))

    # --------------------------------------------------------------------------
    # SECTION 3: KETENTUAN & PENGESAHAN (3 TANDA TANGAN)
    # --------------------------------------------------------------------------
    note_box = Table(
        [
            [
                Paragraph(
                    "<b>Catatan Pelaksanaan:</b> Pelaksana wajib mematuhi standar keselamatan kerja (K3), "
                    "melampirkan dokumentasi foto progres &amp; BAST setelah pekerjaan selesai, serta "
                    "berkoordinasi dengan tim Vendor Management &amp; Project PT. Connectivity Leads eXcellence.",
                    ParagraphStyle(
                        "NoteText",
                        fontName="Helvetica",
                        fontSize=8.2,
                        leading=11,
                        textColor=COLOR_TEXT_DARK,
                    ),
                )
            ]
        ],
        colWidths=[534],
    )
    note_box.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_HIGHLIGHT_BG),
                ("LINEBEFORE", (0, 0), (0, -1), 3.0, COLOR_BLUE_ACCENT),
                ("BOX", (0, 0), (-1, -1), 0.5, COLOR_BORDER_LIGHT),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )

    t_sign = _build_three_party_signature_block(
        pelaksana_name=pic_name,
        pelaksana_sub=f"Telp: {pic_phone}",
        approved=approved,
    )

    elements.append(
        KeepTogether(
            [
                note_box,
                Spacer(1, 12),
                t_sign,
            ]
        )
    )

    doc.build(
        elements,
        canvasmaker=lambda *args, **kwargs: _SpkNumberedCanvas(
            *args,
            doc_ref=no_spk,
            approved=approved,
            **kwargs,
        ),
    )

    buffer.seek(0)
    return no_spk, buffer.getvalue()


# ==============================================================================
# PDF - OPERATION (MAINTENANCE SERVICE EVCS & BSS)
# ==============================================================================

def _build_ms_scope_paragraph(pekerjaan, sow_type, cell_style):
    combined_label = f"{safe_str(sow_type)} {safe_str(pekerjaan)}".lower()

    if "evcs" in combined_label:
        scope_html = (
            f"<b>{esc_pdf(pekerjaan)} (AC &amp; DC Charger)</b><br/>"
            "<b>1. Preventive Maintenance (PM):</b> Pembersihan unit, <i>charging gun</i>, layar HMI, "
            "dan filter udara kabinet; pengecekan panel listrik (MDP/SDP), <i>grounding</i>, MCB/SPD, "
            "serta uji fungsi <i>Emergency Stop</i> dan koneksi OCPP/jaringan.<br/>"
            "<b>2. Corrective Maintenance (CM):</b> Penanganan tiket gangguan lapangan "
            "(<i>troubleshooting</i>), <i>reset</i> sistem, serta penggantian <i>sparepart</i> rusak "
            "sesuai persetujuan tim teknis &amp; SLA.<br/>"
            "<b>3. Inspeksi Area, K3 &amp; Pelaporan:</b> Pengecekan kondisi sipil (<i>pedestal</i>, "
            "marka, penerangan), penerapan standar K3 kelistrikan, serta penyampaian <i>checklist</i> "
            "digital, foto dokumentasi (<i>Before-After</i>), dan laporan SLA bulanan."
        )
    elif "bss" in combined_label:
        scope_html = (
            f"<b>{esc_pdf(pekerjaan)} (Battery Swapping Station)</b><br/>"
            "<b>1. Preventive Maintenance (PM):</b> Pembersihan kabinet dan slot baterai; pengecekan "
            "mekanisme <i>door lock</i>, pin konektor daya &amp; komunikasi, modul <i>charger</i> per slot, "
            "sistem <i>grounding</i>, kipas pendingin, sensor suhu/asap, dan sinkronisasi IoT.<br/>"
            "<b>2. Corrective Maintenance (CM):</b> Penanganan gangguan operasional slot/kabinet "
            "(<i>troubleshooting</i>), <i>reset</i> sistem, serta penggantian komponen/modul sesuai "
            "persetujuan tim teknis &amp; SLA.<br/>"
            "<b>3. Inspeksi Area, K3 &amp; Pelaporan:</b> Pengecekan kondisi fisik area sekitar site, "
            "penerapan standar K3 kelistrikan, serta penyampaian <i>checklist</i> digital, foto dokumentasi "
            "(<i>Before-After</i>), dan laporan SLA bulanan."
        )
    else:
        scope_html = (
            f"<b>{esc_pdf(pekerjaan)}</b><br/>"
            "Melaksanakan kegiatan Preventive Maintenance (PM), Corrective Maintenance (CM), "
            "pengecekan berkala perangkat, serta pelaporan SLA harian/bulanan sesuai ketentuan "
            "PT. Connectivity Leads eXcellence."
        )

    return Paragraph(scope_html, cell_style)


def generate_ms_pdf_bytes(spk_metadata, approved=False):
    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=portrait(A4),
        leftMargin=0.42 * inch,
        rightMargin=0.42 * inch,
        topMargin=0.40 * inch,
        bottomMargin=0.55 * inch,
    )

    st_styles = get_pdf_styles()
    elements = []

    build_pdf_header(elements, st_styles)

    no_spk = safe_str(spk_metadata.get("no_spk"), "-")
    pekerjaan = safe_str(spk_metadata.get("pekerjaan"), "-")
    mitra = safe_str(spk_metadata.get("mitra"), "-")
    periode_start = safe_str(spk_metadata.get("periode_start"), "-")
    periode_end = safe_str(spk_metadata.get("periode_end"), "-")
    sow_type = safe_str(spk_metadata.get("sow_type"), "Maintenance Service")
    date_str = safe_str(
        spk_metadata.get("date_spk"),
        datetime.datetime.now().strftime("%d/%m/%Y"),
    )

    _build_title_and_status_banner(
        elements,
        st_styles,
        no_spk=no_spk,
        date_str=date_str,
        category_tag="SPK OPERATION • MAINTENANCE",
        approved=approved,
    )

    meta_table_data = [
        [
            Paragraph("Kategori SPK", st_styles["meta_label"]),
            Paragraph(": <b>Operation / Managed Services</b>", st_styles["meta_val"]),
            Paragraph("Jenis SOW", st_styles["meta_label"]),
            Paragraph(f": <b>{esc_pdf(sow_type)}</b>", st_styles["meta_val"]),
        ],
        [
            Paragraph("Pekerjaan", st_styles["meta_label"]),
            Paragraph(f": {esc_pdf(pekerjaan)}", st_styles["meta_val"]),
            Paragraph("Mitra Pelaksana", st_styles["meta_label"]),
            Paragraph(f": <b>{esc_pdf(mitra)}</b>", st_styles["meta_val"]),
        ],
        [
            Paragraph("Periode Mulai", st_styles["meta_label"]),
            Paragraph(f": <b>{esc_pdf(periode_start)}</b>", st_styles["meta_val"]),
            Paragraph("Periode Selesai", st_styles["meta_label"]),
            Paragraph(f": <b>{esc_pdf(periode_end)}</b>", st_styles["meta_val"]),
        ],
        [
            Paragraph("Penerbit SPK", st_styles["meta_label"]),
            Paragraph(": PT. Connectivity Leads eXcellence", st_styles["meta_val"]),
            Paragraph("Tanggal Terbit", st_styles["meta_label"]),
            Paragraph(f": {esc_pdf(date_str)}", st_styles["meta_val"]),
        ],
    ]

    t_meta = Table(meta_table_data, colWidths=[88, 179, 98, 169])
    t_meta.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_CARD_BG),
                ("LINEBEFORE", (0, 0), (0, -1), 3.5, COLOR_BLUE_ACCENT),
                ("BOX", (0, 0), (-1, -1), 0.6, COLOR_BORDER_LIGHT),
                ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#E2E8F0")),
            ]
        )
    )

    elements.append(t_meta)
    elements.append(Spacer(1, 14))

    elements.append(_build_section_highlight_header("1. DETAIL RUANG LINGKUP PEKERJAAN OPERATION", st_styles))
    elements.append(Spacer(1, 5))

    operation_table = Table(
        [
            [
                Paragraph("NO", st_styles["cell_head"]),
                Paragraph("URAIAN PEKERJAAN &amp; RUANG LINGKUP", st_styles["cell_head"]),
                Paragraph("PERIODE PELAKSANAAN", st_styles["cell_head"]),
            ],
            [
                Paragraph("1", st_styles["cell_center"]),
                _build_ms_scope_paragraph(pekerjaan, sow_type, st_styles["cell_body"]),
                Paragraph(
                    f"<b>{esc_pdf(periode_start)}</b><br/>s/d<br/><b>{esc_pdf(periode_end)}</b>",
                    st_styles["cell_center"],
                ),
            ],
        ],
        colWidths=[36, 358, 140],
        repeatRows=1,
    )

    operation_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), COLOR_NAVY),
                ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDER_LIGHT),
                ("BOX", (0, 0), (-1, -1), 0.8, COLOR_NAVY),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, 0), 6),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                ("TOPPADDING", (0, 1), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 1), (-1, -1), 8),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )

    elements.append(operation_table)
    elements.append(Spacer(1, 16))

    note_box = Table(
        [
            [
                Paragraph(
                    "<b>Ketentuan Operation:</b> Mitra wajib menjaga ketersediaan teknisi sesuai SLA yang berlaku, "
                    "mengisi log monitoring &amp; laporan corrective maintenance secara tepat waktu, serta "
                    "berkoordinasi dengan tim Vendor Management &amp; Operation PT. Connectivity Leads eXcellence.",
                    ParagraphStyle(
                        "MsNoteText",
                        fontName="Helvetica",
                        fontSize=8.3,
                        leading=11.5,
                        textColor=COLOR_TEXT_DARK,
                    ),
                )
            ]
        ],
        colWidths=[534],
    )
    note_box.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_HIGHLIGHT_BG),
                ("LINEBEFORE", (0, 0), (0, -1), 3.0, COLOR_BLUE_ACCENT),
                ("BOX", (0, 0), (-1, -1), 0.5, COLOR_BORDER_LIGHT),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )

    t_sign = _build_three_party_signature_block(
        pelaksana_name=mitra,
        pelaksana_sub="Mitra Operation",
        approved=approved,
    )

    elements.append(
        KeepTogether(
            [
                note_box,
                Spacer(1, 14),
                t_sign,
            ]
        )
    )

    doc.build(
        elements,
        canvasmaker=lambda *args, **kwargs: _SpkNumberedCanvas(
            *args,
            doc_ref=no_spk,
            approved=approved,
            **kwargs,
        ),
    )

    buffer.seek(0)
    return no_spk, buffer.getvalue()


# ==============================================================================
# RECREATE PROJECT PDF FROM DATABASE (WITH QUERY & DROPDOWN ENRICHMENT)
# ==============================================================================

def _enrich_site_metadata_from_query(site_names):
    """Ambil Province, PIC + Contact, dan Gmaps dari sheet Query saat re-download PDF."""
    lookup = {}
    try:
        q_rows = load_sheet_values_cached(SHEET_QUERY)
        if not q_rows or len(q_rows) <= 1:
            return lookup
        df_q = dataframe_from_sheet_rows(q_rows)
        q_site_col = get_query_data_column(
            df_q,
            ["Site Name", "Site", "Nama Site", "Site ID / Site Name"],
            legacy_index=5,
        )
        q_prov_col = get_query_data_column(
            df_q,
            ["Province", "Provinsi", "Province Name"],
            legacy_index=8,
        )
        q_pic_col = get_query_data_column(
            df_q,
            ["PIC + Contact", "PIC", "PIC Name", "PIC Contact", "Contact Person", "Penanggung Jawab"],
            legacy_index=10,
        )
        q_gmaps_col = get_query_data_column(
            df_q,
            ["Gmaps", "Google Maps", "Google Map", "Maps", "Map"],
            legacy_index=7,
        )

        if not q_site_col:
            return lookup

        target_set = {safe_str(s).lower() for s in site_names if safe_str(s)}
        for _, r in df_q.iterrows():
            s_name = safe_str(r.get(q_site_col, ""))
            key = s_name.lower()
            if key in target_set and key not in lookup:
                lookup[key] = {
                    "province": safe_str(r.get(q_prov_col, "-"), "-") if q_prov_col else "-",
                    "pic": safe_str(r.get(q_pic_col, "-"), "-") if q_pic_col else "-",
                    "gmaps": safe_str(r.get(q_gmaps_col, "-"), "-") if q_gmaps_col else "-",
                }
    except Exception:
        pass
    return lookup


def _lookup_mitra_phone(mitra_name):
    target = safe_str(mitra_name).lower()
    if not target or target == "-":
        return "-"
    try:
        df_dd = load_master_dropdown()
        if not df_dd.empty and df_dd.shape[1] >= 2:
            m_col = df_dd.columns[0]
            p_col = df_dd.columns[1]
            hit = df_dd[df_dd[m_col].astype(str).str.strip().str.lower() == target]
            if not hit.empty:
                return safe_str(hit.iloc[0][p_col], "-")
    except Exception:
        pass
    return "-"


def generate_project_pdf_from_database(
    df_project,
    selected_spk,
    approved=False,
):
    if df_project is None or df_project.empty:
        return None, None

    col_no_spk = find_column(
        df_project,
        ["No. SPK", "No SPK", "SPK Number"],
        fallback="No. SPK",
    )

    matched = df_project[
        df_project[col_no_spk].astype(str).str.strip() == safe_str(selected_spk)
    ].copy()

    if matched.empty:
        return None, None

    first_row = matched.iloc[0]

    col_date = find_column(matched, ["Date SPK", "Date"], fallback=None)
    col_wo = find_column(matched, ["No. WO", "No WO", "WO Number"], fallback=None)
    col_pekerjaan = find_column(matched, ["Pekerjaan", "Job"], fallback=None)
    col_mitra = find_column(matched, ["Mitra", "Mitra Name"], fallback=None)
    col_site = find_column(matched, ["Site Name", "Site", "Nama Site"], fallback=None)
    col_charge = find_column(matched, ["Charger Type", "Charging Type"], fallback=None)
    col_sow = find_column(matched, ["SOW", "WO Release"], fallback=None)

    selected_sites = pd.DataFrame()
    provinces_collected = []

    if col_site:
        site_list = matched[col_site].tolist()
        query_lookup = _enrich_site_metadata_from_query(site_list)

        prov_list = []
        pic_list = []
        gmaps_list = []

        for s in site_list:
            info = query_lookup.get(safe_str(s).lower(), {})
            p_val = info.get("province", "-")
            prov_list.append(p_val)
            if p_val and p_val != "-":
                provinces_collected.append(p_val)
            pic_list.append(info.get("pic", "-"))
            gmaps_list.append(info.get("gmaps", "-"))

        selected_sites["col_site"] = matched[col_site].values
        selected_sites["col_charge"] = (
            matched[col_charge].values if col_charge else ["-"] * len(matched)
        )
        selected_sites["col_province"] = prov_list
        selected_sites["col_pic"] = pic_list
        selected_sites["col_gmaps"] = gmaps_list

    selected_wo = safe_str(first_row[col_wo]) if col_wo else ""
    pekerjaan = safe_str(first_row[col_pekerjaan]) if col_pekerjaan else "-"
    mitra = safe_str(first_row[col_mitra]) if col_mitra else "-"
    sow_type = safe_str(first_row[col_sow]) if col_sow else "Construction"
    date_spk = (
        safe_str(first_row[col_date])
        if col_date
        else datetime.datetime.now().strftime("%d/%m/%Y")
    )

    unique_provs = sorted({p for p in provinces_collected if p and p != "-"})
    lokasi_str = ", ".join(unique_provs) if unique_provs else "-"
    mitra_phone = _lookup_mitra_phone(mitra)

    spk_metadata = {
        "no_spk": safe_str(selected_spk),
        "proyek": "V-Green",
        "pekerjaan": pekerjaan,
        "lokasi": lokasi_str,
        "pic_name": mitra,
        "pic_phone": mitra_phone,
        "sow_type": sow_type,
        "date_spk": date_spk,
    }

    sow_details = pd.DataFrame()
    try:
        master = load_master_sow()
        if not master.empty:
            sow_key = find_column(
                master,
                ["SOW", "Jenis SOW", "Master SOW", "Pekerjaan"],
                fallback=master.columns[0],
            )
            if sow_key:
                sow_details = master[
                    master[sow_key].astype(str).str.strip().str.casefold()
                    == sow_type.strip().casefold()
                ].copy()
    except Exception:
        pass

    return generate_spk_pdf_bytes(
        selected_wo,
        selected_sites,
        spk_metadata,
        sow_details,
        approved=approved,
    )


# ==============================================================================
# RECREATE OPERATION PDF FROM DATABASE
# ==============================================================================

def generate_ms_pdf_from_database(
    df_ms,
    selected_spk,
    approved=False,
):
    if df_ms is None or df_ms.empty:
        return None, None

    col_no_spk = find_column(
        df_ms,
        ["No. SPK", "No SPK", "SPK Number"],
        fallback="No. SPK",
    )

    matched = df_ms[
        df_ms[col_no_spk].astype(str).str.strip() == safe_str(selected_spk)
    ].copy()

    if matched.empty:
        return None, None

    row = matched.iloc[0]

    col_date = find_column(matched, ["Date SPK", "Date"], fallback=None)
    col_pekerjaan = find_column(matched, ["Pekerjaan", "Job"], fallback=None)
    col_mitra = find_column(matched, ["Mitra", "Mitra Name"], fallback=None)
    col_start = find_column(matched, ["Periode Start", "Period Start"], fallback=None)
    col_end = find_column(matched, ["Periode End", "Period End"], fallback=None)

    date_spk = (
        safe_str(row[col_date])
        if col_date
        else datetime.datetime.now().strftime("%d/%m/%Y")
    )
    pekerjaan = safe_str(row[col_pekerjaan]) if col_pekerjaan else "-"
    mitra = safe_str(row[col_mitra]) if col_mitra else "-"
    periode_start = safe_str(row[col_start]) if col_start else "-"
    periode_end = safe_str(row[col_end]) if col_end else "-"

    pekerjaan_lower = pekerjaan.lower()
    if is_maintenance_sow(pekerjaan):
        inferred_sow = pekerjaan
    elif "evcs" in pekerjaan_lower:
        inferred_sow = MAINTENANCE_SOW_EVCS
    elif "bss" in pekerjaan_lower:
        inferred_sow = MAINTENANCE_SOW_BSS
    else:
        inferred_sow = "Maintenance Service"

    spk_metadata = {
        "no_spk": safe_str(selected_spk),
        "date_spk": date_spk,
        "pekerjaan": pekerjaan,
        "mitra": mitra,
        "periode_start": periode_start,
        "periode_end": periode_end,
        "sow_type": inferred_sow,
    }

    return generate_ms_pdf_bytes(spk_metadata, approved=approved)


# ==============================================================================
# MAIN PAGE
# ==============================================================================

def show_spk_page():
    st.markdown(PASTEL_ORANGE_CSS, unsafe_allow_html=True)

    sh = get_cached_connection()
    if not sh:
        st.error("❌ Koneksi ke Database tidak tersedia.")
        return

    # ==========================================================================
    # MAIN MENU
    # ==========================================================================
    main_menu = st.radio(
        "📌 **PILIH MENU OPERASIONAL:**",
        options=[
            "📄 Create SPK Project",
            "🔧 Create SPK Operation",
            "🔄 Take Over Site",
            "🔒 CSO Approval Dashboard",
            "📥 Download PDF",
        ],
        horizontal=True,
        key="main_spk_menu",
    )

    st.markdown("---")

    # ==========================================================================
    # CREATE SPK PROJECT
    # ==========================================================================
    if main_menu == "📄 Create SPK Project":
        st.subheader("📄 Create SPK Project")
        st.markdown(
            "Menu ini khusus untuk pembuatan SPK Project "
            "di luar **Maintenance Service EVCS** dan **Maintenance Service BSS**."
        )

        if get_vm_signature_path() is None:
            st.warning("⚠️ File tanda tangan `assets/Sign VM.jpg` belum ditemukan di folder `assets/`.")

        try:
            query_rows = load_sheet_values_cached(SHEET_QUERY)
            df_dropdown = load_master_dropdown()
            df_master_sow = load_master_sow()

            if not query_rows or len(query_rows) <= 1:
                st.warning("⚠️ Belum ada data pada sheet `Query`.")
            else:
                df_query = dataframe_from_sheet_rows(query_rows)

                sow_dropdown_list = []
                master_sow_col = find_column(df_dropdown, ["Master SOW"], fallback=None)

                if master_sow_col and not df_dropdown.empty:
                    sow_dropdown_list = (
                        df_dropdown[master_sow_col].dropna().unique().tolist()
                    )

                sow_dropdown_list = [
                    safe_str(s) for s in sow_dropdown_list if safe_str(s)
                ]

                if not sow_dropdown_list:
                    sow_dropdown_list = [
                        "Survey BSS",
                        "Instalasi BSS",
                        "Instalasi EVC",
                    ]

                sow_dropdown_list = [
                    s for s in sow_dropdown_list if not is_maintenance_sow(s)
                ]

                col1, col2, col3 = st.columns(3)

                with col1:
                    selected_sow_type = st.selectbox(
                        "Pilih Jenis SOW",
                        options=sow_dropdown_list,
                        key="project_sow_type_select",
                    )

                    sow_category = get_sow_category(selected_sow_type)
                    is_survey = sow_category == "SURVEY"

                    pekerjaan_default = "Survey Location" if is_survey else "Construction"
                    pekerjaan_input = st.text_input(
                        "Pekerjaan",
                        value=pekerjaan_default,
                        key=f"project_pekerjaan_{selected_sow_type}",
                    )

                    target_wo_col = get_query_wo_column(df_query, selected_sow_type)

                    if target_wo_col is not None:
                        raw_wos = (
                            df_query.iloc[:, target_wo_col]
                            .fillna("")
                            .astype(str)
                            .str.strip()
                            .tolist()
                        )
                        wo_list = []
                        seen_wo = set()
                        for wo in raw_wos:
                            if is_sheet_error_value(wo):
                                continue
                            wo_clean = safe_str(wo).strip()
                            if (
                                not wo_clean
                                or wo_clean.lower() in {"nan", "none"}
                                or wo_clean == "-"
                            ):
                                continue
                            if wo_clean not in seen_wo:
                                seen_wo.add(wo_clean)
                                wo_list.append(wo_clean)
                        wo_list = sorted(wo_list, key=lambda x: x.lower())
                    else:
                        wo_list = []

                    wo_label = (
                        f"Pilih Nomor WO ({'Kolom L - Survey' if is_survey else 'Kolom W - Cons'})"
                    )

                    if not wo_list and target_wo_col is None:
                        st.warning(
                            "⚠️ Kolom No. WO tidak ditemukan pada sheet `Query`."
                        )
                    elif not wo_list:
                        detected_wo_header = (
                            df_query.columns[target_wo_col]
                            if isinstance(target_wo_col, int)
                            and target_wo_col < len(df_query.columns)
                            else target_wo_col
                        )
                        st.warning(
                            f"⚠️ Kolom WO terdeteksi sebagai `{detected_wo_header}`, "
                            "tetapi tidak ada nomor WO yang valid."
                        )

                    selected_wo = st.selectbox(
                        wo_label,
                        options=wo_list if wo_list else ["- Tidak ada WO -"],
                        key=f"project_wo_{selected_sow_type}",
                    )

                    proyek_input = st.text_input(
                        "Proyek",
                        value="V-Green",
                        key="project_proyek_input",
                    )

                with col2:
                    if not df_dropdown.empty and len(df_dropdown.columns) > 0:
                        mitra_col_name = df_dropdown.columns[0]
                        mitra_list = (
                            df_dropdown[mitra_col_name].dropna().unique().tolist()
                        )
                    else:
                        mitra_list = []

                    mitra_list = [safe_str(m) for m in mitra_list if safe_str(m)]
                    mitra_list_clean = [
                        m for m in mitra_list if safe_str(m).upper() != "IN HOUSE"
                    ]
                    if "IN HOUSE" not in [safe_str(x).upper() for x in mitra_list_clean]:
                        mitra_list_clean.append("IN HOUSE")

                    selected_pic = st.selectbox(
                        "Penanggung Jawab (Mitra)",
                        options=mitra_list_clean if mitra_list_clean else ["IN HOUSE"],
                        key=f"project_selected_pic_{selected_sow_type}",
                    )

                    manual_pic_name = ""
                    if safe_str(selected_pic).upper() == "IN HOUSE":
                        st.info("🏢 **IN HOUSE** dipilih. Silakan isi nama Penanggung Jawab secara manual.")
                        manual_pic_name = st.text_input(
                            "Nama Penanggung Jawab (IN HOUSE)",
                            value="",
                            key=f"project_manual_pic_{selected_sow_type}",
                        )

                    final_pic_name = resolve_pic_name(selected_pic, manual_pic_name)

                    default_phone = "" if safe_str(selected_pic).upper() == "IN HOUSE" else "0851-8259-6296"
                    if (
                        safe_str(selected_pic).upper() != "IN HOUSE"
                        and not df_dropdown.empty
                        and len(df_dropdown.columns) > 1
                    ):
                        phone_col_name = df_dropdown.columns[1]
                        matched_row = df_dropdown[
                            df_dropdown[mitra_col_name].astype(str).str.strip()
                            == safe_str(selected_pic)
                        ]
                        if not matched_row.empty:
                            default_phone = safe_str(
                                matched_row.iloc[0][phone_col_name],
                                default_phone,
                            )

                    pic_phone = st.text_input(
                        "No. Telepon Penanggung Jawab",
                        value=default_phone,
                        key=f"project_pic_phone_{selected_sow_type}",
                    )

                with col3:
                    st.markdown("**Kontak Penerbit SPK (PT. CLX):**")
                    st.markdown("1. Wikantiyoso Suyono (0878-8855-0300)")
                    st.markdown("2. Hartono (0818-0690-9317)")
                    st.markdown("**Alur Pengesahan:**")
                    st.markdown("• Diperiksa: **Vendor Management** (`Sign VM.jpg`)")
                    st.markdown("• Disetujui: **Chief Services Officer (CSO)**")

                filtered_df = pd.DataFrame()
                selected_sites = pd.DataFrame()

                if selected_wo and selected_wo != "- Tidak ada WO -":
                    if target_wo_col is None:
                        filtered_df = pd.DataFrame()
                    else:
                        wo_series = (
                            df_query.iloc[:, target_wo_col]
                            .fillna("")
                            .astype(str)
                            .str.strip()
                        )
                        filtered_df = df_query[
                            (~wo_series.str.upper().isin(SHEET_ERROR_VALUES))
                            & (wo_series == selected_wo)
                        ].copy()

                    query_charge_col = get_query_data_column(
                        df_query,
                        ["Charging Type", "Charger Type", "Charge Type", "Charging"],
                        legacy_index=2,
                    )
                    query_site_col = get_query_data_column(
                        df_query,
                        ["Site Name", "Site", "Nama Site", "Site ID / Site Name"],
                        legacy_index=5,
                    )
                    query_gmaps_col = get_query_data_column(
                        df_query,
                        ["Gmaps", "Google Maps", "Google Map", "Maps", "Map"],
                        legacy_index=7,
                    )
                    query_province_col = get_query_data_column(
                        df_query,
                        ["Province", "Provinsi", "Province Name"],
                        legacy_index=8,
                    )
                    query_pic_col = get_query_data_column(
                        df_query,
                        ["PIC + Contact", "PIC", "PIC Name", "PIC Contact", "Contact Person", "Penanggung Jawab"],
                        legacy_index=10,
                    )

                    filtered_df["col_charge"] = filtered_df[query_charge_col] if query_charge_col else "-"
                    filtered_df["col_site"] = filtered_df[query_site_col] if query_site_col else "-"
                    filtered_df["col_gmaps"] = filtered_df[query_gmaps_col] if query_gmaps_col else "-"
                    filtered_df["col_province"] = filtered_df[query_province_col] if query_province_col else "-"
                    filtered_df["col_pic"] = filtered_df[query_pic_col] if query_pic_col else "-"
                    filtered_df["col_sow"] = selected_sow_type

                    st.markdown("---")
                    st.markdown(f"### 📍 Daftar Site untuk WO: `{selected_wo}`")

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
                            site_editor_df["Pilih"] = True
                        elif source_col in filtered_df.columns:
                            site_editor_df[display_column_names[source_col]] = filtered_df[source_col].values

                    edited_display_df = st.data_editor(
                        site_editor_df,
                        use_container_width=True,
                        hide_index=True,
                        key=f"project_site_editor_{selected_wo}_{selected_sow_type}",
                        column_config={
                            "Pilih": st.column_config.CheckboxColumn("Pilih", default=True),
                            "Site Name": st.column_config.TextColumn("Site Name", disabled=True),
                            "Charging Type": st.column_config.TextColumn("Charging Type", disabled=True),
                            "Province": st.column_config.TextColumn("Province", disabled=True),
                            "PIC + Contact": st.column_config.TextColumn("PIC + Contact", disabled=True),
                            "Gmaps": st.column_config.TextColumn("Gmaps", disabled=True),
                            "SOW": st.column_config.TextColumn("SOW", disabled=True),
                        },
                    )

                    selected_mask = edited_display_df["Pilih"] == True  # noqa: E712
                    selected_row_indices = edited_display_df.index[selected_mask].tolist()

                    if selected_row_indices:
                        selected_sites = filtered_df.iloc[selected_row_indices].copy()

                can_generate = (
                    selected_wo
                    and selected_wo != "- Tidak ada WO -"
                    and not selected_sites.empty
                )

                if can_generate:
                    st.markdown("---")
                    if st.button(
                        "🚀 Generate SPK Project & Save",
                        type="primary",
                        key="generate_project_spk",
                    ):
                        if safe_str(selected_pic).upper() == "IN HOUSE" and not safe_str(manual_pic_name):
                            st.error("❌ Nama Penanggung Jawab IN HOUSE wajib diisi.")
                            st.stop()

                        if not safe_str(pic_phone):
                            st.error("❌ No. Telepon Penanggung Jawab wajib diisi.")
                            st.stop()

                        with st.spinner("Membuat SPK Project..."):
                            target_db_name = SHEET_SURVEY if is_survey else SHEET_CONS
                            target_sheet = ensure_spk_sheet(sh, target_db_name)
                            existing_rows = load_sheet_values_cached(target_db_name)

                            if not existing_rows or len(existing_rows) <= 1:
                                existing_rows = [DEFAULT_SPK_HEADERS]
                                target_sheet.update(
                                    "A1:N1",
                                    [DEFAULT_SPK_HEADERS],
                                    value_input_option="USER_ENTERED",
                                )
                                clear_sheet_cache(target_db_name)

                            seq_number = get_next_spk_sequence_from_rows(existing_rows)
                            no_spk = generate_spk_number(selected_sow_type, seq_number)

                            provinces = (
                                selected_sites["col_province"]
                                .dropna()
                                .astype(str)
                                .str.strip()
                                .tolist()
                            )
                            unique_provinces = sorted({p for p in provinces if p and p != "-"})
                            auto_lokasi = ", ".join(unique_provinces) if unique_provinces else "-"

                            matched_sow_df = pd.DataFrame()
                            if not df_master_sow.empty:
                                kode_col = df_master_sow.columns[0]
                                matched_sow_df = df_master_sow[
                                    df_master_sow[kode_col].astype(str).str.strip().str.lower()
                                    == safe_str(selected_sow_type).lower()
                                ]

                            spk_metadata = {
                                "no_spk": no_spk,
                                "proyek": proyek_input,
                                "pekerjaan": pekerjaan_input,
                                "lokasi": auto_lokasi,
                                "pic_name": final_pic_name,
                                "pic_phone": pic_phone,
                                "sow_type": selected_sow_type,
                            }

                            no_spk, pdf_bytes = generate_spk_pdf_bytes(
                                selected_wo,
                                selected_sites,
                                spk_metadata,
                                matched_sow_df,
                                approved=False,
                            )

                            now = datetime.datetime.now()
                            current_date_str = now.strftime("%d/%m/%Y")
                            current_time_str = now.strftime("%Y-%m-%d %H:%M:%S")

                            new_rows = []
                            for _, row in selected_sites.iterrows():
                                new_rows.append(
                                    [
                                        current_date_str,
                                        no_spk,
                                        selected_wo,
                                        pekerjaan_input,
                                        "PT CLX",
                                        safe_str(row.get("col_site", "-"), "-"),
                                        safe_str(row.get("col_charge", "-"), "-"),
                                        selected_sow_type,
                                        current_time_str,
                                        final_pic_name,
                                        "",
                                        "",
                                        "",
                                        STATUS_PENDING,
                                    ]
                                )

                            target_sheet.append_rows(new_rows, value_input_option="USER_ENTERED")
                            clear_sheet_cache(target_db_name)

                            safe_filename = f"SPK_{no_spk.replace('/', '_')}_DRAFT.pdf"
                            st.success(
                                f"✅ SPK Project `{no_spk}` berhasil dibuat (termasuk tanda tangan Vendor Management) "
                                "dan dikirim ke **CSO** untuk Approval."
                            )

                            st.download_button(
                                "📥 Download PDF Draft",
                                data=pdf_bytes,
                                file_name=safe_filename,
                                mime="application/pdf",
                                key=f"draft_project_pdf_{no_spk}",
                            )

        except Exception as e:
            st.error(f"❌ Terjadi kesalahan pada Create SPK Project: {e}")

    # ==========================================================================
    # CREATE SPK OPERATION
    # ==========================================================================
    elif main_menu == "🔧 Create SPK Operation":
        st.subheader("🔧 Create SPK Operation")
        st.markdown(
            "Menu ini **khusus** untuk **Maintenance Service EVCS** dan **Maintenance Service BSS**."
        )

        if get_vm_signature_path() is None:
            st.warning("⚠️ File tanda tangan `assets/Sign VM.jpg` belum ditemukan di folder `assets/`.")

        try:
            df_dropdown = load_master_dropdown()
            col1, col2 = st.columns(2)

            with col1:
                selected_ms_sow = st.selectbox(
                    "Pilih Jenis Maintenance",
                    options=MAINTENANCE_SOW_LIST,
                    key="ms_sow_select",
                )

                pekerjaan_ms = st.text_input(
                    "Pekerjaan",
                    value=selected_ms_sow,
                    key="ms_pekerjaan",
                )

                periode_start = st.date_input(
                    "Periode Start",
                    value=datetime.date.today(),
                    key="ms_period_start",
                    format="DD/MM/YYYY",
                )

                periode_end = st.date_input(
                    "Periode End",
                    value=datetime.date.today(),
                    key="ms_period_end",
                    format="DD/MM/YYYY",
                )

            with col2:
                if not df_dropdown.empty and len(df_dropdown.columns) > 0:
                    mitra_col_name = df_dropdown.columns[0]
                    mitra_list = (
                        df_dropdown[mitra_col_name]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .unique()
                        .tolist()
                    )
                    mitra_list = [x for x in mitra_list if x]
                else:
                    mitra_list = []

                if "IN HOUSE" not in [safe_str(x).upper() for x in mitra_list]:
                    mitra_list.append("IN HOUSE")

                selected_ms_mitra = st.selectbox(
                    "Penanggung Jawab / Mitra",
                    options=mitra_list if mitra_list else ["IN HOUSE"],
                    key="ms_mitra_select",
                )

                manual_ms_mitra = ""
                if safe_str(selected_ms_mitra).upper() == "IN HOUSE":
                    manual_ms_mitra = st.text_input(
                        "Nama Pelaksana IN HOUSE",
                        key="ms_manual_mitra",
                    )

                final_ms_mitra = resolve_pic_name(selected_ms_mitra, manual_ms_mitra)

                st.markdown("**Database Target:**")
                st.code(SHEET_MS)

            if periode_end < periode_start:
                st.error("❌ Periode End tidak boleh lebih kecil dari Periode Start.")
            else:
                st.markdown("---")
                st.markdown("### 📋 Preview Database `DB SPK MS`")

                preview_date = datetime.datetime.now().strftime("%d/%m/%Y")
                preview_df = pd.DataFrame(
                    [
                        [
                            preview_date,
                            "Auto Generate",
                            pekerjaan_ms,
                            final_ms_mitra,
                            periode_start.strftime("%d/%m/%Y"),
                            periode_end.strftime("%d/%m/%Y"),
                            STATUS_PENDING,
                        ]
                    ],
                    columns=DEFAULT_MS_HEADERS,
                )

                st.dataframe(preview_df, use_container_width=True, hide_index=True)

                if st.button(
                    "🚀 Generate SPK Operation & Save",
                    type="primary",
                    key="generate_ms_spk",
                ):
                    if safe_str(selected_ms_mitra).upper() == "IN HOUSE" and not safe_str(manual_ms_mitra):
                        st.error("❌ Nama Pelaksana IN HOUSE wajib diisi.")
                        st.stop()

                    with st.spinner("Membuat SPK Operation..."):
                        target_sheet = ensure_ms_sheet(sh)
                        existing_rows = load_sheet_values_cached(SHEET_MS)

                        if not existing_rows or len(existing_rows) <= 1:
                            existing_rows = [DEFAULT_MS_HEADERS]
                            target_sheet.update(
                                "A1:G1",
                                [DEFAULT_MS_HEADERS],
                                value_input_option="USER_ENTERED",
                            )
                            clear_sheet_cache(SHEET_MS)

                        seq_number = get_next_ms_sequence_from_rows(existing_rows)
                        no_spk = generate_ms_number(seq_number)

                        now = datetime.datetime.now()
                        current_date_str = now.strftime("%d/%m/%Y")
                        start_str = periode_start.strftime("%d/%m/%Y")
                        end_str = periode_end.strftime("%d/%m/%Y")

                        new_row = [
                            current_date_str,
                            no_spk,
                            pekerjaan_ms,
                            final_ms_mitra,
                            start_str,
                            end_str,
                            STATUS_PENDING,
                        ]

                        target_sheet.append_rows([new_row], value_input_option="USER_ENTERED")
                        clear_sheet_cache(SHEET_MS)

                        spk_metadata = {
                            "no_spk": no_spk,
                            "date_spk": current_date_str,
                            "pekerjaan": pekerjaan_ms,
                            "mitra": final_ms_mitra,
                            "periode_start": start_str,
                            "periode_end": end_str,
                            "sow_type": selected_ms_sow,
                        }

                        _, pdf_bytes = generate_ms_pdf_bytes(spk_metadata, approved=False)

                        safe_filename = f"SPK_{no_spk.replace('/', '_')}_DRAFT.pdf"
                        st.success(
                            f"✅ SPK Operation `{no_spk}` untuk **{selected_ms_sow}** "
                            "berhasil dibuat dan dikirim ke **CSO** untuk Approval."
                        )
                        st.info(f"📌 Periode: **{start_str} s/d {end_str}**")

                        st.download_button(
                            "📥 Download PDF Draft",
                            data=pdf_bytes,
                            file_name=safe_filename,
                            mime="application/pdf",
                            key=f"draft_ms_pdf_{no_spk}",
                        )

        except Exception as e:
            st.error(f"❌ Terjadi kesalahan pada Create SPK Operation: {e}")

    # ==========================================================================
    # TAKE OVER SITE
    # ==========================================================================
    elif main_menu == "🔄 Take Over Site":
        st.subheader("🔄 Menu Take Over Site")
        st.markdown(
            "Pilih **Sheet Target** (Survey / Cons) dan **Site Name** yang akan di-*Take Over*. "
            "Sistem akan mencatat nomor SPK Take Over baru pada riwayat kolom K:M."
        )

        try:
            db_target_type = st.radio(
                "Pilih Kategori Database SPK",
                options=[SHEET_CONS, SHEET_SURVEY],
                horizontal=True,
                key="takeover_db_choice",
            )

            target_sheet_to = sh.worksheet(db_target_type)
            to_rows = load_sheet_values_cached(db_target_type)

            if len(to_rows) <= 1:
                st.warning(f"⚠️ Sheet `{db_target_type}` belum memiliki data.")
            else:
                df_to = dataframe_from_sheet_rows(to_rows)

                col_site_name = find_column(df_to, ["Site Name", "Site", "Nama Site"], fallback="Site Name")
                col_spk_num = find_column(df_to, ["No. SPK", "No SPK", "SPK"], fallback="No. SPK")
                col_mitra_name = find_column(df_to, ["Mitra", "Mitra Name"], fallback="Mitra")

                record_options = {}
                for _, r in df_to.iterrows():
                    s_val = safe_str(r.get(col_site_name, ""))
                    spk_val = safe_str(r.get(col_spk_num, "-"), "-")
                    if not s_val or s_val == "-":
                        continue
                    label = f"{s_val}  |  {spk_val}"
                    if label not in record_options:
                        record_options[label] = r

                col_to1, col_to2 = st.columns(2)

                with col_to1:
                    selected_label_to = st.selectbox(
                        "Pilih Site Name & No. SPK yang akan di Take Over",
                        options=list(record_options.keys()) if record_options else ["- Tidak Ada Site -"],
                        key="takeover_site_select",
                    )

                if selected_label_to and selected_label_to != "- Tidak Ada Site -":
                    matched_row_to = record_options[selected_label_to]
                    selected_site_to = safe_str(matched_row_to.get(col_site_name, "-"))
                    old_spk = safe_str(matched_row_to.get(col_spk_num, "-"), "-")
                    old_mitra = safe_str(matched_row_to.get(col_mitra_name, "-"), "-")
                    target_row_number = int(matched_row_to["_excel_row"])

                    with col_to2:
                        st.info(f"📌 **No SPK Lama:** {old_spk}")
                        st.info(f"👤 **Mitra Lama:** {old_mitra}")

                    st.markdown("### 📝 Informasi Take Over Baru")
                    col_to3, col_to4 = st.columns(2)

                    df_dropdown_to = load_master_dropdown()
                    new_mitra_list = []

                    if not df_dropdown_to.empty and len(df_dropdown_to.columns) > 0:
                        mitra_col = df_dropdown_to.columns[0]
                        new_mitra_list = (
                            df_dropdown_to[mitra_col]
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .unique()
                            .tolist()
                        )
                        new_mitra_list = [x for x in new_mitra_list if x]

                    with col_to3:
                        new_mitra = st.selectbox(
                            "Pilih Mitra Baru (Pengambil Alih)",
                            options=new_mitra_list if new_mitra_list else ["Mitra Baru"],
                            key="takeover_mitra_select",
                        )
                        sow_type_to = "Survey BSS" if "Survey" in db_target_type else "Instalasi BSS"

                    with col_to4:
                        today_date_str = datetime.datetime.now().strftime("%d/%m/%Y")
                        st.text_input(
                            "Tanggal Take Over (Otomatis)",
                            value=today_date_str,
                            disabled=True,
                            key="takeover_date_display",
                        )

                    if st.button(
                        "🔥 Proses & Simpan Take Over Site",
                        type="primary",
                        key="process_takeover",
                    ):
                        with st.spinner("Memproses Take Over..."):
                            seq_number = get_next_spk_sequence_from_rows(to_rows)
                            new_spk_no = generate_spk_number(sow_type_to, sequence_num=seq_number)

                            takeover_updates = [
                                {
                                    "range": f"K{target_row_number}:M{target_row_number}",
                                    "values": [
                                        [
                                            today_date_str,
                                            new_spk_no,
                                            safe_str(new_mitra),
                                        ]
                                    ],
                                }
                            ]

                            batch_update_cells(target_sheet_to, takeover_updates)
                            clear_sheet_cache(db_target_type)

                            st.success(
                                f"🎉 Take Over Berhasil!\n\n"
                                f"• Site: `{selected_site_to}`\n"
                                f"• No. SPK Take Over Baru: `{new_spk_no}`\n"
                                f"• SPK Lama: `{old_spk}`\n"
                                f"• Mitra Baru: `{new_mitra}`"
                            )

        except Exception as e:
            st.error(f"❌ Terjadi kesalahan pada Menu Take Over: {e}")

    # ==========================================================================
    # CSO APPROVAL DASHBOARD (PROTECTED WITH PASSWORD: clx0007)
    # ==========================================================================
    elif main_menu == "🔒 CSO Approval Dashboard":
        st.subheader("🔒 CSO Approval Dashboard")
        st.markdown(
            "Halaman khusus **Chief Services Officer (CSO)** untuk meninjau, menyetujui (**Approve**), "
            "atau menolak (**Reject**) pengajuan SPK Project maupun Operation."
        )

        pin_input = st.text_input(
            "Masukkan Password Khusus CSO:",
            type="password",
            placeholder="Masukkan password CSO...",
            key="cso_pin_input",
        )

        if pin_input == CSO_PIN_SECRET:
            st.success("🔓 Akses Diterima! Selamat datang, Chief Services Officer (CSO).")

            if get_cso_signature_path() is None:
                st.warning("⚠️ File tanda tangan `assets/Sign CSO.jpg` belum ditemukan di folder `assets/`.")

            col_db_sel, col_ref_btn = st.columns([4, 1])
            with col_db_sel:
                db_approval_type = st.radio(
                    "Pilih Kategori Database SPK",
                    options=[SHEET_CONS, SHEET_SURVEY, SHEET_MS],
                    horizontal=True,
                    key="approval_db_choice",
                )
            with col_ref_btn:
                if st.button("🔄 Refresh Data", key="refresh_cso_approval_data"):
                    clear_sheet_cache(db_approval_type)
                    st.rerun()

            try:
                sheet_appr = sh.worksheet(db_approval_type)
                appr_rows = load_sheet_values_cached(db_approval_type)

                if len(appr_rows) <= 1:
                    st.info(f"ℹ️ Belum ada data pada `{db_approval_type}`.")
                else:
                    df_appr = dataframe_from_sheet_rows(appr_rows)
                    df_appr, col_status = ensure_status_column(sheet_appr, df_appr)

                    col_no_spk = find_column(
                        df_appr,
                        ["No. SPK", "No SPK", "SPK Number", "Nomor SPK"],
                        fallback=None,
                    )
                    col_date_spk = find_column(
                        df_appr,
                        ["Date SPK", "Date", "Tanggal SPK", "Tanggal"],
                        fallback=None,
                    )
                    col_no_wo = find_column(
                        df_appr,
                        ["No. WO", "No WO", "WO Number", "Nomor WO"],
                        fallback=None,
                    )
                    col_pekerjaan = find_column(
                        df_appr,
                        ["Pekerjaan", "Job", "Work"],
                        fallback=None,
                    )
                    col_site = find_column(
                        df_appr,
                        ["Site Name", "Site", "Nama Site"],
                        fallback=None,
                    )
                    col_mitra = find_column(
                        df_appr,
                        ["Mitra", "Mitra Name", "Partner"],
                        fallback=None,
                    )

                    if not col_no_spk:
                        st.error("❌ Kolom No. SPK tidak ditemukan.")
                        st.stop()

                    df_appr[col_status] = (
                        df_appr[col_status].fillna("").astype(str).str.strip()
                    )

                    pending_mask = [
                        _is_pending_status(st_val, spk_val)
                        for st_val, spk_val in zip(
                            df_appr[col_status].tolist(),
                            df_appr[col_no_spk].tolist(),
                        )
                    ]
                    pending_df = df_appr[pending_mask].copy()

                    st.markdown("---")
                    st.markdown(f"### 📋 Daftar SPK Menunggu Approval CSO ({len(pending_df)} baris)")

                    if pending_df.empty:
                        st.success("🎉 Tidak ada SPK yang menunggu approval CSO.")
                    else:
                        pending_spks = [
                            x
                            for x in pending_df[col_no_spk]
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .unique()
                            .tolist()
                            if x
                        ]

                        selected_appr_spk = st.selectbox(
                            "Pilih Nomor SPK yang akan ditinjau:",
                            options=pending_spks,
                            key=f"approval_spk_select_{db_approval_type}",
                        )

                        if selected_appr_spk:
                            spk_details = pending_df[
                                pending_df[col_no_spk].astype(str).str.strip()
                                == safe_str(selected_appr_spk)
                            ]

                            display_columns = []
                            for col in [
                                col_date_spk,
                                col_no_spk,
                                col_no_wo,
                                col_pekerjaan,
                                col_site,
                                col_mitra,
                            ]:
                                if col and col in spk_details.columns and col not in display_columns:
                                    display_columns.append(col)

                            if db_approval_type == SHEET_MS:
                                for col in ["Periode Start", "Periode End"]:
                                    if col in spk_details.columns:
                                        display_columns.append(col)

                            if col_status and col_status in spk_details.columns and col_status not in display_columns:
                                display_columns.append(col_status)

                            if display_columns:
                                st.dataframe(
                                    spk_details[display_columns],
                                    use_container_width=True,
                                    hide_index=True,
                                )
                            else:
                                st.dataframe(
                                    spk_details,
                                    use_container_width=True,
                                    hide_index=True,
                                )

                            col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 1.3])

                            with col_btn1:
                                if st.button(
                                    "✅ APPROVE SPK (CSO)",
                                    type="primary",
                                    key=f"approve_{db_approval_type}_{selected_appr_spk}",
                                ):
                                    data_cols = [c for c in df_appr.columns if c != "_excel_row"]
                                    status_col_index = data_cols.index(col_status) + 1
                                    status_letter = column_letter(status_col_index)

                                    update_ranges = [
                                        {
                                            "range": f"{status_letter}{int(excel_r)}",
                                            "values": [[STATUS_APPROVED]],
                                        }
                                        for excel_r in spk_details["_excel_row"].tolist()
                                    ]

                                    if update_ranges:
                                        batch_update_cells(sheet_appr, update_ranges)
                                        clear_sheet_cache(db_approval_type)
                                        st.success(
                                            f"✅ SPK `{selected_appr_spk}` BERHASIL DI-APPROVE OLEH CSO!"
                                        )
                                        st.rerun()

                            with col_btn2:
                                if st.button(
                                    "❌ REJECT SPK",
                                    key=f"reject_{db_approval_type}_{selected_appr_spk}",
                                ):
                                    data_cols = [c for c in df_appr.columns if c != "_excel_row"]
                                    status_col_index = data_cols.index(col_status) + 1
                                    status_letter = column_letter(status_col_index)

                                    update_ranges = [
                                        {
                                            "range": f"{status_letter}{int(excel_r)}",
                                            "values": [[STATUS_REJECTED]],
                                        }
                                        for excel_r in spk_details["_excel_row"].tolist()
                                    ]

                                    if update_ranges:
                                        batch_update_cells(sheet_appr, update_ranges)
                                        clear_sheet_cache(db_approval_type)
                                        st.warning(f"❌ SPK `{selected_appr_spk}` DITOLAK!")
                                        st.rerun()

                            with col_btn3:
                                if db_approval_type == SHEET_MS:
                                    _, preview_pdf = generate_ms_pdf_from_database(
                                        df_appr, selected_appr_spk, approved=False
                                    )
                                else:
                                    _, preview_pdf = generate_project_pdf_from_database(
                                        df_appr, selected_appr_spk, approved=False
                                    )
                                if preview_pdf:
                                    st.download_button(
                                        "👁️ Preview PDF Draft",
                                        data=preview_pdf,
                                        file_name=f"PREVIEW_{selected_appr_spk.replace('/', '_')}.pdf",
                                        mime="application/pdf",
                                        key=f"cso_preview_{selected_appr_spk}",
                                    )

            except Exception as e:
                st.error(f"❌ Terjadi kesalahan pada Dashboard Approval CSO: {e}")

        elif pin_input != "":
            st.error("❌ Password Salah! Akses CSO ditolak.")

    # ==========================================================================
    # DOWNLOAD PDF
    # ==========================================================================
    elif main_menu == "📥 Download PDF":
        st.subheader("📥 Download SPK PDF (Approved by CSO)")
        st.markdown(
            "PDF resmi dengan tanda tangan lengkap (**Vendor Management** & **CSO**) "
            "dapat di-download setelah SPK mendapatkan status **Approved by CSO**."
        )

        try:
            download_type = st.radio(
                "Pilih Jenis SPK",
                options=["SPK Project", "SPK Operation"],
                horizontal=True,
                key="download_pdf_type",
            )

            if download_type == "SPK Project":
                project_db = st.radio(
                    "Pilih Database Project",
                    options=[SHEET_CONS, SHEET_SURVEY],
                    horizontal=True,
                    key="download_project_db",
                )

                rows = load_sheet_values_cached(project_db)

                if len(rows) <= 1:
                    st.info("ℹ️ Belum ada data SPK.")
                else:
                    df_download = dataframe_from_sheet_rows(rows)
                    col_no_spk = find_column(
                        df_download,
                        ["No. SPK", "No SPK", "SPK Number"],
                        fallback=None,
                    )
                    col_status = find_column(
                        df_download,
                        ["Status Approval", "Approval Status", "Status"],
                        fallback=None,
                    )

                    if not col_no_spk or not col_status:
                        st.error("❌ Kolom No. SPK atau Status Approval tidak ditemukan.")
                    else:
                        approved_df = df_download[
                            df_download[col_status].apply(_is_approved_status)
                        ]

                        approved_spks = [
                            x
                            for x in approved_df[col_no_spk]
                            .astype(str)
                            .str.strip()
                            .unique()
                            .tolist()
                            if x
                        ]

                        if not approved_spks:
                            st.info("ℹ️ Belum ada SPK Project yang Approved by CSO.")
                        else:
                            selected_download_spk = st.selectbox(
                                "Pilih SPK Approved",
                                options=approved_spks,
                                key="download_project_spk",
                            )

                            if selected_download_spk:
                                st.success(
                                    f"✅ SPK `{selected_download_spk}` sudah Approved dan siap didownload."
                                )

                                if st.button(
                                    "📄 Generate PDF Approved",
                                    type="primary",
                                    key="generate_approved_project_pdf",
                                ):
                                    with st.spinner("Menyiapkan PDF Approved..."):
                                        no_spk_pdf, pdf_bytes = generate_project_pdf_from_database(
                                            df_download,
                                            selected_download_spk,
                                            approved=True,
                                        )

                                        if pdf_bytes:
                                            filename = (
                                                f"SPK_{no_spk_pdf.replace('/', '_')}_APPROVED.pdf"
                                            )
                                            st.download_button(
                                                "📥 Download PDF SPK Project",
                                                data=pdf_bytes,
                                                file_name=filename,
                                                mime="application/pdf",
                                                type="primary",
                                                key=f"download_project_pdf_{selected_download_spk}",
                                            )

            else:
                rows = load_sheet_values_cached(SHEET_MS)

                if len(rows) <= 1:
                    st.info("ℹ️ Belum ada data pada `DB SPK MS`.")
                else:
                    df_download = dataframe_from_sheet_rows(rows)
                    col_no_spk = find_column(
                        df_download,
                        ["No. SPK", "No SPK", "SPK Number"],
                        fallback=None,
                    )
                    col_status = find_column(
                        df_download,
                        ["Status Approval", "Approval Status", "Status"],
                        fallback=None,
                    )

                    if not col_no_spk or not col_status:
                        st.error("❌ Kolom No. SPK atau Status Approval tidak ditemukan pada DB SPK MS.")
                    else:
                        approved_df = df_download[
                            df_download[col_status].apply(_is_approved_status)
                        ]

                        approved_spks = [
                            x
                            for x in approved_df[col_no_spk]
                            .astype(str)
                            .str.strip()
                            .unique()
                            .tolist()
                            if x
                        ]

                        if not approved_spks:
                            st.info("ℹ️ Belum ada SPK Operation yang Approved by CSO.")
                        else:
                            selected_download_spk = st.selectbox(
                                "Pilih SPK Operation Approved",
                                options=approved_spks,
                                key="download_ms_spk",
                            )

                            if selected_download_spk:
                                selected_rows = approved_df[
                                    approved_df[col_no_spk].astype(str).str.strip()
                                    == selected_download_spk
                                ]
                                display_cols = [c for c in selected_rows.columns if c != "_excel_row"]
                                st.dataframe(
                                    selected_rows[display_cols],
                                    use_container_width=True,
                                    hide_index=True,
                                )

                                st.success(
                                    f"✅ SPK Operation `{selected_download_spk}` sudah Approved dan siap didownload."
                                )

                                if st.button(
                                    "📄 Generate PDF Approved",
                                    type="primary",
                                    key="generate_approved_ms_pdf",
                                ):
                                    with st.spinner("Menyiapkan PDF Operation..."):
                                        no_spk_pdf, pdf_bytes = generate_ms_pdf_from_database(
                                            df_download,
                                            selected_download_spk,
                                            approved=True,
                                        )

                                        if pdf_bytes:
                                            filename = (
                                                f"SPK_{no_spk_pdf.replace('/', '_')}_APPROVED.pdf"
                                            )
                                            st.download_button(
                                                "📥 Download PDF SPK Operation",
                                                data=pdf_bytes,
                                                file_name=filename,
                                                mime="application/pdf",
                                                type="primary",
                                                key=f"download_ms_pdf_{selected_download_spk}",
                                            )

        except Exception as e:
            st.error(f"❌ Terjadi kesalahan pada Download PDF: {e}")


# ==============================================================================
# RUN APPLICATION
# ==============================================================================

if __name__ == "__main__":
    show_spk_page()
