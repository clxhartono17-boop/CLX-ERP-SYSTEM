# ==============================================================================
# DATABASE.PY
# ERP DATABASE LAYER - GOOGLE SHEETS
#
# VERSION:
# DATABASE.PY V2.2
# Rate Limit Protected / Cached / Retry / Backoff / Multi User Safer
#
# FEATURES:
# 1. Streamlit Resource Cache
# 2. Streamlit Data Cache
# 3. Worksheet Resource Cache
# 4. Centralized Google Sheets Read Limiter
# 5. Exponential Backoff + Jitter
# 6. Centralized Safe Google Sheets Read
# 7. Domain-Based Cache Invalidation
# 8. Reduced API Read Requests
# 9. Reimbursement Database
# 10. Material Out / Delivery Order Database
# 11. Query Database
# 12. Authorization Database
# 13. Google Drive Image Upload via Apps Script
# 14. Targeted DO Update - NO worksheet.clear()
# 15. Central Document Sequence
# 16. Multi-User Safer Document Number Allocation
# 17. Internal Row Number Allocation from Google Sheets Append Response
# 18. Database Health Check
# 19. Rate Limit Debugging
# 20. Compatibility Helper: get_sheet_values()
# 21. Global 429 Cooldown
# 22. Retry-After Detection
# 23. Multi-Thread Read Protection
#
# IMPORTANT:
# - Sheet "DB Sequence" akan dibuat otomatis jika belum ada.
# - Jangan menghapus isi Sheet "DB Sequence" setelah sistem mulai digunakan.
# - Nomor dokumen yang sudah teralokasi dapat memiliki gap jika user membuat
#   nomor lalu membatalkan transaksi. Ini NORMAL untuk sistem ERP.
# ==============================================================================


# ==============================================================================
# 0. IMPORT
# ==============================================================================

import os
import re
import time
import base64
import random
import threading
import uuid

from collections import deque
from datetime import datetime

import requests
import gspread
import pandas as pd
import streamlit as st

from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, WorksheetNotFound


# ==============================================================================
# 1. GLOBAL CONFIGURATION
# ==============================================================================

SPREADSHEET_ID = (
    "1FU1lL3ls3jP_hAxBdx_Fu35Z9Ap4ICdHmOpMvCyA3gY"
)


# ------------------------------------------------------------------------------
# DATABASE SHEETS
# ------------------------------------------------------------------------------

SHEET_REIMBURSEMENT = "DB Reimbursement"

SHEET_MATERIAL_OUT = "DB Material Out"

SHEET_QUERY = "Query"

SHEET_AUTHORIZATION = "Otorisasi"

SHEET_SEQUENCE = "DB Sequence"


# ------------------------------------------------------------------------------
# GOOGLE DRIVE
# ------------------------------------------------------------------------------

GOOGLE_DRIVE_ROOT_FOLDER_ID = (
    "1fto5kD7X_pYT21F6Qr1RLfmBSmEb1O3o"
)


# ------------------------------------------------------------------------------
# APPS SCRIPT WEB APP
# ------------------------------------------------------------------------------

APPS_SCRIPT_WEB_APP_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbwAx3pGoDtMLI7CZV58WoNSeKo2oHx3jCs8IARlAagUvaAVRAWkoLeZ1H_4P0RMpD6p/"
    "exec"
)


# ------------------------------------------------------------------------------
# GOOGLE API SCOPES
# ------------------------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]


# ==============================================================================
# 2. RATE LIMIT CONFIGURATION
# ==============================================================================

# ------------------------------------------------------------------------------
# IMPORTANT
#
# Google API quota:
#
#   Read requests per minute per user = 60
#
# Kita sengaja menggunakan limit internal yang lebih rendah agar lebih aman.
#
# 20 request/minute dipilih agar masih tersedia ruang untuk:
# - worksheet discovery
# - spreadsheet connection
# - concurrent Streamlit users
# - request lain yang mungkin terjadi di luar database.py
# ------------------------------------------------------------------------------

MAX_READ_REQUESTS_PER_MINUTE = 20

RATE_WINDOW_SECONDS = 60

READ_BACKOFF_MAX_SECONDS = 60

# Cooldown global setelah menerima HTTP 429.
#
# Ketika satu request terkena quota:
# seluruh thread akan berhenti sementara sebelum mencoba request berikutnya.
#
# Ini sangat penting untuk multi-user Streamlit.
RATE_LIMIT_COOLDOWN_SECONDS = 10


# ------------------------------------------------------------------------------
# READ REQUEST TIMESTAMP STORAGE
# ------------------------------------------------------------------------------

_read_request_times = deque()

_read_lock = threading.Lock()


# ------------------------------------------------------------------------------
# GLOBAL RATE LIMIT COOLDOWN
# ------------------------------------------------------------------------------

_next_allowed_read_time = 0.0


# ------------------------------------------------------------------------------
# WRITE LOCK
# ------------------------------------------------------------------------------

_write_lock = threading.Lock()


# ==============================================================================
# 3. RATE LIMITER
# ==============================================================================

def wait_for_read_slot():
    """
    Central internal READ rate limiter.

    Maksimum:
        20 logical READ request / 60 detik.

    Selain rate window, fungsi ini juga menghormati global cooldown
    apabila sebelumnya Google mengembalikan HTTP 429.

    Semua READ yang masuk ke Google Sheets sebaiknya melewati fungsi ini.
    """

    global _next_allowed_read_time

    while True:

        now = time.time()

        with _read_lock:

            # ------------------------------------------------------------------
            # REMOVE EXPIRED REQUEST TIMESTAMPS
            # ------------------------------------------------------------------

            while (
                _read_request_times
                and
                (
                    now
                    -
                    _read_request_times[0]
                    >= RATE_WINDOW_SECONDS
                )
            ):

                _read_request_times.popleft()

            # ------------------------------------------------------------------
            # GLOBAL 429 COOLDOWN
            # ------------------------------------------------------------------

            cooldown_remaining = (
                _next_allowed_read_time
                -
                now
            )

            if cooldown_remaining > 0:

                wait_time = cooldown_remaining

            # ------------------------------------------------------------------
            # NORMAL RATE WINDOW
            # ------------------------------------------------------------------

            elif (
                len(_read_request_times)
                <
                MAX_READ_REQUESTS_PER_MINUTE
            ):

                _read_request_times.append(now)

                return

            # ------------------------------------------------------------------
            # RATE WINDOW FULL
            # ------------------------------------------------------------------

            else:

                wait_time = (
                    RATE_WINDOW_SECONDS
                    -
                    (
                        now
                        -
                        _read_request_times[0]
                    )
                    +
                    0.5
                )

        time.sleep(
            max(
                wait_time,
                0.5
            )
        )


# ==============================================================================
# 4. RATE LIMIT ERROR DETECTOR
# ==============================================================================

def is_rate_limit_error(error):
    """
    Mendeteksi berbagai bentuk error quota/rate limit Google.
    """

    error_text = str(
        error
    ).upper()

    keywords = [
        "429",
        "RESOURCE_EXHAUSTED",
        "RATE_LIMIT_EXCEEDED",
        "QUOTA_EXCEEDED",
        "TOO MANY REQUESTS",
        "READREQUESTSPERMINUTE",
        "USER_RATE_LIMIT_EXCEEDED",
        "PER_USER_RATE_LIMIT",
        "RATE LIMIT"
    ]

    return any(
        keyword in error_text
        for keyword in keywords
    )


# ==============================================================================
# 4A. RETRY-AFTER DETECTOR
# ==============================================================================

def _get_retry_after_seconds(
    error
):
    """
    Mencoba mengambil nilai Retry-After dari response Google.

    Jika tidak tersedia, return 0.
    """

    try:

        response = getattr(
            error,
            "response",
            None
        )

        if response is None:

            return 0

        headers = getattr(
            response,
            "headers",
            {}
        )

        retry_after = headers.get(
            "Retry-After"
        )

        if retry_after is None:

            return 0

        return max(
            float(
                retry_after
            ),
            0
        )

    except Exception:

        return 0


# ==============================================================================
# 4B. GLOBAL 429 COOLDOWN
# ==============================================================================

def _activate_rate_limit_cooldown(
    seconds=None
):
    """
    Mengaktifkan cooldown global setelah menerima 429.

    Semua thread yang menggunakan wait_for_read_slot()
    akan ikut menunggu cooldown ini.
    """

    global _next_allowed_read_time

    if seconds is None:

        seconds = (
            RATE_LIMIT_COOLDOWN_SECONDS
        )

    try:

        seconds = float(
            seconds
        )

    except Exception:

        seconds = (
            RATE_LIMIT_COOLDOWN_SECONDS
        )

    seconds = max(
        seconds,
        RATE_LIMIT_COOLDOWN_SECONDS
    )

    with _read_lock:

        cooldown_until = (
            time.time()
            +
            seconds
        )

        if (
            cooldown_until
            >
            _next_allowed_read_time
        ):

            _next_allowed_read_time = (
                cooldown_until
            )


# ==============================================================================
# 5. SAFE GOOGLE SHEETS READ
# ==============================================================================

def safe_sheet_read(
    read_function,
    max_retries=6
):
    """
    Centralized Google Sheets READ handler.

    Proteksi:
        1. Internal rate limiter
        2. 429 detection
        3. Global cooldown
        4. Exponential backoff
        5. Random jitter
        6. Retry-After support
        7. Multi-thread protection
    """

    last_error = None

    for attempt in range(
        max_retries
    ):

        try:

            # --------------------------------------------------------------
            # WAIT UNTIL REQUEST IS ALLOWED
            # --------------------------------------------------------------

            wait_for_read_slot()

            # --------------------------------------------------------------
            # ACTUAL GOOGLE SHEETS READ
            # --------------------------------------------------------------

            return read_function()

        except Exception as e:

            last_error = e

            # --------------------------------------------------------------
            # NON RATE-LIMIT ERROR
            # --------------------------------------------------------------

            if not is_rate_limit_error(
                e
            ):

                raise

            # --------------------------------------------------------------
            # RETRY-AFTER
            # --------------------------------------------------------------

            retry_after = (
                _get_retry_after_seconds(
                    e
                )
            )

            # --------------------------------------------------------------
            # EXPONENTIAL BACKOFF
            # --------------------------------------------------------------

            base_delay = min(
                2 ** attempt,
                READ_BACKOFF_MAX_SECONDS
            )

            # --------------------------------------------------------------
            # JITTER
            # --------------------------------------------------------------

            jitter = random.uniform(
                0.75,
                1.50
            )

            calculated_delay = (
                base_delay
                *
                jitter
            )

            # --------------------------------------------------------------
            # FINAL DELAY
            # --------------------------------------------------------------

            sleep_time = max(
                calculated_delay,
                retry_after,
                RATE_LIMIT_COOLDOWN_SECONDS
            )

            # --------------------------------------------------------------
            # GLOBAL COOLDOWN
            # --------------------------------------------------------------

            _activate_rate_limit_cooldown(
                sleep_time
            )

            print(
                f"⚠️ Google Sheets RATE LIMIT / 429. "
                f"Retry {attempt + 1}/{max_retries} "
                f"dalam {sleep_time:.1f} detik..."
            )

            time.sleep(
                sleep_time
            )

    # --------------------------------------------------------------------------
    # ALL RETRIES FAILED
    # --------------------------------------------------------------------------

    raise last_error


# ==============================================================================
# 6. GOOGLE SHEET CONNECTION
# ==============================================================================

@st.cache_resource(
    show_spinner=False
)
def get_google_sheet_connection():
    """
    Membuka koneksi utama Google Spreadsheet.
    """

    try:

        # ----------------------------------------------------------------------
        # Streamlit Secrets
        # ----------------------------------------------------------------------

        if (
            "gcp_service_account"
            in
            st.secrets
        ):

            creds_dict = dict(
                st.secrets[
                    "gcp_service_account"
                ]
            )

            creds = (
                Credentials
                .from_service_account_info(
                    creds_dict,
                    scopes=SCOPES
                )
            )

        # ----------------------------------------------------------------------
        # Local service_account.json
        # ----------------------------------------------------------------------

        elif os.path.exists(
            "service_account.json"
        ):

            creds = (
                Credentials
                .from_service_account_file(
                    "service_account.json",
                    scopes=SCOPES
                )
            )

        else:

            st.error(
                "❌ Konfigurasi Service Account tidak ditemukan.\n\n"
                "Pastikan salah satu tersedia:\n"
                "• st.secrets['gcp_service_account']\n"
                "• service_account.json"
            )

            st.stop()

        # ----------------------------------------------------------------------
        # Authorize
        # ----------------------------------------------------------------------

        gc = gspread.authorize(
            creds
        )

        # ----------------------------------------------------------------------
        # Open spreadsheet
        # ----------------------------------------------------------------------

        sh = gc.open_by_key(
            SPREADSHEET_ID
        )

        return sh

    except Exception as e:

        st.error(
            "❌ Gagal terhubung ke Google Sheets.\n\n"
            f"{type(e).__name__}: {str(e)}"
        )

        st.stop()


# ==============================================================================
# 7. CACHED WORKSHEET ACCESS
# ==============================================================================

@st.cache_resource(
    show_spinner=False
)
def _get_worksheet_cached(
    sheet_name
):
    """
    Mengambil object worksheet dan menyimpannya sebagai resource.

    CATATAN:
    sh.worksheet() sendiri merupakan READ request.
    Karena function ini sudah menggunakan cache_resource,
    request hanya dilakukan saat worksheet belum tersedia di cache.
    """

    sh = get_google_sheet_connection()

    return sh.worksheet(
        sheet_name
    )


def get_worksheet(
    sheet_name
):
    """
    Public helper untuk mendapatkan worksheet.
    """

    return _get_worksheet_cached(
        sheet_name
    )


# ==============================================================================
# 8. GENERAL SHEET READ COMPATIBILITY HELPERS
# ==============================================================================

def get_sheet_values(
    sheet_name,
    default=None
):
    """
    COMPATIBILITY HELPER.

    Fungsi ini dibuat agar modul lain dapat menggunakan:

        from core.database import get_sheet_values

    tanpa harus mengetahui detail worksheet/cache/rate limiter.

    Return:
        list[list[str]]

    Contoh:

        values = get_sheet_values("Query")

    Semua READ tetap melewati safe_sheet_read().
    """

    if default is None:
        default = []

    try:

        worksheet = get_worksheet(
            sheet_name
        )

        def read_values():

            return worksheet.get_all_values()

        values = safe_sheet_read(
            read_values
        )

        if values is None:
            return default

        return values

    except Exception as e:

        if is_rate_limit_error(
            e
        ):

            st.warning(
                f"⚠️ Google Sheets API quota tercapai "
                f"saat membaca Sheet '{sheet_name}'. "
                f"Silakan tunggu beberapa detik."
            )

        else:

            st.warning(
                f"⚠️ Gagal membaca Sheet '{sheet_name}': {e}"
            )

        return default


def get_sheet_dataframe(
    sheet_name
):
    """
    Compatibility helper tambahan.

    Membaca worksheet menjadi DataFrame.

    Tidak mengubah fungsi database lama.
    """

    try:

        values = get_sheet_values(
            sheet_name
        )

        if not values or len(values) < 2:

            return pd.DataFrame()

        headers = list(
            values[0]
        )

        rows = []

        for row in values[1:]:

            row = list(row)

            if len(row) < len(headers):

                row += (
                    [""] *
                    (
                        len(headers)
                        -
                        len(row)
                    )
                )

            elif len(row) > len(headers):

                row = row[
                    :len(headers)
                ]

            rows.append(
                row
            )

        return pd.DataFrame(
            rows,
            columns=headers
        )

    except Exception:

        return pd.DataFrame()


# ==============================================================================
# 9. GENERAL HELPERS
# ==============================================================================

def get_roman_month(
    month_int
):

    roman_months = {

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
        12: "XII"

    }

    return roman_months.get(
        month_int,
        "I"
    )


def safe_int(
    value,
    default=0
):

    try:

        if value is None:

            return default

        text = str(
            value
        ).strip()

        if not text:

            return default

        return int(
            float(
                text.replace(
                    ",",
                    ""
                )
            )
        )

    except Exception:

        return default


def safe_float(
    value,
    default=0.0
):

    try:

        if value is None:

            return default

        text = (
            str(value)
            .replace(
                ",",
                ""
            )
            .replace(
                "Rp",
                ""
            )
            .replace(
                "rp",
                ""
            )
            .strip()
        )

        if not text:

            return default

        return float(
            text
        )

    except Exception:

        return default


def normalize_text(
    value
):

    if value is None:

        return ""

    return (
        str(value)
        .strip()
        .lower()
    )


def _parse_row_number_from_a1(
    range_text
):

    if not range_text:

        return None

    match = re.search(
        r"!?[A-Z]+(\d+)",
        str(
            range_text
        )
    )

    if not match:

        return None

    try:

        return int(
            match.group(1)
        )

    except Exception:

        return None


def _normalize_20_columns(
    row
):

    normalized = list(
        row[:20]
    )

    while len(
        normalized
    ) < 20:

        normalized.append("")

    return normalized


def _build_material_out_row(
    row
):

    return [

        row.get(
            "No",
            ""
        ),

        row.get(
            "No. DO",
            ""
        ),

        row.get(
            "Delv. Date",
            ""
        ),

        row.get(
            "Material Code",
            ""
        ),

        row.get(
            "Material Name",
            ""
        ),

        row.get(
            "Qty",
            0
        ),

        row.get(
            "UoM",
            row.get(
                "uom",
                "Pcs"
            )
        ),

        row.get(
            "Charging Type",
            ""
        ),

        row.get(
            "Site Alocation",
            row.get(
                "Site Allocation",
                ""
            )
        ),

        row.get(
            "Remarks",
            ""
        ),

        row.get(
            "To",
            ""
        ),

        row.get(
            "Phone No.",
            ""
        ),

        row.get(
            "Address",
            ""
        ),

        row.get(
            "EPC",
            ""
        ),

        row.get(
            "Date Reloc.",
            ""
        ),

        row.get(
            "No. DO Reloc.",
            ""
        ),

        row.get(
            "Qty Reloc.",
            ""
        ),

        row.get(
            "Site Reloc.",
            ""
        ),

        row.get(
            "Mitra Reloc.",
            ""
        ),

        row.get(
            "Remarks Reloc.",
            ""
        )

    ]


# ==============================================================================
# 10. CACHE INVALIDATION
# ==============================================================================

def _clear_function_cache(
    function_name
):

    try:

        function = globals().get(
            function_name
        )

        if (
            function
            and
            hasattr(
                function,
                "clear"
            )
        ):

            function.clear()

    except Exception:

        pass


def clear_reimbursement_cache():

    for name in [

        "_get_reimbursement_raw",

        "get_all_reimbursements"

    ]:

        _clear_function_cache(
            name
        )


def clear_material_out_cache():

    for name in [

        "_get_material_out_raw",

        "get_all_do_numbers",

        "get_do_by_number",

        "get_used_sites_from_db_material_out"

    ]:

        _clear_function_cache(
            name
        )


def clear_query_cache():

    for name in [

        "get_query_sheet_data",

        "get_epc_and_charging_dropdown_options",

        "get_sites_by_epc_and_charging"

    ]:

        _clear_function_cache(
            name
        )


def clear_authorization_cache():

    _clear_function_cache(
        "get_authorization_data"
    )


def clear_sequence_cache():

    _clear_function_cache(
        "_get_sequence_raw"
    )


def clear_database_cache(
    domain=None
):

    domain_clean = normalize_text(
        domain
    )

    if domain_clean == "reimbursement":

        clear_reimbursement_cache()

    elif domain_clean in [
        "material_out",
        "do",
        "material"
    ]:

        clear_material_out_cache()

    elif domain_clean == "query":

        clear_query_cache()

    elif domain_clean in [
        "authorization",
        "otorisasi",
        "auth"
    ]:

        clear_authorization_cache()

    elif domain_clean == "sequence":

        clear_sequence_cache()

    else:

        clear_reimbursement_cache()

        clear_material_out_cache()

        clear_query_cache()

        clear_authorization_cache()

        clear_sequence_cache()


# ==============================================================================
# 11. GOOGLE DRIVE IMAGE UPLOAD
# ==============================================================================

def upload_image_to_gdrive(
    file_bytes,
    file_name,
    pic_name,
    date_str
):

    print(
        "🔥 MENGIRIM FILE KE APPS SCRIPT WEB APP"
    )

    try:

        if not file_bytes:

            print(
                "❌ ERROR: file_bytes kosong."
            )

            return ""

        file_base64 = (
            base64
            .b64encode(
                file_bytes
            )
            .decode(
                "utf-8"
            )
        )

        cleaned_file_name = (
            str(
                file_name
            )
            .strip()
            .lower()
        )

        if not (
            cleaned_file_name.endswith(
                ".jpg"
            )
            or
            cleaned_file_name.endswith(
                ".jpeg"
            )
        ):

            cleaned_file_name += ".jpg"

        payload = {

            "fileName":
                cleaned_file_name,

            "fileData":
                file_base64,

            "picName":
                pic_name,

            "dateStr":
                date_str

        }

        response = requests.post(
            APPS_SCRIPT_WEB_APP_URL,
            json=payload,
            timeout=30
        )

        print(
            "DEBUG RESPONSE JSON:",
            response.text
        )

        if response.status_code != 200:

            st.error(
                "❌ Gagal HTTP POST ke Apps Script: "
                f"Status code {response.status_code}"
            )

            return ""

        try:

            res_json = response.json()

        except Exception:

            st.error(
                "❌ Response Apps Script bukan JSON valid."
            )

            return ""

        print(
            res_json
        )

        if (
            res_json.get(
                "status"
            )
            ==
            "success"
        ):

            file_id = (
                res_json.get(
                    "fileId",
                    ""
                )
            )

            if file_id:

                uploaded_data = (
                    "https://drive.google.com/thumbnail"
                    f"?id={file_id}&sz=w1200"
                )

                print(
                    "URL THUMBNAIL:",
                    uploaded_data
                )

                return uploaded_data

            return ""

        st.error(
            "❌ Gagal di Apps Script: "
            f"{res_json.get('message', 'Unknown error')}"
        )

        return ""

    except Exception as e:

        print(
            "❌ Error uploading image:",
            e
        )

        st.error(
            f"❌ Upload gambar gagal: {e}"
        )

        return ""


# ==============================================================================
# 12. SEQUENCE ENGINE
# ==============================================================================

def _ensure_sequence_sheet():

    try:

        return get_worksheet(
            SHEET_SEQUENCE
        )

    except WorksheetNotFound:

        pass

    sh = get_google_sheet_connection()

    with _write_lock:

        try:

            worksheet = sh.worksheet(
                SHEET_SEQUENCE
            )

            return worksheet

        except WorksheetNotFound:

            worksheet = sh.add_worksheet(
                title=SHEET_SEQUENCE,
                rows=1000,
                cols=4
            )

            worksheet.update(
                "A1:D1",
                [[
                    "Token",
                    "Prefix",
                    "Sequence",
                    "Created At"
                ]],
                value_input_option="USER_ENTERED"
            )

            try:

                _get_worksheet_cached.clear()

            except Exception:

                pass

            return worksheet


@st.cache_data(
    ttl=60,
    show_spinner=False
)
def _get_sequence_raw():

    try:

        def read_sequence():

            worksheet = _ensure_sequence_sheet()

            return worksheet.get_all_values()

        return safe_sheet_read(
            read_sequence
        )

    except Exception:

        return []


def _get_existing_max_document_number(
    source_sheet,
    document_column_index,
    prefix_suffix
):

    try:

        worksheet = get_worksheet(
            source_sheet
        )

        def read_existing():

            return worksheet.get_all_values()

        all_data = safe_sheet_read(
            read_existing
        )

        max_number = 0

        if not all_data:

            return 0

        for row in all_data[1:]:

            if len(row) <= document_column_index:

                continue

            no = str(
                row[
                    document_column_index
                ]
            ).strip()

            if not no:

                continue

            if prefix_suffix not in no:

                continue

            try:

                first_part = (
                    no.split(
                        "/"
                    )[0]
                )

                number = int(
                    first_part
                )

                if number > max_number:

                    max_number = number

            except Exception:

                continue

        return max_number

    except Exception:

        return 0


def _sequence_prefix_has_seed(
    sequence_rows,
    prefix_suffix
):

    if not sequence_rows:

        return False

    target = str(
        prefix_suffix
    ).strip()

    for row in sequence_rows[1:]:

        if len(row) < 2:

            continue

        token = str(
            row[0]
        ).strip()

        prefix = str(
            row[1]
        ).strip()

        if (
            token == "__SEED__"
            and
            prefix == target
        ):

            return True

    return False


def _ensure_sequence_seed(
    prefix_suffix,
    source_sheet,
    document_column_index
):

    sequence_rows = _get_sequence_raw()

    if _sequence_prefix_has_seed(
        sequence_rows,
        prefix_suffix
    ):

        return

    with _write_lock:

        try:

            _get_sequence_raw.clear()

        except Exception:

            pass

        sequence_rows = _get_sequence_raw()

        if _sequence_prefix_has_seed(
            sequence_rows,
            prefix_suffix
        ):

            return

        max_existing = (
            _get_existing_max_document_number(
                source_sheet,
                document_column_index,
                prefix_suffix
            )
        )

        worksheet = _ensure_sequence_sheet()

        seed_row = [

            "__SEED__",

            prefix_suffix,

            max_existing,

            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )

        ]

        worksheet.append_row(
            seed_row,
            value_input_option="USER_ENTERED"
        )

        try:

            _get_sequence_raw.clear()

        except Exception:

            pass


def reserve_document_number(
    prefix_suffix,
    source_sheet,
    document_column_index
):

    prefix_suffix = str(
        prefix_suffix
    ).strip()

    if not prefix_suffix:

        raise ValueError(
            "prefix_suffix tidak boleh kosong."
        )

    _ensure_sequence_seed(
        prefix_suffix,
        source_sheet,
        document_column_index
    )

    worksheet = _ensure_sequence_sheet()

    token = str(
        uuid.uuid4()
    )

    created_at = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    formula = (
        '=IF('
        'INDEX(B:B,ROW())="",'
        '"",'
        'IFERROR('
        'MAX('
        'FILTER('
        '$C$2:INDEX($C:$C,ROW()-1),'
        '$B$2:INDEX($B:$B,ROW()-1)'
        '='
        'INDEX(B:B,ROW())'
        ')'
        ')'
        '+1,'
        '1'
        ')'
        ')'
    )

    row = [

        token,

        prefix_suffix,

        formula,

        created_at

    ]

    response = worksheet.append_rows(

        [row],

        value_input_option="USER_ENTERED",

        table_range="A1:D1",

        include_values_in_response=True

    )

    sequence_value = None

    try:

        updates = (
            response.get(
                "updates",
                {}
            )
        )

        updated_data = (
            updates.get(
                "updatedData",
                {}
            )
        )

        values = (
            updated_data.get(
                "values",
                []
            )
        )

        if values:

            returned_row = values[0]

            if len(returned_row) >= 3:

                sequence_value = safe_int(
                    returned_row[2],
                    0
                )

    except Exception:

        sequence_value = None

    if (
        not sequence_value
        or
        sequence_value <= 0
    ):

        updated_range = (
            response
            .get(
                "updates",
                {}
            )
            .get(
                "updatedRange",
                ""
            )
        )

        row_number = (
            _parse_row_number_from_a1(
                updated_range
            )
        )

        if row_number:

            def read_sequence_cell():

                return worksheet.acell(
                    f"C{row_number}"
                ).value

            sequence_value = safe_int(
                safe_sheet_read(
                    read_sequence_cell
                ),
                0
            )

    if (
        not sequence_value
        or
        sequence_value <= 0
    ):

        raise RuntimeError(
            "Gagal mendapatkan sequence number "
            "dari DB Sequence."
        )

    return (
        f"{sequence_value:04d}"
        f"{prefix_suffix}"
    )


# ==============================================================================
# 13. REIMBURSEMENT
# ==============================================================================

def generate_reimbursement_no():

    now = datetime.now()

    roman_m = get_roman_month(
        now.month
    )

    year_str = now.strftime(
        "%Y"
    )

    prefix_suffix = (
        f"/CLX/RMS/{roman_m}/{year_str}"
    )

    try:

        return reserve_document_number(

            prefix_suffix=prefix_suffix,

            source_sheet=SHEET_REIMBURSEMENT,

            document_column_index=2

        )

    except Exception as e:

        st.warning(
            f"⚠️ Gagal generate nomor reimbursement: {e}"
        )

        try:

            max_existing = (
                _get_existing_max_document_number(
                    SHEET_REIMBURSEMENT,
                    2,
                    prefix_suffix
                )
            )

            return (
                f"{max_existing + 1:04d}"
                f"{prefix_suffix}"
            )

        except Exception:

            return (
                f"0001"
                f"{prefix_suffix}"
            )


# ==============================================================================
# SAVE REIMBURSEMENT
# ==============================================================================

def save_reimbursement_to_sheet(
    payload
):

    try:

        worksheet = get_worksheet(
            SHEET_REIMBURSEMENT
        )

        rows_to_append = []

        items = payload.get(
            "items",
            []
        )

        for item in items:

            item_evident_link = (
                item.get(
                    "evident",
                    ""
                )
            )

            row = [

                "",

                str(
                    payload.get(
                        "pic",
                        ""
                    )
                ),

                str(
                    payload.get(
                        "form_no",
                        ""
                    )
                ),

                str(
                    payload.get(
                        "date",
                        ""
                    )
                ),

                str(
                    item.get(
                        "description",
                        ""
                    )
                ),

                safe_int(
                    item.get(
                        "qty",
                        1
                    ),
                    1
                ),

                safe_float(
                    item.get(
                        "amount",
                        0
                    )
                ),

                safe_float(
                    item.get(
                        "total",
                        0
                    )
                ),

                str(
                    payload.get(
                        "remarks",
                        ""
                    )
                ),

                str(
                    payload.get(
                        "status_coo",
                        "Pending"
                    )
                ),

                str(
                    payload.get(
                        "status_cfo",
                        "Pending"
                    )
                ),

                (
                    item_evident_link
                    if item_evident_link
                    else ""
                )

            ]

            rows_to_append.append(
                row
            )

        if not rows_to_append:

            return False

        response = worksheet.append_rows(

            rows_to_append,

            value_input_option="USER_ENTERED",

            table_range="A1:L1"

        )

        updated_range = (
            response
            .get(
                "updates",
                {}
            )
            .get(
                "updatedRange",
                ""
            )
        )

        start_row = (
            _parse_row_number_from_a1(
                updated_range
            )
        )

        if start_row:

            no_values = [

                [start_row + index]

                for index in range(
                    len(
                        rows_to_append
                    )
                )

            ]

            end_row = (
                start_row
                +
                len(rows_to_append)
                -
                1
            )

            worksheet.update(
                f"A{start_row}:A{end_row}",
                no_values,
                value_input_option="USER_ENTERED"
            )

        clear_reimbursement_cache()

        return True

    except Exception as e:

        st.error(
            "❌ Gagal menyimpan data Reimbursement: "
            f"{type(e).__name__} - {str(e)}"
        )

        return False


# ==============================================================================
# RAW REIMBURSEMENT DATA
# ==============================================================================

@st.cache_data(
    ttl=600,
    show_spinner=False
)
def _get_reimbursement_raw():

    try:

        def read_reimbursement():

            worksheet = get_worksheet(
                SHEET_REIMBURSEMENT
            )

            return worksheet.get_all_values()

        return safe_sheet_read(
            read_reimbursement
        )

    except Exception as e:

        if is_rate_limit_error(
            e
        ):

            st.warning(
                "⚠️ Google Sheets API quota tercapai "
                "saat membaca Reimbursement."
            )

        else:

            st.error(
                "❌ Gagal membaca DB Reimbursement: "
                f"{e}"
            )

        return []


# ==============================================================================
# GET ALL REIMBURSEMENTS
# ==============================================================================

@st.cache_data(
    ttl=600,
    show_spinner=False
)
def get_all_reimbursements():

    try:

        all_data = (
            _get_reimbursement_raw()
        )

        if (
            not all_data
            or
            len(all_data) < 2
        ):

            return []

        grouped = {}

        for row in all_data[1:]:

            if len(row) < 8:

                continue

            form_no = str(
                row[2]
            ).strip()

            if not form_no:

                continue

            pic = (
                row[1]
                if len(row) > 1
                else ""
            )

            date_str = (
                row[3]
                if len(row) > 3
                else ""
            )

            desc = (
                row[4]
                if len(row) > 4
                else ""
            )

            qty = safe_int(
                row[5],
                1
            )

            amt = safe_float(
                row[6]
            )

            tot = safe_float(
                row[7]
            )

            remarks = (
                row[8]
                if len(row) > 8
                else ""
            )

            status_coo = (
                row[9]
                if len(row) > 9
                and row[9]
                else "Pending"
            )

            status_cfo = (
                row[10]
                if len(row) > 10
                and row[10]
                else "Pending"
            )

            raw_evident = (
                row[11]
                if len(row) > 11
                else ""
            )

            clean_evident = str(
                raw_evident
            ).strip()

            if clean_evident.lower() in [

                "",

                "0",

                "0.0",

                "none",

                "null"

            ]:

                clean_evident = ""

            if (
                status_coo
                ==
                "Rejected"
                or
                status_cfo
                ==
                "Rejected"
            ):

                overall_status = "Rejected"

            elif (
                status_coo
                ==
                "Approved"
                and
                status_cfo
                ==
                "Approved"
            ):

                overall_status = "Approved"

            elif (
                status_coo
                ==
                "Approved"
            ):

                overall_status = "Pending CFO"

            else:

                overall_status = "Pending COO"

            if form_no not in grouped:

                grouped[form_no] = {

                    "form_no":
                        form_no,

                    "pic":
                        pic,

                    "date":
                        date_str,

                    "remarks":
                        remarks,

                    "status_coo":
                        status_coo,

                    "status_cfo":
                        status_cfo,

                    "status":
                        overall_status,

                    "items":
                        [],

                    "grand_total":
                        0.0,

                    "image_links":
                        []

                }

            grouped[
                form_no
            ][
                "items"
            ].append({

                "no":
                    len(
                        grouped[
                            form_no
                        ][
                            "items"
                        ]
                    )
                    +
                    1,

                "description":
                    desc,

                "qty":
                    qty,

                "amount":
                    amt,

                "total":
                    tot,

                "evident":
                    clean_evident

            })

            if (
                clean_evident
                and
                clean_evident.startswith(
                    "http"
                )
                and
                clean_evident
                not in grouped[
                    form_no
                ][
                    "image_links"
                ]
            ):

                grouped[
                    form_no
                ][
                    "image_links"
                ].append(
                    clean_evident
                )

            grouped[
                form_no
            ][
                "grand_total"
            ] += tot

        return list(
            grouped.values()
        )

    except Exception as e:

        st.error(
            "❌ Gagal mengambil data Reimbursement: "
            f"{e}"
        )

        return []


# ==============================================================================
# UPDATE REIMBURSEMENT STATUS
# ==============================================================================

def update_reimbursement_status(
    form_no_target,
    new_status,
    approver_role,
    note=""
):

    try:

        worksheet = get_worksheet(
            SHEET_REIMBURSEMENT
        )

        all_rows = (
            _get_reimbursement_raw()
        )

        if (
            not all_rows
            or
            len(all_rows) < 2
        ):

            return False

        if (
            normalize_text(
                approver_role
            )
            ==
            "coo"
        ):

            col_index = 10

        else:

            col_index = 11

        if (
            "Approved"
            in
            str(
                new_status
            )
            or
            new_status
            ==
            "Pending CFO"
        ):

            status_value = "Approved"

        else:

            status_value = "Rejected"

        cells_to_update = []

        for row_idx, row in enumerate(
            all_rows[1:],
            start=2
        ):

            if (
                len(row) >= 3
                and
                str(
                    row[2]
                ).strip()
                ==
                str(
                    form_no_target
                ).strip()
            ):

                cells_to_update.append(
                    gspread.Cell(
                        row_idx,
                        col_index,
                        status_value
                    )
                )

        if not cells_to_update:

            return False

        worksheet.update_cells(
            cells_to_update,
            value_input_option="USER_ENTERED"
        )

        clear_reimbursement_cache()

        return True

    except Exception as e:

        st.error(
            "❌ Gagal memperbarui status Approval: "
            f"{e}"
        )

        return False


# ==============================================================================
# 14. DELIVERY ORDER / MATERIAL OUT
# ==============================================================================

def generate_do_number(
    is_reloc=False
):

    now = datetime.now()

    month_roman = get_roman_month(
        now.month
    )

    year_str = now.strftime(
        "%Y"
    )

    code_type = (
        "DO-RELOC"
        if is_reloc
        else
        "DO"
    )

    prefix_suffix = (
        f"/CLX/{code_type}/{month_roman}/{year_str}"
    )

    try:

        return reserve_document_number(

            prefix_suffix=prefix_suffix,

            source_sheet=SHEET_MATERIAL_OUT,

            document_column_index=1

        )

    except Exception as e:

        st.warning(
            f"⚠️ Gagal generate nomor DO: {e}"
        )

        try:

            max_existing = (
                _get_existing_max_document_number(
                    SHEET_MATERIAL_OUT,
                    1,
                    prefix_suffix
                )
            )

            return (
                f"{max_existing + 1:04d}"
                f"{prefix_suffix}"
            )

        except Exception:

            return (
                f"0001"
                f"{prefix_suffix}"
            )


# ==============================================================================
# SAVE DO
# ==============================================================================

def save_do_to_db_material_out(
    rows_data
):

    try:

        worksheet = get_worksheet(
            SHEET_MATERIAL_OUT
        )

        records = []

        for row in rows_data:

            records.append(
                _build_material_out_row(
                    row
                )
            )

        if not records:

            return False

        worksheet.append_rows(

            records,

            value_input_option="USER_ENTERED",

            table_range="A1:T1"

        )

        clear_material_out_cache()

        return True

    except Exception as e:

        st.error(
            "❌ Gagal menyimpan data DO: "
            f"{type(e).__name__} - {str(e)}"
        )

        return False


# ==============================================================================
# RAW MATERIAL OUT DATA
# ==============================================================================

@st.cache_data(
    ttl=600,
    show_spinner=False
)
def _get_material_out_raw():

    try:

        def read_material_out():

            worksheet = get_worksheet(
                SHEET_MATERIAL_OUT
            )

            return worksheet.get_all_values()

        return safe_sheet_read(
            read_material_out
        )

    except Exception as e:

        if is_rate_limit_error(
            e
        ):

            st.warning(
                "⚠️ Google Sheets API quota tercapai "
                "saat membaca Material Out."
            )

        else:

            st.warning(
                f"⚠️ Gagal membaca Material Out: {e}"
            )

        return []


# ==============================================================================
# GET ALL DO NUMBERS
# ==============================================================================

@st.cache_data(
    ttl=600,
    show_spinner=False
)
def get_all_do_numbers():

    try:

        all_data = (
            _get_material_out_raw()
        )

        if (
            not all_data
            or
            len(all_data) < 2
        ):

            return []

        unique_dos = set()

        for row in all_data[1:]:

            if len(row) > 1:

                no_do = str(
                    row[1]
                ).strip()

                if no_do:

                    unique_dos.add(
                        no_do
                    )

        return sorted(
            list(
                unique_dos
            )
        )

    except Exception as e:

        if is_rate_limit_error(
            e
        ):

            st.warning(
                "⚠️ Google Sheets API quota tercapai."
            )

        else:

            st.warning(
                f"⚠️ Gagal mengambil nomor DO: {e}"
            )

        return []


# ==============================================================================
# GET DO BY NUMBER
# ==============================================================================

@st.cache_data(
    ttl=300,
    show_spinner=False
)
def get_do_by_number(
    no_do_search
):

    try:

        all_data = (
            _get_material_out_raw()
        )

        if (
            not all_data
            or
            len(all_data) < 2
        ):

            return None

        matching_rows = []

        target = str(
            no_do_search
        ).strip()

        for row in all_data[1:]:

            if (
                len(row) >= 2
                and
                str(
                    row[1]
                ).strip()
                ==
                target
            ):

                matching_rows.append({

                    "No":
                        row[0]
                        if len(row) > 0
                        else "",

                    "No. DO":
                        row[1]
                        if len(row) > 1
                        else "",

                    "Delv. Date":
                        row[2]
                        if len(row) > 2
                        else "",

                    "Material Code":
                        row[3]
                        if len(row) > 3
                        else "",

                    "Material Name":
                        row[4]
                        if len(row) > 4
                        else "",

                    "Qty":
                        safe_int(
                            row[5]
                            if len(row) > 5
                            else 0
                        ),

                    "UoM":
                        row[6]
                        if len(row) > 6
                        else "Pcs",

                    "uom":
                        row[6]
                        if len(row) > 6
                        else "Pcs",

                    "Charging Type":
                        row[7]
                        if len(row) > 7
                        else "",

                    "Site Alocation":
                        row[8]
                        if len(row) > 8
                        else "",

                    "Site Allocation":
                        row[8]
                        if len(row) > 8
                        else "",

                    "Remarks":
                        row[9]
                        if len(row) > 9
                        else "",

                    "To":
                        row[10]
                        if len(row) > 10
                        else "",

                    "Phone No.":
                        row[11]
                        if len(row) > 11
                        else "",

                    "Address":
                        row[12]
                        if len(row) > 12
                        else "",

                    "EPC":
                        row[13]
                        if len(row) > 13
                        else "",

                    "Date Reloc.":
                        row[14]
                        if len(row) > 14
                        else "",

                    "No. DO Reloc.":
                        row[15]
                        if len(row) > 15
                        else "",

                    "Qty Reloc.":
                        row[16]
                        if len(row) > 16
                        else "",

                    "Site Reloc.":
                        row[17]
                        if len(row) > 17
                        else "",

                    "Mitra Reloc.":
                        row[18]
                        if len(row) > 18
                        else "",

                    "Remarks Reloc.":
                        row[19]
                        if len(row) > 19
                        else ""

                })

        if not matching_rows:

            return None

        first = matching_rows[0]

        return {

            "no_do":
                first[
                    "No. DO"
                ],

            "date":
                first[
                    "Delv. Date"
                ],

            "to":
                first[
                    "To"
                ],

            "contact":
                first[
                    "Phone No."
                ],

            "address":
                first[
                    "Address"
                ],

            "epc":
                first[
                    "EPC"
                ],

            "charging_type":
                first[
                    "Charging Type"
                ],

            "materials":
                matching_rows

        }

    except Exception as e:

        st.error(
            f"❌ Error mengambil data DO: {e}"
        )

        return None


# ==============================================================================
# UPDATE DO - TARGETED UPDATE
# ==============================================================================

def update_do_in_db_material_out(
    no_do_target,
    updated_rows_data
):

    try:

        worksheet = get_worksheet(
            SHEET_MATERIAL_OUT
        )

        all_rows = (
            _get_material_out_raw()
        )

        if (
            not all_rows
            or
            len(all_rows) < 2
        ):

            return False

        target = str(
            no_do_target
        ).strip()

        target_sheet_rows = []

        for python_index, row in enumerate(
            all_rows[1:],
            start=1
        ):

            if (
                len(row) > 1
                and
                str(
                    row[1]
                ).strip()
                ==
                target
            ):

                target_sheet_rows.append(
                    python_index + 1
                )

        if not target_sheet_rows:

            records = []

            for row in updated_rows_data:

                records.append(
                    _build_material_out_row(
                        row
                    )
                )

            if not records:

                return False

            worksheet.append_rows(

                records,

                value_input_option="USER_ENTERED",

                table_range="A1:T1"

            )

            clear_material_out_cache()

            return True

        new_records = [

            _build_material_out_row(
                row
            )

            for row in updated_rows_data

        ]

        old_count = len(
            target_sheet_rows
        )

        new_count = len(
            new_records
        )

        common_count = min(
            old_count,
            new_count
        )

        if common_count > 0:

            groups = []

            current_group = [
                target_sheet_rows[0]
            ]

            for row_number in target_sheet_rows[1:]:

                if (
                    row_number
                    ==
                    current_group[-1] + 1
                ):

                    current_group.append(
                        row_number
                    )

                else:

                    groups.append(
                        current_group
                    )

                    current_group = [
                        row_number
                    ]

            groups.append(
                current_group
            )

            record_cursor = 0

            for group in groups:

                group_length = min(
                    len(group),
                    common_count - record_cursor
                )

                if group_length <= 0:

                    break

                first_row = group[0]

                last_row = (
                    first_row
                    +
                    group_length
                    -
                    1
                )

                values = new_records[
                    record_cursor:
                    record_cursor
                    +
                    group_length
                ]

                worksheet.update(

                    f"A{first_row}:T{last_row}",

                    values,

                    value_input_option="USER_ENTERED"

                )

                record_cursor += (
                    group_length
                )

        if new_count > old_count:

            additional_records = (
                new_records[
                    old_count:
                ]
            )

            worksheet.append_rows(

                additional_records,

                value_input_option="USER_ENTERED",

                table_range="A1:T1"

            )

        elif old_count > new_count:

            rows_to_clear = (
                target_sheet_rows[
                    new_count:
                ]
            )

            clear_groups = []

            if rows_to_clear:

                current_group = [
                    rows_to_clear[0]
                ]

                for row_number in rows_to_clear[1:]:

                    if (
                        row_number
                        ==
                        current_group[-1] + 1
                    ):

                        current_group.append(
                            row_number
                        )

                    else:

                        clear_groups.append(
                            current_group
                        )

                        current_group = [
                            row_number
                        ]

                clear_groups.append(
                    current_group
                )

            clear_ranges = []

            for group in clear_groups:

                first_row = group[0]

                last_row = group[-1]

                clear_ranges.append(
                    f"A{first_row}:T{last_row}"
                )

            if clear_ranges:

                worksheet.batch_clear(
                    clear_ranges
                )

        clear_material_out_cache()

        return True

    except Exception as e:

        st.error(
            "❌ Gagal memperbarui data DO: "
            f"{type(e).__name__} - {str(e)}"
        )

        return False


# ==============================================================================
# 15. QUERY SHEET
# ==============================================================================

@st.cache_data(
    ttl=1800,
    show_spinner=False
)
def get_query_sheet_data():

    try:

        def read_query():

            worksheet = get_worksheet(
                SHEET_QUERY
            )

            return worksheet.get_all_values()

        all_values = safe_sheet_read(
            read_query
        )

        if not all_values:

            return pd.DataFrame()

        if len(all_values) <= 1:

            return pd.DataFrame()

        headers = list(
            all_values[0]
        )

        rows = all_values[1:]

        normalized_rows = []

        for row in rows:

            row = list(row)

            if len(row) < len(headers):

                row += (
                    [""] *
                    (
                        len(headers)
                        -
                        len(row)
                    )
                )

            elif len(row) > len(headers):

                row = row[
                    :len(headers)
                ]

            normalized_rows.append(
                row
            )

        return pd.DataFrame(
            normalized_rows,
            columns=headers
        )

    except Exception as e:

        if is_rate_limit_error(
            e
        ):

            st.warning(
                "⚠️ Google Sheets API sedang mencapai quota "
                "pada Sheet Query."
            )

        else:

            st.warning(
                f"⚠️ Gagal membaca Sheet Query: {e}"
            )

        return pd.DataFrame()


# ==============================================================================
# GET USED SITES
# ==============================================================================

@st.cache_data(
    ttl=600,
    show_spinner=False
)
def get_used_sites_from_db_material_out():

    try:

        all_rows = (
            _get_material_out_raw()
        )

        if (
            not all_rows
            or
            len(all_rows) < 2
        ):

            return []

        used_sites = set()

        for row in all_rows[1:]:

            if len(row) > 8:

                site = str(
                    row[8]
                ).strip()

                if site:

                    used_sites.add(
                        site
                    )

        return sorted(
            list(
                used_sites
            )
        )

    except Exception as e:

        if is_rate_limit_error(
            e
        ):

            st.warning(
                "⚠️ Google Sheets API quota tercapai "
                "saat membaca Used Sites."
            )

        else:

            st.warning(
                f"⚠️ Gagal mengambil Used Sites: {e}"
            )

        return []


# ==============================================================================
# EPC & CHARGING TYPE DROPDOWN
# ==============================================================================

@st.cache_data(
    ttl=1800,
    show_spinner=False
)
def get_epc_and_charging_dropdown_options():

    try:

        df = get_query_sheet_data()

        if df.empty:

            return [], []

        epc_col = (
            df.columns[1]
            if len(df.columns) > 1
            else None
        )

        charging_col = (
            df.columns[2]
            if len(df.columns) > 2
            else None
        )

        if epc_col is None:

            epc_list = []

        else:

            epc_list = sorted(

                [

                    x

                    for x in (

                        df[
                            epc_col
                        ]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .unique()

                    )

                    if x

                ]

            )

        if charging_col is None:

            charging_list = []

        else:

            charging_list = sorted(

                [

                    x

                    for x in (

                        df[
                            charging_col
                        ]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .unique()

                    )

                    if x

                ]

            )

        return (
            epc_list,
            charging_list
        )

    except Exception as e:

        st.error(
            "❌ Gagal mengambil opsi EPC & Charging: "
            f"{e}"
        )

        return [], []


# ==============================================================================
# GET SITES BY EPC & CHARGING
# ==============================================================================

@st.cache_data(
    ttl=1800,
    show_spinner=False
)
def get_sites_by_epc_and_charging(
    epc_target,
    charging_type_target
):

    try:

        df = get_query_sheet_data()

        if (
            df.empty
            or
            len(df.columns) <= 5
        ):

            return []

        epc_clean = normalize_text(
            epc_target
        )

        charging_clean = normalize_text(
            charging_type_target
        )

        col_epc = df.columns[1]

        col_charging = df.columns[2]

        col_status = df.columns[3]

        col_site = df.columns[5]

        mask = (

            (
                df[
                    col_epc
                ]
                .astype(str)
                .str.strip()
                .str.lower()
                ==
                epc_clean
            )

            &

            (
                df[
                    col_charging
                ]
                .astype(str)
                .str.strip()
                .str.lower()
                ==
                charging_clean
            )

            &

            (
                ~df[
                    col_status
                ]
                .astype(str)
                .str.strip()
                .str.lower()
                .str.contains(
                    "cancel|drop",
                    regex=True,
                    na=False
                )
            )

        )

        filtered_df = df[
            mask
        ]

        sites = (

            filtered_df[
                col_site
            ]

            .dropna()

            .astype(str)

            .str.strip()

            .unique()

            .tolist()

        )

        return [

            site

            for site in sites

            if site

        ]

    except Exception as e:

        st.error(
            "❌ Gagal mengambil data Site dari Query: "
            f"{e}"
        )

        return []


# ==============================================================================
# 16. AUTHORIZATION / OTORISASI
# ==============================================================================

@st.cache_data(
    ttl=300,
    show_spinner=False
)
def get_authorization_data():

    try:

        def read_authorization():

            worksheet = get_worksheet(
                SHEET_AUTHORIZATION
            )

            return worksheet.get_all_values()

        all_values = safe_sheet_read(
            read_authorization
        )

        if not all_values:

            return pd.DataFrame()

        if len(all_values) < 2:

            return pd.DataFrame()

        headers = list(
            all_values[0]
        )

        rows = all_values[1:]

        normalized_rows = []

        for row in rows:

            row = list(row)

            if len(row) < len(headers):

                row += (
                    [""] *
                    (
                        len(headers)
                        -
                        len(row)
                    )
                )

            elif len(row) > len(headers):

                row = row[
                    :len(headers)
                ]

            normalized_rows.append(
                row
            )

        return pd.DataFrame(
            normalized_rows,
            columns=headers
        )

    except Exception as e:

        if is_rate_limit_error(
            e
        ):

            st.warning(
                "⚠️ Google Sheets API quota tercapai "
                "saat membaca Sheet Otorisasi. "
                "Sistem akan mencoba kembali setelah cooldown."
            )

        else:

            st.error(
                "❌ Gagal menghubungkan ke Sheet Otorisasi: "
                f"{e}"
            )

        return pd.DataFrame()


# ==============================================================================
# GET AUTHORIZED USER
# ==============================================================================

def get_authorized_user(
    username=None,
    email=None
):

    try:

        df = get_authorization_data()

        if df.empty:

            return None

        username_columns = [

            "username",
            "user",
            "userid",
            "user id",
            "login",
            "nik"

        ]

        email_columns = [

            "email",
            "e-mail",
            "email address"

        ]

        username_col = None

        email_col = None

        normalized_columns = {

            normalize_text(
                col
            ):
                col

            for col in df.columns

        }

        for candidate in username_columns:

            if candidate in normalized_columns:

                username_col = (
                    normalized_columns[
                        candidate
                    ]
                )

                break

        for candidate in email_columns:

            if candidate in normalized_columns:

                email_col = (
                    normalized_columns[
                        candidate
                    ]
                )

                break

        if (
            username
            and
            username_col
        ):

            target = normalize_text(
                username
            )

            result = df[

                df[
                    username_col
                ]
                .astype(str)
                .str.strip()
                .str.lower()
                ==
                target

            ]

            if not result.empty:

                return (
                    result.iloc[
                        0
                    ]
                    .to_dict()
                )

        if (
            email
            and
            email_col
        ):

            target = normalize_text(
                email
            )

            result = df[

                df[
                    email_col
                ]
                .astype(str)
                .str.strip()
                .str.lower()
                ==
                target

            ]

            if not result.empty:

                return (
                    result.iloc[
                        0
                    ]
                    .to_dict()
                )

        return None

    except Exception as e:

        st.error(
            f"❌ Gagal mencari user authorization: {e}"
        )

        return None


# ==============================================================================
# 17. DATABASE HEALTH CHECK
# ==============================================================================

def google_sheet_health_check():

    try:

        def read_test():

            sh = get_google_sheet_connection()

            return sh.worksheets()

        worksheets = safe_sheet_read(
            read_test
        )

        return {

            "status":
                "OK",

            "message":
                "Google Sheets connection OK",

            "worksheets":
                len(
                    worksheets
                )

        }

    except Exception as e:

        return {

            "status":
                "ERROR",

            "message":
                str(e),

            "worksheets":
                0

        }


# ==============================================================================
# 18. DEBUG API RATE LIMIT
# ==============================================================================

def get_current_read_usage():

    with _read_lock:

        now = time.time()

        while (

            _read_request_times

            and

            (
                now
                -
                _read_request_times[0]
            )
            >=
            RATE_WINDOW_SECONDS

        ):

            _read_request_times.popleft()

        return len(
            _read_request_times
        )


# ==============================================================================
# 18A. DEBUG RATE LIMIT COOLDOWN
# ==============================================================================

def get_rate_limit_cooldown():

    with _read_lock:

        remaining = (
            _next_allowed_read_time
            -
            time.time()
        )

        return max(
            remaining,
            0
        )


# ==============================================================================
# 19. DATABASE CACHE STATUS
# ==============================================================================

def get_database_cache_status():

    return {

        "read_limit_per_minute":
            MAX_READ_REQUESTS_PER_MINUTE,

        "rate_window_seconds":
            RATE_WINDOW_SECONDS,

        "current_read_usage":
            get_current_read_usage(),

        "rate_limit_cooldown_seconds":
            RATE_LIMIT_COOLDOWN_SECONDS,

        "current_cooldown_remaining":
            round(
                get_rate_limit_cooldown(),
                2
            ),

        "cache_domains": [

            "reimbursement",
            "material_out",
            "query",
            "authorization",
            "sequence"

        ],

        "sequence_sheet":
            SHEET_SEQUENCE,

        "compatibility_helpers": [

            "get_sheet_values",
            "get_sheet_dataframe"

        ]

    }


# ==============================================================================
# 20. OPTIONAL DATABASE WARMUP
# ==============================================================================

def warmup_database(
    include_reimbursement=False,
    include_material_out=False,
    include_query=True,
    include_authorization=False
):

    result = {

        "query":
            False,

        "authorization":
            False,

        "material_out":
            False,

        "reimbursement":
            False

    }

    try:

        if include_query:

            get_query_sheet_data()

            result[
                "query"
            ] = True

        if include_authorization:

            get_authorization_data()

            result[
                "authorization"
            ] = True

        if include_material_out:

            _get_material_out_raw()

            result[
                "material_out"
            ] = True

        if include_reimbursement:

            _get_reimbursement_raw()

            result[
                "reimbursement"
            ] = True

    except Exception as e:

        result[
            "error"
        ] = str(e)

    return result


# ==============================================================================
# END OF DATABASE.PY V2.2
# ==============================================================================
