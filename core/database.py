# ==============================================================================
# DATABASE.PY
# ERP DATABASE LAYER - GOOGLE SHEETS
#
# VERSION:
# Rate Limit Protected / Cached / Retry / Backoff
#
# FEATURES:
# 1. Streamlit Resource Cache
# 2. Streamlit Data Cache
# 3. Internal Google Sheets Read Rate Limiter
# 4. Exponential Backoff + Jitter
# 5. Centralized Safe Google Sheets Read
# 6. Reduced API Read Requests
# 7. Reimbursement Database
# 8. Material Out / Delivery Order Database
# 9. Query Database
# 10. Authorization Database
# 11. Google Drive Image Upload via Apps Script
# ==============================================================================


# ==============================================================================
# 0. IMPORT
# ==============================================================================

import os
import io
import time
import base64
import random
import threading

from collections import deque
from datetime import datetime

import requests
import gspread
import pandas as pd
import streamlit as st

from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError


# ==============================================================================
# 1. GLOBAL CONFIGURATION
# ==============================================================================

SPREADSHEET_ID = "1FU1lL3ls3jP_hAxBdx_Fu35Z9Ap4ICdHmOpMvCyA3gY"

# ------------------------------------------------------------------------------
# DATABASE SHEETS
# ------------------------------------------------------------------------------

SHEET_REIMBURSEMENT = "DB Reimbursement"
SHEET_MATERIAL_OUT = "DB Material Out"
SHEET_QUERY = "Query"

# Sheet authorization.
# Jika nama tab kamu berbeda, cukup ubah value ini.
SHEET_AUTHORIZATION = "Otorisasi"

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
# 2. GOOGLE SHEETS RATE LIMIT PROTECTION
# ==============================================================================

# Google limit yang muncul pada error kamu:
#
# ReadRequestsPerMinutePerUser = 60
#
# Kita sengaja menggunakan 45 request/minute sebagai safety margin.
# Ini BUKAN menaikkan quota Google.
# Ini membatasi aplikasi agar tidak menembak API terlalu agresif.

MAX_READ_REQUESTS_PER_MINUTE = 45

RATE_WINDOW_SECONDS = 60

_read_request_times = deque()

_read_lock = threading.Lock()


def wait_for_read_slot():
    """
    Internal rate limiter untuk READ Google Sheets.

    Maksimal:
        45 READ request / 60 detik

    Jika limit internal tercapai, aplikasi akan menunggu
    sampai slot tersedia.
    """

    while True:

        now = time.time()

        with _read_lock:

            # Hapus timestamp yang sudah lebih dari 60 detik
            while (
                _read_request_times
                and (
                    now - _read_request_times[0]
                    >= RATE_WINDOW_SECONDS
                )
            ):
                _read_request_times.popleft()

            # Masih tersedia slot
            if len(_read_request_times) < MAX_READ_REQUESTS_PER_MINUTE:

                _read_request_times.append(now)

                return

            # Hitung waktu tunggu
            wait_time = (
                RATE_WINDOW_SECONDS
                - (
                    now
                    - _read_request_times[0]
                )
                + 0.5
            )

        time.sleep(
            max(wait_time, 0.5)
        )


# ==============================================================================
# 3. RATE LIMIT ERROR DETECTOR
# ==============================================================================

def is_rate_limit_error(error):
    """
    Mendeteksi berbagai bentuk error quota/rate limit Google API.
    """

    error_text = str(error).upper()

    keywords = [
        "429",
        "RESOURCE_EXHAUSTED",
        "RATE_LIMIT_EXCEEDED",
        "QUOTA_EXCEEDED",
        "TOO MANY REQUESTS",
        "READREQUESTSPERMINUTE"
    ]

    return any(
        keyword in error_text
        for keyword in keywords
    )


# ==============================================================================
# 4. SAFE GOOGLE SHEETS READ
# ==============================================================================

def safe_sheet_read(
    read_function,
    max_retries=5
):
    """
    Centralized Google Sheets READ handler.

    Semua READ Google Sheets sebaiknya melewati fungsi ini.

    Proteksi:
        1. Internal rate limiter
        2. 429 detection
        3. Exponential backoff
        4. Random jitter
        5. Retry
    """

    last_error = None

    for attempt in range(max_retries):

        try:

            # --------------------------------------------------------------
            # STEP 1
            # Internal Rate Limiter
            # --------------------------------------------------------------

            wait_for_read_slot()

            # --------------------------------------------------------------
            # STEP 2
            # Execute API Read
            # --------------------------------------------------------------

            result = read_function()

            return result

        except Exception as e:

            last_error = e

            # --------------------------------------------------------------
            # Jika bukan rate limit, jangan retry
            # --------------------------------------------------------------

            if not is_rate_limit_error(e):

                raise

            # --------------------------------------------------------------
            # STEP 3
            # Exponential Backoff
            # --------------------------------------------------------------

            base_delay = min(
                2 ** attempt,
                32
            )

            # --------------------------------------------------------------
            # STEP 4
            # Random Jitter
            # --------------------------------------------------------------

            jitter = random.uniform(
                0.5,
                1.5
            )

            sleep_time = (
                base_delay
                * jitter
            )

            print(
                f"⚠️ Google Sheets RATE LIMIT. "
                f"Retry {attempt + 1}/{max_retries} "
                f"dalam {sleep_time:.1f} detik..."
            )

            time.sleep(
                sleep_time
            )

    # Semua retry gagal
    raise last_error


# ==============================================================================
# 5. GOOGLE SHEET CONNECTION
# ==============================================================================

@st.cache_resource(
    show_spinner=False
)
def get_google_sheet_connection():
    """
    Membuka koneksi utama Google Spreadsheet.

    Connection di-cache sehingga Streamlit tidak membuat
    koneksi baru pada setiap rerun.
    """

    try:

        # ------------------------------------------------------------------
        # Streamlit Secrets
        # ------------------------------------------------------------------

        if "gcp_service_account" in st.secrets:

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

        # ------------------------------------------------------------------
        # Local service_account.json
        # ------------------------------------------------------------------

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

        # ------------------------------------------------------------------
        # Authorize
        # ------------------------------------------------------------------

        gc = gspread.authorize(
            creds
        )

        # ------------------------------------------------------------------
        # Open Spreadsheet
        # ------------------------------------------------------------------

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
# 6. GENERAL HELPERS
# ==============================================================================

def get_roman_month(month_int):
    """
    Mengubah angka bulan menjadi angka Romawi.
    """

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


def safe_int(value, default=0):
    """
    Konversi value menjadi integer secara aman.
    """

    try:

        if value is None:
            return default

        text = str(value).strip()

        if not text:
            return default

        return int(
            float(
                text.replace(",", "")
            )
        )

    except Exception:

        return default


def safe_float(value, default=0.0):
    """
    Konversi value menjadi float secara aman.
    """

    try:

        if value is None:
            return default

        text = (
            str(value)
            .replace(",", "")
            .replace("Rp", "")
            .replace("rp", "")
            .strip()
        )

        if not text:
            return default

        return float(text)

    except Exception:

        return default


def normalize_text(value):
    """
    Normalisasi text untuk pencocokan.
    """

    if value is None:
        return ""

    return (
        str(value)
        .strip()
        .lower()
    )


# ==============================================================================
# 7. CACHE INVALIDATION
# ==============================================================================

def clear_database_cache():
    """
    Membersihkan cache database setelah INSERT/UPDATE.

    Aman dipanggil setelah write operation.
    """

    try:
        st.cache_data.clear()
    except Exception:
        pass


# ==============================================================================
# 8. GOOGLE DRIVE IMAGE UPLOAD
# ==============================================================================

def upload_image_to_gdrive(
    file_bytes,
    file_name,
    pic_name,
    date_str
):
    """
    Upload gambar ke Google Drive melalui Apps Script Web App.
    """

    print(
        "🔥 MENGIRIM FILE KE APPS SCRIPT WEB APP"
    )

    try:

        if not file_bytes:

            print(
                "❌ ERROR: file_bytes kosong."
            )

            return ""

        # ------------------------------------------------------------------
        # Convert bytes -> Base64
        # ------------------------------------------------------------------

        file_base64 = (
            base64
            .b64encode(file_bytes)
            .decode("utf-8")
        )

        # ------------------------------------------------------------------
        # Normalize file name
        # ------------------------------------------------------------------

        cleaned_file_name = (
            str(file_name)
            .strip()
            .lower()
        )

        if not (
            cleaned_file_name.endswith(".jpg")
            or
            cleaned_file_name.endswith(".jpeg")
        ):

            cleaned_file_name += ".jpg"

        # ------------------------------------------------------------------
        # Payload
        # ------------------------------------------------------------------

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

        # ------------------------------------------------------------------
        # POST
        # ------------------------------------------------------------------

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

        # ------------------------------------------------------------------
        # Success
        # ------------------------------------------------------------------

        if res_json.get(
            "status"
        ) == "success":

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

        # ------------------------------------------------------------------
        # Apps Script Error
        # ------------------------------------------------------------------

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
# 9. REIMBURSEMENT
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

        def read_data():

            sh = get_google_sheet_connection()

            worksheet = sh.worksheet(
                SHEET_REIMBURSEMENT
            )

            return worksheet.get_all_values()

        all_data = safe_sheet_read(
            read_data
        )

        existing_numbers = []

        if len(all_data) > 1:

            for row in all_data[1:]:

                if len(row) <= 2:
                    continue

                no = str(
                    row[2]
                ).strip()

                if prefix_suffix in no:

                    try:

                        num_part = int(
                            no.split("/")[0]
                        )

                        existing_numbers.append(
                            num_part
                        )

                    except ValueError:

                        continue

        next_num = (
            max(existing_numbers) + 1
            if existing_numbers
            else 1
        )

        return (
            f"{next_num:04d}"
            f"{prefix_suffix}"
        )

    except Exception as e:

        st.warning(
            f"⚠️ Gagal generate nomor reimbursement: {e}"
        )

        return (
            f"0001"
            f"{prefix_suffix}"
        )


def save_reimbursement_to_sheet(
    payload
):

    try:

        sh = get_google_sheet_connection()

        worksheet = sh.worksheet(
            SHEET_REIMBURSEMENT
        )

        # --------------------------------------------------------------
        # Tidak perlu READ get_all_values()
        #
        # Sebelumnya:
        # next_no = len(all_values)
        #
        # Sekarang kita gunakan timestamp-based row sequence.
        # Nomor database internal tidak digunakan sebagai nomor dokumen.
        # --------------------------------------------------------------

        rows_to_append = []

        items = payload.get(
            "items",
            []
        )

        # Ambil row number hanya sekali menggunakan worksheet row_count
        # Tidak melakukan READ API tambahan.

        next_no = (
            worksheet.row_count
            + 1
        )

        for item in items:

            item_evident_link = (
                item.get(
                    "evident",
                    ""
                )
            )

            row = [

                str(next_no),

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

            next_no += 1

        if not rows_to_append:
            return False

        # --------------------------------------------------------------
        # WRITE
        # --------------------------------------------------------------

        worksheet.append_rows(
            rows_to_append,
            value_input_option="USER_ENTERED"
        )

        clear_database_cache()

        return True

    except Exception as e:

        st.error(
            "❌ Gagal menyimpan data Reimbursement: "
            f"{type(e).__name__} - {str(e)}"
        )

        return False


# ==============================================================================
# GET ALL REIMBURSEMENTS
# ==============================================================================

@st.cache_data(
    ttl=600,
    show_spinner=False
)
def get_all_reimbursements():

    try:

        def read_reimbursement():

            sh = get_google_sheet_connection()

            worksheet = sh.worksheet(
                SHEET_REIMBURSEMENT
            )

            return worksheet.get_all_values()

        all_data = safe_sheet_read(
            read_reimbursement
        )

        if not all_data or len(all_data) < 2:

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

            # ----------------------------------------------------------
            # Overall status
            # ----------------------------------------------------------

            if (
                status_coo == "Rejected"
                or
                status_cfo == "Rejected"
            ):

                overall_status = "Rejected"

            elif (
                status_coo == "Approved"
                and
                status_cfo == "Approved"
            ):

                overall_status = "Approved"

            elif status_coo == "Approved":

                overall_status = "Pending CFO"

            else:

                overall_status = "Pending COO"

            # ----------------------------------------------------------
            # Create parent form
            # ----------------------------------------------------------

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

            # ----------------------------------------------------------
            # Add item
            # ----------------------------------------------------------

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
                    ) + 1,

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

            # ----------------------------------------------------------
            # Image links
            # ----------------------------------------------------------

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

            # ----------------------------------------------------------
            # Grand total
            # ----------------------------------------------------------

            grouped[
                form_no
            ][
                "grand_total"
            ] += tot

        return list(
            grouped.values()
        )

    except Exception as e:

        if is_rate_limit_error(e):

            st.warning(
                "⚠️ Google Sheets API quota tercapai "
                "saat membaca Reimbursement."
            )

        else:

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

        sh = get_google_sheet_connection()

        worksheet = sh.worksheet(
            SHEET_REIMBURSEMENT
        )

        def read_data():

            return worksheet.get_all_values()

        all_rows = safe_sheet_read(
            read_data
        )

        if not all_rows or len(all_rows) < 2:
            return False

        # COO = column J = 10
        # CFO = column K = 11

        if approver_role == "coo":

            col_index = 10

        else:

            col_index = 11

        if (
            "Approved" in str(new_status)
            or
            new_status == "Pending CFO"
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
                str(row[2]).strip()
                == str(form_no_target).strip()
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

        clear_database_cache()

        return True

    except Exception as e:

        st.error(
            "❌ Gagal memperbarui status Approval: "
            f"{e}"
        )

        return False


# ==============================================================================
# 10. DELIVERY ORDER / MATERIAL OUT
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
        else "DO"
    )

    prefix_suffix = (
        f"/CLX/{code_type}/{month_roman}/{year_str}"
    )

    try:

        def read_data():

            sh = get_google_sheet_connection()

            worksheet = sh.worksheet(
                SHEET_MATERIAL_OUT
            )

            return worksheet.get_all_values()

        all_data = safe_sheet_read(
            read_data
        )

        existing_numbers = []

        if len(all_data) > 1:

            for row in all_data[1:]:

                if len(row) <= 1:
                    continue

                no = str(
                    row[1]
                ).strip()

                if prefix_suffix in no:

                    try:

                        num_part = int(
                            no.split("/")[0]
                        )

                        existing_numbers.append(
                            num_part
                        )

                    except ValueError:

                        continue

        next_num = (
            max(existing_numbers) + 1
            if existing_numbers
            else 1
        )

        return (
            f"{next_num:04d}"
            f"{prefix_suffix}"
        )

    except Exception as e:

        st.warning(
            f"⚠️ Gagal generate nomor DO: {e}"
        )

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

        sh = get_google_sheet_connection()

        worksheet = sh.worksheet(
            SHEET_MATERIAL_OUT
        )

        records = []

        for row in rows_data:

            records.append([

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
            ])

        if not records:
            return False

        worksheet.append_rows(
            records,
            value_input_option="USER_ENTERED"
        )

        clear_database_cache()

        return True

    except Exception as e:

        st.error(
            "❌ Gagal menyimpan data DO: "
            f"{e}"
        )

        return False


# ==============================================================================
# GET ALL DO NUMBERS
# ==============================================================================

@st.cache_data(
    ttl=600,
    show_spinner=False
)
def get_all_do_numbers():

    try:

        def read_data():

            sh = get_google_sheet_connection()

            worksheet = sh.worksheet(
                SHEET_MATERIAL_OUT
            )

            return worksheet.get_all_values()

        all_data = safe_sheet_read(
            read_data
        )

        if not all_data or len(all_data) < 2:
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
            list(unique_dos)
        )

    except Exception as e:

        if is_rate_limit_error(e):

            st.warning(
                "⚠️ Google Sheets API quota tercapai."
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

        def read_data():

            sh = get_google_sheet_connection()

            worksheet = sh.worksheet(
                SHEET_MATERIAL_OUT
            )

            return worksheet.get_all_values()

        all_data = safe_sheet_read(
            read_data
        )

        if not all_data or len(all_data) < 2:
            return None

        matching_rows = []

        target = str(
            no_do_search
        ).strip()

        for row in all_data[1:]:

            if (
                len(row) >= 2
                and
                str(row[1]).strip()
                == target
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
                first["No. DO"],

            "date":
                first["Delv. Date"],

            "to":
                first["To"],

            "contact":
                first["Phone No."],

            "address":
                first["Address"],

            "epc":
                first["EPC"],

            "charging_type":
                first["Charging Type"],

            "materials":
                matching_rows
        }

    except Exception as e:

        st.error(
            f"❌ Error mengambil data DO: {e}"
        )

        return None


# ==============================================================================
# UPDATE DO
# ==============================================================================

def update_do_in_db_material_out(
    no_do_target,
    updated_rows_data
):

    try:

        sh = get_google_sheet_connection()

        worksheet = sh.worksheet(
            SHEET_MATERIAL_OUT
        )

        # --------------------------------------------------------------
        # Read existing data
        # --------------------------------------------------------------

        def read_data():

            return worksheet.get_all_values()

        all_rows = safe_sheet_read(
            read_data
        )

        if not all_rows or len(all_rows) < 2:
            return False

        header = all_rows[0]

        kept_rows = [
            header
        ]

        target = str(
            no_do_target
        ).strip()

        # --------------------------------------------------------------
        # Keep all DO except target
        # --------------------------------------------------------------

        for row in all_rows[1:]:

            if (
                len(row) > 1
                and
                str(row[1]).strip()
                != target
            ):

                # Pastikan selalu 20 kolom
                normalized = (
                    list(row[:20])
                )

                while len(
                    normalized
                ) < 20:

                    normalized.append("")

                kept_rows.append(
                    normalized
                )

        # --------------------------------------------------------------
        # Add updated rows
        # --------------------------------------------------------------

        for row in updated_rows_data:

            kept_rows.append([

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
            ])

        # --------------------------------------------------------------
        # WRITE
        # --------------------------------------------------------------

        worksheet.clear()

        worksheet.update(
            "A1",
            kept_rows,
            value_input_option="USER_ENTERED"
        )

        clear_database_cache()

        return True

    except Exception as e:

        st.error(
            "❌ Gagal memperbarui data DO: "
            f"{e}"
        )

        return False


# ==============================================================================
# 11. QUERY SHEET
# ==============================================================================

@st.cache_data(
    ttl=1800,
    show_spinner=False
)
def get_query_sheet_data():

    try:

        def read_query():

            sh = get_google_sheet_connection()

            worksheet = sh.worksheet(
                SHEET_QUERY
            )

            # SATU READ REQUEST
            return worksheet.get_all_values()

        all_values = safe_sheet_read(
            read_query
        )

        if not all_values:
            return pd.DataFrame()

        if len(all_values) <= 1:
            return pd.DataFrame()

        headers = all_values[0]

        rows = all_values[1:]

        # --------------------------------------------------------------
        # Normalize column count
        # --------------------------------------------------------------

        normalized_rows = []

        for row in rows:

            row = list(row)

            if len(row) < len(headers):

                row += [
                    ""
                ] * (
                    len(headers)
                    - len(row)
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

        if is_rate_limit_error(e):

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

        def read_material_out():

            sh = get_google_sheet_connection()

            worksheet = sh.worksheet(
                SHEET_MATERIAL_OUT
            )

            return worksheet.get_all_values()

        all_rows = safe_sheet_read(
            read_material_out
        )

        if not all_rows or len(all_rows) < 2:
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
            list(used_sites)
        )

    except Exception as e:

        if is_rate_limit_error(e):

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

        # Sesuai struktur Query lama:
        #
        # Column B = EPC
        # Column C = Charging Type

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

        if not epc_col:

            epc_list = []

        else:

            epc_list = sorted(
                [
                    x
                    for x in
                    df[
                        epc_col
                    ]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .unique()
                    if x
                ]
            )

        if not charging_col:

            charging_list = []

        else:

            charging_list = sorted(
                [
                    x
                    for x in
                    df[
                        charging_col
                    ]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .unique()
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

        # --------------------------------------------------------------
        # Struktur Query lama:
        #
        # B = EPC
        # C = Charging Type
        # D = Status
        # F = Site
        # --------------------------------------------------------------

        col_epc = df.columns[1]

        col_charging = df.columns[2]

        col_status = df.columns[3]

        col_site = df.columns[5]

        # --------------------------------------------------------------
        # Filter
        # --------------------------------------------------------------

        mask = (

            (
                df[col_epc]
                .astype(str)
                .str.strip()
                .str.lower()
                ==
                epc_clean
            )

            &

            (
                df[col_charging]
                .astype(str)
                .str.strip()
                .str.lower()
                ==
                charging_clean
            )

            &

            (
                ~df[col_status]
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
# 12. AUTHORIZATION / OTORISASI
# ==============================================================================

@st.cache_data(
    ttl=300,
    show_spinner=False
)
def get_authorization_data():

    """
    Membaca Sheet Otorisasi.

    CACHE:
        5 menit

    Tujuan:
        Mencegah login/authentication melakukan READ Google Sheets
        pada setiap Streamlit rerun.

    Struktur kolom mengikuti header yang ada di Sheet Otorisasi.
    """

    try:

        def read_authorization():

            sh = get_google_sheet_connection()

            worksheet = sh.worksheet(
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

        headers = all_values[0]

        rows = all_values[1:]

        normalized_rows = []

        for row in rows:

            row = list(row)

            if len(row) < len(headers):

                row += [
                    ""
                ] * (
                    len(headers)
                    - len(row)
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

        if is_rate_limit_error(e):

            st.warning(
                "⚠️ Google Sheets API quota tercapai "
                "saat membaca Sheet Otorisasi. "
                "Silakan tunggu beberapa detik."
            )

        else:

            st.error(
                "❌ Gagal menghubungkan ke Sheet Otorisasi: "
                f"{e}"
            )

        return pd.DataFrame()


# ==============================================================================
# OPTIONAL AUTHORIZATION HELPER
# ==============================================================================

def get_authorized_user(
    username=None,
    email=None
):

    """
    Helper untuk mencari user dari Sheet Otorisasi.

    Bisa dipanggil berdasarkan username atau email.

    Karena struktur kolom Otorisasi bisa berbeda,
    fungsi ini mencoba mencari nama kolom umum.
    """

    try:

        df = get_authorization_data()

        if df.empty:
            return None

        # --------------------------------------------------------------
        # Cari kolom username
        # --------------------------------------------------------------

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
            normalize_text(col): col
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

        # --------------------------------------------------------------
        # Username search
        # --------------------------------------------------------------

        if username and username_col:

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

                return result.iloc[0].to_dict()

        # --------------------------------------------------------------
        # Email search
        # --------------------------------------------------------------

        if email and email_col:

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

                return result.iloc[0].to_dict()

        return None

    except Exception as e:

        st.error(
            f"❌ Gagal mencari user authorization: {e}"
        )

        return None


# ==============================================================================
# 13. DATABASE HEALTH CHECK
# ==============================================================================

def google_sheet_health_check():

    """
    Fungsi sederhana untuk mengecek koneksi Google Sheets.

    Hanya melakukan satu READ request.
    """

    try:

        def read_test():

            sh = get_google_sheet_connection()

            # Mengambil worksheet list.
            # Ini digunakan hanya sebagai test koneksi.
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
                len(worksheets)
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
# 14. DEBUG API RATE LIMIT
# ==============================================================================

def get_current_read_usage():

    """
    Debug helper.

    Mengembalikan jumlah READ request yang tercatat
    oleh rate limiter internal dalam window 60 detik.
    """

    with _read_lock:

        now = time.time()

        while (
            _read_request_times
            and
            now
            - _read_request_times[0]
            >= RATE_WINDOW_SECONDS
        ):

            _read_request_times.popleft()

        return len(
            _read_request_times
        )


# ==============================================================================
# END OF DATABASE.PY
# ==============================================================================
