import os
import io
import time
import base64
import requests
import gspread
import pandas as pd
from datetime import datetime
import streamlit as st

# Google Service Account API Imports
from google.oauth2.service_account import Credentials

# ==============================================================================
# KONFIGURASI GLOBAL DATABASE ERP & SCOPES
# ==============================================================================
SPREADSHEET_ID = "1FU1lL3ls3jP_hAxBdx_Fu35Z9Ap4ICdHmOpMvCyA3gY"
SHEET_REIMBURSEMENT = "DB Reimbursement"
SHEET_MATERIAL_OUT = "DB Material Out"
SHEET_QUERY = "Query"
GOOGLE_DRIVE_ROOT_FOLDER_ID = "1fto5kD7X_pYT21F6Qr1RLfmBSmEb1O3o"

# URL Apps Script Web App untuk Upload Gambar ke Google Drive
APPS_SCRIPT_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwAx3pGoDtMLI7CZV58WoNSeKo2oHx3jCs8IARlAagUvaAVRAWkoLeZ1H_4P0RMpD6p/exec"

# Scope yang dibutuhkan untuk Google Sheets & Drive
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]


# ==============================================================================
# 1. KONEKSI UTAMA SERVICE ACCOUNT (KHUSUS GOOGLE SHEETS)
# ==============================================================================

@st.cache_resource
def get_google_sheet_connection():
    """Mengembalikan koneksi gspread Spreadsheet utama menggunakan Service Account Credentials."""
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        elif os.path.exists("service_account.json"):
            creds = Credentials.from_service_account_file("service_account.json", scopes=SCOPES)
        else:
            st.error("❌ Konfigurasi Service Account tidak ditemukan di `st.secrets['gcp_service_account']` maupun file `service_account.json`.")
            st.stop()

        gc = gspread.authorize(creds)
        sh = gc.open_by_key(SPREADSHEET_ID)
        return sh
        
    except Exception as e:
        st.error(f"❌ Gagal terhubung ke Google Sheets via Service Account: {type(e).__name__} - {str(e)}")
        st.stop()


def get_roman_month(month_int):
    """Utility helper mengubah angka bulan ke format Romawi"""
    roman_months = {
        1: "I", 2: "II", 3: "III", 4: "IV", 5: "V", 6: "VI",
        7: "VII", 8: "VIII", 9: "IX", 10: "X", 11: "XI", 12: "XII"
    }
    return roman_months.get(month_int, "I")


# ==============================================================================
# 2. HELPER UPLOAD GAMBAR KE GOOGLE DRIVE (VIA APPS SCRIPT WEB APP)
# ==============================================================================

def upload_image_to_gdrive(file_bytes, file_name, pic_name, date_str):
    """Mengunggah gambar ke Google Drive melalui HTTP POST ke Apps Script Web App."""
    print("🔥 MENGIRIM FILE KE APPS SCRIPT WEB APP")
    try:
        if not file_bytes:
            print("❌ ERROR: file_bytes kosong atau bernilai None!")
            return ""

        file_base64 = base64.b64encode(file_bytes).decode('utf-8')

        cleaned_file_name = file_name.lower()
        if not cleaned_file_name.endswith(".jpg") and not cleaned_file_name.endswith(".jpeg"):
            cleaned_file_name += ".jpg"

        payload = {
            "fileName": cleaned_file_name,
            "fileData": file_base64,
            "picName": pic_name,
            "dateStr": date_str
        }

        response = requests.post(APPS_SCRIPT_WEB_APP_URL, json=payload, timeout=30)
        print("DEBUG RESPONSE JSON:", response.text)
        
        if response.status_code == 200:
            res_json = response.json()
            print(res_json)
            
            if res_json.get("status") == "success":
                file_id = res_json.get("fileId", "")
                if file_id:
                    uploaded_data = f"https://drive.google.com/thumbnail?id={file_id}&sz=w1200"
                    print("URL THUMBNAIL:", uploaded_data)
                    return uploaded_data
                return ""
            else:
                st.error(f"Gagal di Apps Script: {res_json.get('message')}")
                return ""
        else:
            st.error(f"Gagal HTTP POST: Status code {response.status_code}")
            return ""

    except Exception as e:
        print(f"Error uploading image via Apps Script Web App: {e}")
        return ""


# ==============================================================================
# 3. HELPER REIMBURSEMENT
# ==============================================================================

def generate_reimbursement_no():
    now = datetime.now()
    roman_m = get_roman_month(now.month)
    year_str = now.strftime("%Y")
    prefix_suffix = f"/CLX/RMS/{roman_m}/{year_str}"
    
    try:
        sh = get_google_sheet_connection()
        worksheet = sh.worksheet(SHEET_REIMBURSEMENT)
        form_nos = worksheet.col_values(3)
        
        existing_numbers = []
        for no in form_nos[1:]:
            if prefix_suffix in str(no):
                try:
                    num_part = int(str(no).split('/')[0])
                    existing_numbers.append(num_part)
                except ValueError:
                    continue
        
        next_num = max(existing_numbers) + 1 if existing_numbers else 1
        return f"{next_num:04d}{prefix_suffix}"
        
    except Exception:
        return f"0001{prefix_suffix}"


def save_reimbursement_to_sheet(payload):
    try:
        sh = get_google_sheet_connection()
        
        try:
            worksheet = sh.worksheet(SHEET_REIMBURSEMENT)
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"❌ Tab Sheet '{SHEET_REIMBURSEMENT}' tidak ditemukan di Google Sheets!")
            return False
        
        all_values = worksheet.get_all_values()
        next_no = len(all_values)

        items = payload.get("items", [])
        rows_to_append = []
        
        for item in items:
            item_evident_link = item.get("evident", "")
            print("EVIDENT LINK YANG MASUK SHEET:", item_evident_link)

            row = [
                str(next_no),
                str(payload.get("pic", "")),
                str(payload.get("form_no", "")),
                str(payload.get("date", "")),
                str(item.get("description", "")),
                int(item.get("qty", 1)),
                float(item.get("amount", 0)),
                float(item.get("total", 0)),
                str(payload.get("remarks", "")),
                str(payload.get("status_coo", "Pending")),
                str(payload.get("status_cfo", "Pending")),
                item_evident_link if item_evident_link else ""
            ]
            rows_to_append.append(row)
            next_no += 1
        
        if rows_to_append:
            worksheet.append_rows(rows_to_append)
            st.cache_data.clear()  # Clear cache setelah ada penulisan data baru
            return True
        return False

    except Exception as e:
        st.error(f"❌ Gagal menyimpan data Reimbursement: {type(e).__name__} - {str(e)}")
        return False


@st.cache_data(ttl=300)
def get_all_reimbursements():
    max_retries = 3
    retry_delay = 3
    for attempt in range(max_retries):
        try:
            sh = get_google_sheet_connection()
            worksheet = sh.worksheet(SHEET_REIMBURSEMENT)
            all_data = worksheet.get_all_values()

            if not all_data or len(all_data) < 2:
                return []

            grouped = {}
            for row in all_data[1:]:
                if len(row) < 8:
                    continue

                form_no = row[2]
                if not form_no:
                    continue

                pic = row[1]
                date_str = row[3]
                desc = row[4]
                qty = int(row[5]) if str(row[5]).isdigit() else 1
                amt = float(str(row[6]).replace(",", "").replace("Rp", "").strip() or 0)
                tot = float(str(row[7]).replace(",", "").replace("Rp", "").strip() or 0)
                remarks = row[8] if len(row) > 8 else ""
                status_coo = row[9] if len(row) > 9 and row[9] else "Pending"
                status_cfo = row[10] if len(row) > 10 and row[10] else "Pending"

                raw_evident = row[11] if len(row) > 11 else ""
                clean_evident = str(raw_evident).strip()
                if clean_evident in ["0", "0.0", "None", "none"]:
                    clean_evident = ""

                if status_coo == "Rejected" or status_cfo == "Rejected":
                    overall_status = "Rejected"
                elif status_coo == "Approved" and status_cfo == "Approved":
                    overall_status = "Approved"
                elif status_coo == "Approved":
                    overall_status = "Pending CFO"
                else:
                    overall_status = "Pending COO"

                if form_no not in grouped:
                    grouped[form_no] = {
                        "form_no": form_no,
                        "pic": pic,
                        "date": date_str,
                        "remarks": remarks,
                        "status_coo": status_coo,
                        "status_cfo": status_cfo,
                        "status": overall_status,
                        "items": [],
                        "grand_total": 0.0,
                        "image_links": []
                    }

                grouped[form_no]["items"].append({
                    "no": len(grouped[form_no]["items"]) + 1,
                    "description": desc,
                    "qty": qty,
                    "amount": amt,
                    "total": tot,
                    "evident": clean_evident
                })

                if clean_evident and clean_evident.startswith("http") and clean_evident not in grouped[form_no]["image_links"]:
                    grouped[form_no]["image_links"].append(clean_evident)
                    
                grouped[form_no]["grand_total"] += tot

            return list(grouped.values())

        except Exception as e:
            if ("429" in str(e) or "RESOURCE_EXHAUSTED" in str(e)) and attempt < max_retries - 1:
                time.sleep(retry_delay * (attempt + 1))
                continue
            st.error(f"❌ Gagal mengambil data Reimbursement: {e}")
            return []


def update_reimbursement_status(form_no_target, new_status, approver_role, note=""):
    try:
        sh = get_google_sheet_connection()
        worksheet = sh.worksheet(SHEET_REIMBURSEMENT)
        all_rows = worksheet.get_all_values()

        if not all_rows or len(all_rows) < 2:
            return False

        col_index = 10 if approver_role == "coo" else 11
        status_value = "Approved" if "Approved" in new_status or new_status == "Pending CFO" else "Rejected"

        cells_to_update = []
        for row_idx, row in enumerate(all_rows[1:], start=2):
            if len(row) >= 3 and row[2] == form_no_target:
                cells_to_update.append(gspread.Cell(row_idx, col_index, status_value))

        if cells_to_update:
            worksheet.update_cells(cells_to_update)
            st.cache_data.clear()  # Clear cache setelah update
            return True
        return False

    except Exception as e:
        st.error(f"❌ Gagal memperbarui status Approval di Sheet: {e}")
        return False


# ==============================================================================
# 4. HELPER DELIVERY ORDER (SCM - DB MATERIAL OUT)
# ==============================================================================

def generate_do_number(is_reloc=False):
    """Membuat Nomor DO Otomatis (Support DO Reguler dan DO Relokasi)"""
    now = datetime.now()
    month_roman = get_roman_month(now.month)
    year_str = now.strftime("%Y")
    
    code_type = "DO-RELOC" if is_reloc else "DO"
    prefix_suffix = f"/CLX/{code_type}/{month_roman}/{year_str}"

    try:
        sh = get_google_sheet_connection()
        worksheet = sh.worksheet(SHEET_MATERIAL_OUT)
        do_nos = worksheet.col_values(2)  # Kolom B: No. DO
        
        existing_numbers = []
        for no in do_nos[1:]:
            if prefix_suffix in str(no):
                try:
                    num_part = int(str(no).split('/')[0])
                    existing_numbers.append(num_part)
                except ValueError:
                    continue

        next_num = max(existing_numbers) + 1 if existing_numbers else 1
        return f"{next_num:04d}{prefix_suffix}"

    except Exception:
        return f"0001{prefix_suffix}"


def save_do_to_db_material_out(rows_data):
    """Menyimpan atau Menambahkan Baris DO Baru ke Sheet DB Material Out (Kolom A:T)"""
    try:
        sh = get_google_sheet_connection()
        worksheet = sh.worksheet(SHEET_MATERIAL_OUT)

        records = []
        for row in rows_data:
            records.append([
                row.get("No", ""),
                row.get("No. DO", ""),
                row.get("Delv. Date", ""),
                row.get("Material Code", ""),
                row.get("Material Name", ""),
                row.get("Qty", 0),
                row.get("UoM", row.get("uom", "Pcs")),
                row.get("Charging Type", ""),
                row.get("Site Alocation", row.get("Site Allocation", "")),
                row.get("Remarks", ""),
                row.get("To", ""),
                row.get("Phone No.", ""),
                row.get("Address", ""),
                row.get("EPC", ""),
                row.get("Date Reloc.", ""),
                row.get("No. DO Reloc.", ""),
                row.get("Qty Reloc.", ""),
                row.get("Site Reloc.", ""),
                row.get("Mitra Reloc.", ""),
                row.get("Remarks Reloc.", "")
            ])

        if records:
            worksheet.append_rows(records)
            st.cache_data.clear()  # Clear cache setelah data baru ditambahkan
            return True
        return False

    except Exception as e:
        st.error(f"❌ Gagal menyimpan data DO ke DB Material Out: {e}")
        return False


@st.cache_data(ttl=300)
def get_all_do_numbers():
    """Mengambil list unik seluruh Nomor DO yang sudah tersimpan"""
    try:
        sh = get_google_sheet_connection()
        worksheet = sh.worksheet(SHEET_MATERIAL_OUT)
        do_nos = worksheet.col_values(2)
        
        unique_dos = sorted(list(set([str(no).strip() for no in do_nos[1:] if str(no).strip()])))
        return unique_dos
    except Exception:
        return []


def get_do_by_number(no_do_search):
    """Mencari detail DO berdasarkan Nomor DO (Mendukung Struktur Kolom A:T)"""
    try:
        sh = get_google_sheet_connection()
        worksheet = sh.worksheet(SHEET_MATERIAL_OUT)
        all_data = worksheet.get_all_values()

        if not all_data or len(all_data) < 2:
            return None

        matching_rows = []
        for row in all_data[1:]:
            if len(row) >= 2 and row[1] == no_do_search:
                matching_rows.append({
                    "No": row[0] if len(row) > 0 else "",
                    "No. DO": row[1] if len(row) > 1 else "",
                    "Delv. Date": row[2] if len(row) > 2 else "",
                    "Material Code": row[3] if len(row) > 3 else "",
                    "Material Name": row[4] if len(row) > 4 else "",
                    "Qty": int(row[5]) if len(row) > 5 and str(row[5]).isdigit() else (row[5] if len(row) > 5 else 0),
                    "UoM": row[6] if len(row) > 6 else "Pcs",
                    "uom": row[6] if len(row) > 6 else "Pcs",
                    "Charging Type": row[7] if len(row) > 7 else "",
                    "Site Alocation": row[8] if len(row) > 8 else "",
                    "Site Allocation": row[8] if len(row) > 8 else "",
                    "Remarks": row[9] if len(row) > 9 else "",
                    "To": row[10] if len(row) > 10 else "",
                    "Phone No.": row[11] if len(row) > 11 else "",
                    "Address": row[12] if len(row) > 12 else "",
                    "EPC": row[13] if len(row) > 13 else "",
                    "Date Reloc.": row[14] if len(row) > 14 else "",
                    "No. DO Reloc.": row[15] if len(row) > 15 else "",
                    "Qty Reloc.": row[16] if len(row) > 16 else "",
                    "Site Reloc.": row[17] if len(row) > 17 else "",
                    "Mitra Reloc.": row[18] if len(row) > 18 else "",
                    "Remarks Reloc.": row[19] if len(row) > 19 else ""
                })

        if not matching_rows:
            return None

        first = matching_rows[0]
        return {
            "no_do": first["No. DO"],
            "date": first["Delv. Date"],
            "to": first["To"],
            "contact": first["Phone No."],
            "address": first["Address"],
            "epc": first["EPC"],
            "charging_type": first["Charging Type"],
            "materials": matching_rows
        }

    except Exception as e:
        st.error(f"❌ Error saat mengambil data DO: {e}")
        return None


def update_do_in_db_material_out(no_do_target, updated_rows_data):
    """Memperbarui baris data DO di Sheet DB Material Out (Presisi Kolom A:T)"""
    try:
        sh = get_google_sheet_connection()
        worksheet = sh.worksheet(SHEET_MATERIAL_OUT)
        all_rows = worksheet.get_all_values()

        if not all_rows or len(all_rows) < 2:
            return False

        header = all_rows[0]
        kept_rows = [header]

        for row in all_rows[1:]:
            if len(row) > 1 and row[1] != no_do_target:
                kept_rows.append(row)

        for row in updated_rows_data:
            kept_rows.append([
                row.get("No", ""),
                row.get("No. DO", ""),
                row.get("Delv. Date", ""),
                row.get("Material Code", ""),
                row.get("Material Name", ""),
                row.get("Qty", 0),
                row.get("UoM", row.get("uom", "Pcs")),
                row.get("Charging Type", ""),
                row.get("Site Alocation", row.get("Site Allocation", "")),
                row.get("Remarks", ""),
                row.get("To", ""),
                row.get("Phone No.", ""),
                row.get("Address", ""),
                row.get("EPC", ""),
                row.get("Date Reloc.", ""),
                row.get("No. DO Reloc.", ""),
                row.get("Qty Reloc.", ""),
                row.get("Site Reloc.", ""),
                row.get("Mitra Reloc.", ""),
                row.get("Remarks Reloc.", "")
            ])

        worksheet.clear()
        worksheet.update('A1', kept_rows)
        st.cache_data.clear()  # Clear cache setelah update
        return True

    except Exception as e:
        st.error(f"❌ Gagal memperbarui data DO di Google Sheet: {e}")
        return False


# ==============================================================================
# 5. HELPER QUERY SHEET (DILINDUNGI CACHING & RETRY PROTEKSI ERROR 429)
# ==============================================================================

@st.cache_data(ttl=1800)  # Menjadikan cache bertahan 30 menit untuk menekan request API
def get_query_sheet_data():
    """
    Mengambil data Sheet 'Query' secara utuh sebagai DataFrame dengan Proteksi Backoff 429 Rate Limit.
    """
    max_retries = 3
    retry_delay = 3  # Jeda awal 3 detik jika terkena kuota limit
    
    for attempt in range(max_retries):
        try:
            sh = get_google_sheet_connection()
            worksheet = sh.worksheet(SHEET_QUERY)
            data = worksheet.get_all_records()
            
            if data:
                return pd.DataFrame(data)
            
            all_values = worksheet.get_all_values()
            if len(all_values) > 1:
                headers = all_values[0]
                rows = all_values[1:]
                return pd.DataFrame(rows, columns=headers)

            return pd.DataFrame()

        except Exception as e:
            # Jika terkena Google Rate Limit (HTTP 429 / RESOURCE_EXHAUSTED)
            if "429" in str(e) or "RESOURCE_EXHAUSTED" in str(e) or "RATE_LIMIT_EXCEEDED" in str(e):
                if attempt < max_retries - 1:
                    time.sleep(retry_delay * (attempt + 1))  # Menunggu 3s, 6s, dst.
                    continue
            
            st.warning(f"⚠️ Batas kuota pembacaan Google API tercapai pada Sheet '{SHEET_QUERY}'. Menggunakan data fallback/kosong sementara.")
            return pd.DataFrame()


@st.cache_data(ttl=300)
def get_used_sites_from_db_material_out():
    """
    Membaca Sheet 'DB Material Out' untuk mendapatkan daftar site
    yang sudah terpakai (Kolom I / Index 8) dengan proteksi rate limit.
    """
    max_retries = 3
    retry_delay = 3
    
    for attempt in range(max_retries):
        try:
            sh = get_google_sheet_connection()
            worksheet = sh.worksheet(SHEET_MATERIAL_OUT)
            all_rows = worksheet.get_all_values()

            if not all_rows or len(all_rows) < 2:
                return []

            used_sites = set()
            for row in all_rows[1:]:
                if len(row) > 8 and str(row[8]).strip():
                    used_sites.add(str(row[8]).strip())

            return list(used_sites)

        except Exception as e:
            if ("429" in str(e) or "RESOURCE_EXHAUSTED" in str(e)) and attempt < max_retries - 1:
                time.sleep(retry_delay * (attempt + 1))
                continue
            st.warning(f"⚠️ Gagal mengambil used sites dari '{SHEET_MATERIAL_OUT}': {e}")
            return []


def get_epc_and_charging_dropdown_options():
    """Mengambil list unik EPC & Charging Type dari cache DataFrame Query"""
    try:
        df = get_query_sheet_data()
        if df.empty:
            return [], []

        epc_col = df.columns[1] if len(df.columns) > 1 else None
        charging_col = df.columns[2] if len(df.columns) > 2 else None

        epc_list = sorted(list(df[epc_col].dropna().astype(str).str.strip().unique())) if epc_col else []
        charging_list = sorted(list(df[charging_col].dropna().astype(str).str.strip().unique())) if charging_col else []

        return [x for x in epc_list if x], [x for x in charging_list if x]

    except Exception as e:
        st.error(f"❌ Gagal mengambil opsi EPC & Charging: {e}")
        return [], []


def get_sites_by_epc_and_charging(epc_target, charging_type_target):
    """
    Membaca DataFrame 'Query' yang di-cache dan memfilter daftar site
    berdasarkan pencocokan EPC dan Charging Type tanpa memanggil Google API lagi.
    """
    try:
        df = get_query_sheet_data()
        if df.empty or len(df.columns) <= 5:
            return []

        epc_clean = str(epc_target).strip().lower() if epc_target else ""
        charging_clean = str(charging_type_target).strip().lower() if charging_type_target else ""

        col_epc = df.columns[1]
        col_charging = df.columns[2]
        col_status = df.columns[3]
        col_site = df.columns[5]

        # Filter data tanpa hit Google API
        mask = (
            (df[col_epc].astype(str).str.strip().str.lower() == epc_clean) &
            (df[col_charging].astype(str).str.strip().str.lower() == charging_clean) &
            (~df[col_status].astype(str).str.strip().str.lower().str.contains("cancel|drop", regex=True, na=False))
        )
        
        filtered_df = df[mask]
        sites = filtered_df[col_site].dropna().astype(str).str.strip().unique().tolist()
        return [s for s in sites if s]

    except Exception as e:
        st.error(f"❌ Gagal mengambil data site dari Sheet 'Query': {e}")
        return []
