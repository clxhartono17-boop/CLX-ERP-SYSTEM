import os
import io
import base64
import traceback
import requests
import gspread
from datetime import datetime
import streamlit as st
from PIL import Image

# Google Service Account API Imports
from google.oauth2.service_account import Credentials

# ==============================================================================
# KONFIGURASI GLOBAL DATABASE ERP & SCOPES
# ==============================================================================
SPREADSHEET_ID = "1FU1lL3ls3jP_hAxBdx_Fu35Z9Ap4ICdHmOpMvCyA3gY"
SHEET_REIMBURSEMENT = "DB Reimbursement"
SHEET_MATERIAL_OUT = "DB Material Out"
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
        # Mengambil data dictionary dari st.secrets['gcp_service_account']
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
    """
    Mengunggah gambar ke Google Drive melalui HTTP POST ke Apps Script Web App.
    """
    print("🔥 MENGIRIM FILE KE APPS SCRIPT WEB APP")
    try:
        if not file_bytes:
            print("❌ ERROR: file_bytes kosong atau bernilai None!")
            return ""

        # Encode file bytes menjadi base64 agar aman dikirimkan lewat JSON payload
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

        # Kirim HTTP POST ke Google Apps Script Web App
        response = requests.post(APPS_SCRIPT_WEB_APP_URL, json=payload, timeout=30)
        
        print("DEBUG RESPONSE JSON:", response.text) # Tambahan debug untuk memastikan respon Apps Script
        
        if response.status_code == 200:
            res_json = response.json()
            print(res_json) # Debug print sesuai instruksi
            
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
# 3. HELPER REIMBURSEMENT (GOOGLE SHEETS + DRIVE LINKS PER ITEM)
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
            return True
        return False

    except Exception as e:
        st.error(f"❌ Gagal menyimpan data Reimbursement: {type(e).__name__} - {str(e)}")
        return False


def get_all_reimbursements():
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
            
            # --- Karena yang disimpan sudah berupa URL thumbnail, sanitasi /file/d/ tidak diperlukan lagi ---
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
            return True
        return False

    except Exception as e:
        st.error(f"❌ Gagal memperbarui status Approval di Sheet: {e}")
        return False


# ==============================================================================
# 4. HELPER DELIVERY ORDER (SCM - DB MATERIAL OUT)
# ==============================================================================

def generate_do_number():
    now = datetime.now()
    month_roman = get_roman_month(now.month)
    year_str = now.strftime("%Y")
    prefix_suffix = f"/CLX/DO/{month_roman}/{year_str}"

    try:
        sh = get_google_sheet_connection()
        worksheet = sh.worksheet(SHEET_MATERIAL_OUT)
        do_nos = worksheet.col_values(2)
        
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
    try:
        sh = get_google_sheet_connection()
        worksheet = sh.worksheet(SHEET_MATERIAL_OUT)

        records = []
        for row in rows_data:
            records.append([
                row.get("No"),
                row.get("No. DO"),
                row.get("Delv. Date"),
                row.get("Material Code"),
                row.get("Material Name"),
                row.get("Qty"),
                row.get("uom"),
                row.get("Site Alocation"),
                row.get("Remarks"),
                row.get("To"),
                row.get("Phone No."),
                row.get("Address"),
                row.get("EPC")
            ])

        worksheet.append_rows(records)
        return True

    except Exception as e:
        st.error(f"❌ Gagal menyimpan data DO ke DB Material Out: {e}")
        return False


def get_all_do_numbers():
    try:
        sh = get_google_sheet_connection()
        worksheet = sh.worksheet(SHEET_MATERIAL_OUT)
        do_nos = worksheet.col_values(2)
        
        unique_dos = sorted(list(set([no for no in do_nos[1:] if no.strip()])))
        return unique_dos
    except Exception:
        return []


def get_do_by_number(no_do_search):
    try:
        sh = get_google_sheet_connection()
        worksheet = sh.worksheet(SHEET_MATERIAL_OUT)
        all_data = worksheet.get_all_values()

        if not all_data or len(all_data) < 2:
            return None

        matching_rows = []
        for row in all_data[1:]:
            if len(row) >= 13 and row[1] == no_do_search:
                matching_rows.append({
                    "No": row[0],
                    "No. DO": row[1],
                    "Delv. Date": row[2],
                    "Material Code": row[3],
                    "Material Name": row[4],
                    "Qty": int(row[5]) if str(row[5]).isdigit() else row[5],
                    "uom": row[6],
                    "Site Allocation": row[7],
                    "Remarks": row[8],
                    "To": row[9],
                    "Phone No.": row[10],
                    "Address": row[11],
                    "EPC": row[12]
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
            "materials": matching_rows
        }

    except Exception as e:
        st.error(f"❌ Error saat mengambil data DO: {e}")
        return None


def update_do_in_db_material_out(no_do_target, updated_rows_data):
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
                row.get("No"),
                row.get("No. DO"),
                row.get("Delv. Date"),
                row.get("Material Code"),
                row.get("Material Name"),
                row.get("Qty"),
                row.get("uom"),
                row.get("Site Allocation"),
                row.get("Remarks"),
                row.get("To"),
                row.get("Phone No."),
                row.get("Address"),
                row.get("EPC")
            ])

        worksheet.clear()
        worksheet.update('A1', kept_rows)
        return True

    except Exception as e:
        st.error(f"❌ Gagal memperbarui data DO di Google Sheet: {e}")
        return False