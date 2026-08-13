import gspread
from google.oauth2.service_account import Credentials
import streamlit as st


def get_google_sheet_connection():
    """Menghubungkan ke Google Sheets menggunakan st.secrets atau service_account.json"""
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]

        # Menggunakan Streamlit Secrets jika ada
        if "gcp_service_account" in st.secrets:
            creds = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"], scopes=scope
            )
        else:
            # Menggunakan file json lokal
            creds = Credentials.from_service_account_file(
                "service_account.json", scopes=scope
            )

        client = gspread.authorize(creds)
        
        # Sesuai dengan Nama File Spreadsheet Google Sheet kamu di Screenshot
        spreadsheet = client.open("CLX ERP SYSTEM")
        return spreadsheet
    except Exception as e:
        st.sidebar.warning(f"⚠️ Google Sheets offline / tidak terkoneksi: {e}")
        return None