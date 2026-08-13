import os
import base64
import importlib
import traceback
import streamlit as st
import gspread

# Mengimpor koneksi dari folder core/database.py yang sudah menggunakan OAuth 2.0
from core.database import get_google_sheet_connection

# ==============================================================================
# 1. KONFIGURASI HALAMAN STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="CLX ERP SYSTEM",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================================================================
# 2. HELPER STYLING & HELPER COMPONENTS
# ==============================================================================
def load_css(file_name="style.css"):
    if os.path.exists(file_name):
        with open(file_name) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

load_css("style.css")

def apply_login_styling():
    """CSS custom untuk merapikan layout background dan card login"""
    css = """
    <style>
    /* Styling latar belakang halaman login */
    .stApp {
        background-color: #f8fafc;
    }
    
    /* Menghilangkan padding berlebih dari Streamlit */
    .main .block-container {
        padding-top: 3rem;
        padding-bottom: 2rem;
    }

    /* Membuat Card Login Tampil Modern, Clean, & Elegan */
    div[data-testid="stColumn"]:nth-child(2) {
        background-color: #ffffff;
        padding: 40px 32px;
        border-radius: 16px;
        box-shadow: 0px 10px 25px -5px rgba(0, 0, 0, 0.05), 0px 8px 10px -6px rgba(0, 0, 0, 0.01);
        border: 1px solid #e2e8f0;
    }
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

def render_metric_card(title, value, delta=None, is_positive=True):
    delta_html = ""
    if delta:
        delta_class = "metric-delta-positive" if is_positive else "metric-delta-negative"
        arrow = "▲" if is_positive else "▼"
        delta_html = f'<div class="{delta_class}">{arrow} {delta}</div>'
        
    html_code = f"""
    <div class="metric-card">
        <div class="metric-title">{title}</div>
        <div class="metric-value">{value}</div>
        {delta_html}
    </div>
    """
    st.markdown(html_code, unsafe_allow_html=True)


# ==============================================================================
# 3. KONEKSI & BACA/TULIS DATA GOOGLE SHEETS VIA GSPREAD (OAuth 2.0)
# ==============================================================================
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1FU1lL3ls3jP_hAxBdx_Fu35Z9Ap4ICdHmOpMvCyA3gY/edit#gid=2146027013"

def load_users_from_sheets():
    """Membaca data Otorisasi berdasarkan indeks posisi kolom (Aman dari Merged Cell)"""
    try:
        sh = get_google_sheet_connection()
        worksheet = sh.worksheet("Otorisasi")
        
        all_rows = worksheet.get_all_values()
        
        users = {}
        # Data pengguna dimulai dari baris ke-3 (Indeks 2 pada Python)
        for row in all_rows[2:]:
            if len(row) >= 2:
                email = str(row[1]).strip().lower()  # Kolom B = Index 1
                
                if email:
                    def is_true(val):
                        return str(val).upper() in ["TRUE", "1", "YES", "V", "CHECKED"]

                    users[email] = {
                        "name": row[2] if len(row) > 2 else "",
                        "role": row[3] if len(row) > 3 else "",
                        "password": str(row[12]).strip() if len(row) > 12 else "",
                        "access": {
                            "Project Div": is_true(row[4]) if len(row) > 4 else False,
                            "Operation Div": is_true(row[5]) if len(row) > 5 else False,
                            "Commercial Div": is_true(row[6]) if len(row) > 6 else False,
                            "SCM Div": is_true(row[7]) if len(row) > 7 else False,
                            "General / Lapangan": is_true(row[8]) if len(row) > 8 else False,
                        }
                    }
        return users
    except Exception as e:
        st.error(f"Gagal menghubungkan ke Sheet Otorisasi: {e}")
        return {}

def save_password_to_sheet(email_to_update, new_password):
    """Menyimpan/Mengubah Password ke Kolom J di Sheet Otorisasi"""
    try:
        sh = get_google_sheet_connection()
        worksheet = sh.worksheet("Otorisasi")
        
        emails = worksheet.col_values(2)[2:]
        email_clean = email_to_update.strip().lower()
        
        for idx, email in enumerate(emails, start=3):
            if email.strip().lower() == email_clean:
                worksheet.update_cell(idx, 13, str(new_password))
                return True
        return False
    except Exception as e:
        st.error(f"Gagal memperbarui Password di Google Sheets: {e}")
        return False


# ==============================================================================
# 4. IMPORT MODUL VIEW / HALAMAN DYNAMICALLY
# ==============================================================================
try:
    spk_module = importlib.import_module("views.01_project.spk")
    show_spk_page = getattr(spk_module, "show_spk_page", None) or getattr(spk_module, "render", None)
except Exception:
    show_spk_page = None

try:
    boq_module = importlib.import_module("views.01_project.boq")
    show_boq_page = getattr(boq_module, "render", None) or getattr(boq_module, "show", None)
except Exception:
    show_boq_page = None

scm_error_msg = None
try:
    scm_module = importlib.import_module("views.03_scm.do")
    show_scm_page = getattr(scm_module, "render", None) or getattr(scm_module, "show_scm_page", None) or getattr(scm_module, "show", None)
    if show_scm_page is None:
        scm_error_msg = "File `views/03_scm/do.py` berhasil dimuat, tetapi fungsi `render()` atau `show()` tidak ditemukan."
except Exception as e:
    show_scm_page = None
    scm_error_msg = f"Detail Exception: {e}\n\nTraceback Lengkap:\n{traceback.format_exc()}"

reimburse_error_msg = None
try:
    reimburse_module = importlib.import_module("views.04_commercial.reimbursement")
    show_reimburse_page = getattr(reimburse_module, "render", None) or getattr(reimburse_module, "show_reimbursement_page", None) or getattr(reimburse_module, "show", None)
    if show_reimburse_page is None:
        reimburse_error_msg = "File `reimbursement.py` berhasil dimuat, tetapi fungsi `render()` tidak ditemukan."
except Exception as e:
    show_reimburse_page = None
    reimburse_error_msg = f"Detail Exception: {e}\n\nTraceback Lengkap:\n{traceback.format_exc()}"

invoice_error_msg = None
try:
    invoice_module = importlib.import_module("views.04_commercial.create_invoice")
    show_invoice_page = getattr(invoice_module, "render", None) or getattr(invoice_module, "show_invoice_page", None) or getattr(invoice_module, "show", None)
    if show_invoice_page is None:
        invoice_error_msg = "File `create_invoice.py` berhasil dimuat, tetapi fungsi `render()` tidak ditemukan."
except Exception as e:
    show_invoice_page = None
    invoice_error_msg = f"Detail Exception: {e}\n\nTraceback Lengkap:\n{traceback.format_exc()}"


# ==============================================================================
# 5. SISTEM AUTHENTICATION (LOGIN & REGISTER PASSWORD)
# ==============================================================================
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    apply_login_styling()

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        LOGO_PATH = os.path.join("assets", "Logo.png")
        if os.path.exists(LOGO_PATH):
            sub_col1, sub_col2, sub_col3 = st.columns([1, 2, 1])
            with sub_col2:
                st.image(LOGO_PATH, use_container_width=True)

        st.markdown("<h3 style='text-align: center; color: #1e293b; margin-top: 10px;'>🔒 CLX ERP System</h3>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #64748b; font-size: 14px;'>Silakan masuk dengan akun Anda</p>", unsafe_allow_html=True)
        st.write("")

        tab_login, tab_register = st.tabs(["🔐 Login", "🔑 Buat Password Baru"])

        user_db = load_users_from_sheets()

        with tab_login:
            login_email = st.text_input("Email", key="login_email", placeholder="")
            login_password = st.text_input("Password", type="password", key="login_pass")
            btn_login = st.button("Login", use_container_width=True, type="primary")

            if btn_login:
                email_clean = login_email.strip().lower()
                if email_clean in user_db:
                    user_info = user_db[email_clean]
                    
                    if not user_info["password"]:
                        st.warning("⚠️ Email Anda belum memiliki password. Silakan buat password di tab 'Buat Password Baru'.")
                    elif user_info["password"] == login_password.strip():
                        st.session_state["authenticated"] = True
                        st.session_state["user_data"] = user_info
                        st.session_state["user_email"] = email_clean
                        st.success(f"Selamat datang, {user_info['name']}!")
                        st.rerun()
                    else:
                        st.error("❌ Password salah.")
                else:
                    st.error("❌ Email belum terdaftar di sheet Otorisasi.")

        with tab_register:
            st.caption("Khusus user yang emailnya sudah didaftarkan di Sheet Otorisasi tetapi belum memiliki password.")
            reg_email = st.text_input("Email Terdaftar", key="reg_email", placeholder="masukkan email terdaftar...")
            reg_pass1 = st.text_input("Password Baru", type="password", key="reg_pass1")
            reg_pass2 = st.text_input("Konfirmasi Password Baru", type="password", key="reg_pass2")
            btn_register = st.button("Simpan Password", use_container_width=True, type="secondary")

            if btn_register:
                email_clean = reg_email.strip().lower()
                
                if email_clean not in user_db:
                    st.error("❌ Email Anda belum terdaftar di sheet Otorisasi. Silakan hubungi Admin.")
                elif reg_pass1 != reg_pass2:
                    st.error("❌ Konfirmasi password tidak cocok!")
                elif len(reg_pass1) < 3:
                    st.warning("⚠️ Password minimal 3 karakter.")
                else:
                    with st.spinner("Menyimpan password ke Google Sheet..."):
                        if save_password_to_sheet(email_clean, reg_pass1):
                            st.success("✅ Password berhasil disimpan di Sheet Otorisasi! Silakan Login.")
                        else:
                            st.error("🚨 Gagal memperbarui password di Google Sheets.")

    st.stop()


# ==============================================================================
# 6. NAVIGASI SIDEBAR & HAK AKSES USER
# ==============================================================================
current_user = st.session_state["user_data"]
user_access = current_user["access"]

structure_menu = {
    "Project Div": ["Create SPK", "Create BOQ"],
    "Operation Div": ["Operation Status"],
    "Commercial Div": ["Create PO", "Create Invoice", "Form Reimbursement"],
    "SCM Div": ["Create DO"],
    "General / Lapangan": ["Form Reimbursement"],
}

allowed_divs = [div for div, has_access in user_access.items() if has_access and div in structure_menu]

with st.sidebar:
    LOGO_PATH = os.path.join("assets", "Logo.png")
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, use_container_width=True)
    else:
        st.title("🏢 CLX ERP SYSTEM")
        
    st.markdown("---")

    if allowed_divs:
        selected_div = st.selectbox("Pilih Divisi / Modul", allowed_divs)
        selected_menu = st.radio("Navigasi Menu", structure_menu[selected_div])
    else:
        st.warning("Anda tidak memiliki hak akses ke modul manapun.")
        selected_menu = None

    st.markdown("---")
    st.caption(f"User: **{current_user['name']}**")
    st.caption(f"Role: **{current_user['role']}**")
    
    if st.button("Logout", type="secondary", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["user_data"] = None
        st.rerun()


# ==============================================================================
# 7. ROUTER HALAMAN
# ==============================================================================
if selected_menu == "Create SPK":
    if show_spk_page:
        show_spk_page()
    else:
        st.warning("Halaman SPK Generator belum siap.")

elif selected_menu == "Create BOQ":
    if show_boq_page:
        show_boq_page()
    else:
        st.warning("Halaman BOQ Generator belum siap.")

elif selected_menu == "Operation Status":
    st.title("🚧 Operation Div")
    st.info("Modul Operation masih dalam tahap pengembangan (Under Development).")

elif selected_menu == "Create DO":
    if show_scm_page:
        show_scm_page()
    else:
        st.error("🚨 Gagal Memuat Halaman SCM Delivery Order")
        if scm_error_msg: st.code(scm_error_msg, language="python")

elif selected_menu == "Create PO":
    st.title("🛍️ Create PO")
    st.info("Modul Create Purchase Order (PO) masih dalam tahap pengembangan.")

elif selected_menu == "Create Invoice":
    if show_invoice_page:
        show_invoice_page()
    else:
        st.error("🚨 Gagal Memuat Halaman Create Invoice")
        if invoice_error_msg: st.code(invoice_error_msg, language="python")

elif selected_menu == "Form Reimbursement":
    if show_reimburse_page:
        show_reimburse_page()
    else:
        st.error("🚨 Gagal Memuat Halaman Reimbursement")
        if reimburse_error_msg: st.code(reimburse_error_msg, language="python")