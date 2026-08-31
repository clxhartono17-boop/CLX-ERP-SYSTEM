# ==============================================================================
# APP.PY
# CLX ERP SYSTEM
#
# FEATURES:
# 1. Authentication
# 2. User Authorization
# 3. Project Division
# 4. Operation Division
# 5. Commercial Division
# 6. SCM Division
# 7. General / Lapangan
# 8. Configuration
#
# CONFIGURATION:
# - Add Mitra
# - Add Team
# - Add Supplier
# - Add Project
# ==============================================================================

import importlib
import os
import traceback

import streamlit as st

from core.database import get_google_sheet_connection


# ==============================================================================
# 1. PAGE CONFIG
# ==============================================================================

st.set_page_config(
    page_title="CLX ERP SYSTEM",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==============================================================================
# 2. CSS
# ==============================================================================

def load_css(
    file_name="style.css"
):

    if os.path.exists(
        file_name
    ):

        try:

            with open(
                file_name,
                "r",
                encoding="utf-8"
            ) as f:

                st.markdown(
                    f"<style>{f.read()}</style>",
                    unsafe_allow_html=True
                )

        except Exception as e:

            st.warning(
                f"Gagal membaca CSS: {e}"
            )


load_css()


def custom_loading_card(
    text="Memproses data... Mohon tunggu"
):

    return f"""
    <div class="custom-loader-card">

        <div class="custom-loader-dots">
            <div></div>
            <div></div>
            <div></div>
        </div>

        <span style="
            font-weight:600;
            color:#0F172A;
            font-size:0.95rem;
        ">
            {text}
        </span>

    </div>
    """


def apply_login_styling():

    css = """
    <style>

    .stApp {
        background-color: #f8fafc;
    }

    .main .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
    }

    div[data-testid="stColumn"]:nth-child(2) {

        background-color: #ffffff;

        padding: 40px 32px;

        border-radius: 16px;

        box-shadow:
            0px 10px 25px -5px rgba(0,0,0,0.05),
            0px 8px 10px -6px rgba(0,0,0,0.01);

        border: 1px solid #e2e8f0;
    }

    </style>
    """

    st.markdown(
        css,
        unsafe_allow_html=True
    )


# ==============================================================================
# 3. GOOGLE SHEETS
# ==============================================================================

SPREADSHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1FU1lL3ls3jP_hAxBdx_Fu35Z9Ap4ICdHmOpMvCyA3gY"
    "/edit#gid=2146027013"
)


def load_users_from_sheets():

    try:

        sh = get_google_sheet_connection()

        worksheet = sh.worksheet(
            "Otorisasi"
        )

        all_rows = (
            worksheet.get_all_values()
        )

        users = {}

        for row in all_rows[2:]:

            if len(row) < 2:

                continue

            email = str(
                row[1]
            ).strip().lower()

            if not email:

                continue

            def is_true(value):

                return (
                    str(value)
                    .strip()
                    .upper()
                    in [
                        "TRUE",
                        "1",
                        "YES",
                        "V",
                        "CHECKED",
                    ]
                )

            users[email] = {

                "name":
                    row[2]
                    if len(row) > 2
                    else "",

                "role":
                    row[3]
                    if len(row) > 3
                    else "",

                "password":
                    (
                        str(
                            row[4]
                        ).strip()
                        if len(row) > 4
                        else ""
                    ),

                "access": {

                    "Project Div":
                        is_true(row[5])
                        if len(row) > 5
                        else False,

                    "Operation Div":
                        is_true(row[6])
                        if len(row) > 6
                        else False,

                    "Commercial Div":
                        is_true(row[7])
                        if len(row) > 7
                        else False,

                    "SCM Div":
                        is_true(row[8])
                        if len(row) > 8
                        else False,

                    "General / Lapangan":
                        is_true(row[9])
                        if len(row) > 9
                        else False,
                },
            }

        return users

    except Exception as e:

        st.error(
            f"Gagal menghubungkan ke "
            f"Sheet Otorisasi: {e}"
        )

        return {}


def save_password_to_sheet(
    email_to_update,
    new_password
):

    try:

        sh = get_google_sheet_connection()

        worksheet = sh.worksheet(
            "Otorisasi"
        )

        emails = (
            worksheet.col_values(
                2
            )[2:]
        )

        email_clean = (
            email_to_update
            .strip()
            .lower()
        )

        for idx, email in enumerate(
            emails,
            start=3
        ):

            if (
                email.strip().lower()
                == email_clean
            ):

                worksheet.update_cell(
                    idx,
                    13,
                    str(new_password)
                )

                return True

        return False

    except Exception as e:

        st.error(
            "Gagal memperbarui Password "
            f"di Google Sheets: {e}"
        )

        return False


# ==============================================================================
# 4. IMPORT MODULE
# ==============================================================================


# ------------------------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------------------------

show_configuration_page = None
configuration_error_msg = None

try:

    configuration_module = (
        importlib.import_module(
            "views.00_configuration.configuration"
        )
    )

    show_configuration_page = (
        getattr(
            configuration_module,
            "render",
            None
        )
        or
        getattr(
            configuration_module,
            "show",
            None
        )
    )

except Exception as e:

    configuration_error_msg = (
        f"Detail Exception: {e}\n\n"
        f"Traceback Lengkap:\n"
        f"{traceback.format_exc()}"
    )


# ------------------------------------------------------------------------------
# SPK
# ------------------------------------------------------------------------------

show_spk_page = None
spk_error_msg = None

try:

    spk_module = importlib.import_module(
        "views.01_project.spk"
    )

    show_spk_page = (
        getattr(
            spk_module,
            "show_spk_page",
            None
        )
        or
        getattr(
            spk_module,
            "render",
            None
        )
    )

except Exception as e:

    spk_error_msg = (
        f"Detail Exception: {e}\n\n"
        f"Traceback Lengkap:\n"
        f"{traceback.format_exc()}"
    )


# ------------------------------------------------------------------------------
# BOQ
# ------------------------------------------------------------------------------

show_boq_page = None
boq_error_msg = None

try:

    boq_module = importlib.import_module(
        "views.01_project.boq"
    )

    show_boq_page = (
        getattr(
            boq_module,
            "render",
            None
        )
        or
        getattr(
            boq_module,
            "show",
            None
        )
    )

except Exception as e:

    boq_error_msg = (
        f"Detail Exception: {e}\n\n"
        f"Traceback Lengkap:\n"
        f"{traceback.format_exc()}"
    )


# ------------------------------------------------------------------------------
# DO
# ------------------------------------------------------------------------------

show_scm_page = None
scm_error_msg = None

try:

    do_module = importlib.import_module(
        "views.03_scm.do"
    )

    show_scm_page = (
        getattr(
            do_module,
            "render",
            None
        )
        or
        getattr(
            do_module,
            "show_scm_page",
            None
        )
        or
        getattr(
            do_module,
            "show",
            None
        )
    )

except Exception as e:

    scm_error_msg = (
        f"Detail Exception: {e}\n\n"
        f"Traceback Lengkap:\n"
        f"{traceback.format_exc()}"
    )


# ------------------------------------------------------------------------------
# PO
# ------------------------------------------------------------------------------

show_po_page = None
po_error_msg = None

try:

    po_module = importlib.import_module(
        "views.03_scm.po"
    )

    show_po_page = (
        getattr(
            po_module,
            "render_po_module",
            None
        )
        or
        getattr(
            po_module,
            "render",
            None
        )
        or
        getattr(
            po_module,
            "show_po_page",
            None
        )
        or
        getattr(
            po_module,
            "show",
            None
        )
    )

except Exception as e:

    po_error_msg = (
        f"Detail Exception: {e}\n\n"
        f"Traceback Lengkap:\n"
        f"{traceback.format_exc()}"
    )


# ------------------------------------------------------------------------------
# REIMBURSEMENT
# ------------------------------------------------------------------------------

show_reimburse_page = None
reimburse_error_msg = None

try:

    reimburse_module = (
        importlib.import_module(
            "views.04_commercial.reimbursement"
        )
    )

    show_reimburse_page = (
        getattr(
            reimburse_module,
            "render",
            None
        )
        or
        getattr(
            reimburse_module,
            "show_reimbursement_page",
            None
        )
        or
        getattr(
            reimburse_module,
            "show",
            None
        )
    )

except Exception as e:

    reimburse_error_msg = (
        f"Detail Exception: {e}\n\n"
        f"Traceback Lengkap:\n"
        f"{traceback.format_exc()}"
    )


# ------------------------------------------------------------------------------
# INVOICE
# ------------------------------------------------------------------------------

show_invoice_page = None
invoice_error_msg = None

try:

    invoice_module = (
        importlib.import_module(
            "views.04_commercial.create_invoice"
        )
    )

    show_invoice_page = (
        getattr(
            invoice_module,
            "render",
            None
        )
        or
        getattr(
            invoice_module,
            "show_invoice_page",
            None
        )
        or
        getattr(
            invoice_module,
            "show",
            None
        )
    )

except Exception as e:

    invoice_error_msg = (
        f"Detail Exception: {e}\n\n"
        f"Traceback Lengkap:\n"
        f"{traceback.format_exc()}"
    )


# ==============================================================================
# 5. AUTHENTICATION
# ==============================================================================

if "authenticated" not in st.session_state:

    st.session_state.authenticated = False


if not st.session_state.authenticated:

    apply_login_styling()

    col1, col2, col3 = st.columns(
        [1, 2, 1]
    )

    with col2:

        logo_path = os.path.join(
            "assets",
            "CLX.png"
        )

        if os.path.exists(
            logo_path
        ):

            sub1, sub2, sub3 = st.columns(
                [1, 2, 1]
            )

            with sub2:

                st.image(
                    logo_path,
                    use_container_width=True
                )

        st.markdown(
            """
            <h3 style="
                text-align:center;
                color:#1e293b;
                margin-top:10px;
            ">
                🔒 CLX ERP System
            </h3>
            """,
            unsafe_allow_html=True
        )

        st.markdown(
            """
            <p style="
                text-align:center;
                color:#64748b;
                font-size:14px;
            ">
                Silakan masuk dengan akun Anda
            </p>
            """,
            unsafe_allow_html=True
        )

        st.write("")

        tab_login, tab_register = st.tabs(
            [
                "🔐 Login",
                "🔑 Buat Password Baru"
            ]
        )

        user_db = (
            load_users_from_sheets()
        )

        # ----------------------------------------------------------------------
        # LOGIN
        # ----------------------------------------------------------------------

        with tab_login:

            login_email = st.text_input(
                "Email",
                key="login_email"
            )

            login_password = st.text_input(
                "Password",
                type="password",
                key="login_pass"
            )

            if st.button(
                "Login",
                use_container_width=True,
                type="primary"
            ):

                email_clean = (
                    login_email
                    .strip()
                    .lower()
                )

                if email_clean not in user_db:

                    st.error(
                        "❌ Email belum terdaftar "
                        "di sheet Otorisasi."
                    )

                else:

                    user_info = (
                        user_db[
                            email_clean
                        ]
                    )

                    if not user_info[
                        "password"
                    ]:

                        st.warning(
                            "⚠️ Email Anda belum "
                            "memiliki password. "
                            "Silakan buat password "
                            "di tab 'Buat Password Baru'."
                        )

                    elif (
                        user_info["password"]
                        ==
                        login_password.strip()
                    ):

                        st.session_state[
                            "authenticated"
                        ] = True

                        st.session_state[
                            "user_data"
                        ] = user_info

                        st.session_state[
                            "user_email"
                        ] = email_clean

                        st.success(
                            f"Selamat datang, "
                            f"{user_info['name']}!"
                        )

                        st.rerun()

                    else:

                        st.error(
                            "❌ Password salah."
                        )

        # ----------------------------------------------------------------------
        # REGISTER
        # ----------------------------------------------------------------------

        with tab_register:

            st.caption(
                "Khusus user yang emailnya "
                "sudah didaftarkan di Sheet "
                "Otorisasi tetapi belum "
                "memiliki password."
            )

            reg_email = st.text_input(
                "Email Terdaftar",
                key="reg_email"
            )

            reg_pass1 = st.text_input(
                "Password Baru",
                type="password",
                key="reg_pass1"
            )

            reg_pass2 = st.text_input(
                "Konfirmasi Password Baru",
                type="password",
                key="reg_pass2"
            )

            if st.button(
                "Simpan Password",
                use_container_width=True,
                type="secondary"
            ):

                email_clean = (
                    reg_email
                    .strip()
                    .lower()
                )

                if email_clean not in user_db:

                    st.error(
                        "❌ Email Anda belum "
                        "terdaftar di sheet "
                        "Otorisasi. "
                        "Silakan hubungi Admin."
                    )

                elif reg_pass1 != reg_pass2:

                    st.error(
                        "❌ Konfirmasi password "
                        "tidak cocok!"
                    )

                elif len(reg_pass1) < 3:

                    st.warning(
                        "⚠️ Password minimal "
                        "3 karakter."
                    )

                else:

                    loader = st.empty()

                    loader.markdown(
                        custom_loading_card(
                            "Menyimpan password "
                            "ke Google Sheet..."
                        ),
                        unsafe_allow_html=True
                    )

                    success = (
                        save_password_to_sheet(
                            email_clean,
                            reg_pass1
                        )
                    )

                    loader.empty()

                    if success:

                        st.success(
                            "✅ Password berhasil "
                            "disimpan di Sheet "
                            "Otorisasi! "
                            "Silakan Login."
                        )

                    else:

                        st.error(
                            "🚨 Gagal memperbarui "
                            "password di Google Sheets."
                        )

    st.stop()


# ==============================================================================
# 6. USER ACCESS
# ==============================================================================

current_user = (
    st.session_state[
        "user_data"
    ]
)

user_access = (
    current_user[
        "access"
    ]
)


# ==============================================================================
# 7. MENU
# ==============================================================================

structure_menu = {

    "Project Div": [
        "Create SPK",
        "Create BOQ",
    ],

    "Operation Div": [
        "Operation Status",
    ],

    "Commercial Div": [
        "Create PO",
        "Create Invoice",
        "Form Reimbursement",
    ],

    "SCM Div": [
        "Create DO",
    ],

    "General / Lapangan": [
        "Form Reimbursement",
    ],

    # --------------------------------------------------------------------------
    # CONFIGURATION
    #
    # Configuration tidak dimasukkan ke Otorisasi biasa.
    # Akses diberikan khusus berdasarkan Role Admin.
    # --------------------------------------------------------------------------

    "Configuration": [
        "Master Data",
    ],
}


# ==============================================================================
# 8. ALLOWED DIVISIONS
# ==============================================================================

allowed_divs = [

    div

    for div, has_access
    in user_access.items()

    if has_access
    and div in structure_menu

]


# ==============================================================================
# 9. CONFIGURATION ACCESS
# ==============================================================================

current_role = str(
    current_user.get(
        "role",
        ""
    )
).strip().lower()


if current_role in [
    "admin",
    "administrator",
]:

    if "Configuration" not in allowed_divs:

        allowed_divs.append(
            "Configuration"
        )


# ==============================================================================
# 10. SIDEBAR
# ==============================================================================

with st.sidebar:

    logo_path = os.path.join(
        "assets",
        "CLX.png"
    )

    if os.path.exists(
        logo_path
    ):

        st.image(
            logo_path,
            use_container_width=True
        )

    else:

        st.title(
            "🏢 CLX ERP SYSTEM"
        )

    st.markdown("---")

    if allowed_divs:

        selected_div = st.selectbox(
            "Pilih Divisi / Modul",
            allowed_divs
        )

        selected_menu = st.radio(
            "Navigasi Menu",
            structure_menu[
                selected_div
            ]
        )

    else:

        st.warning(
            "Anda tidak memiliki "
            "hak akses ke modul manapun."
        )

        selected_menu = None

    st.markdown("---")

    st.caption(
        f"User: **{current_user['name']}**"
    )

    st.caption(
        f"Role: **{current_user['role']}**"
    )

    if st.button(
        "Logout",
        type="secondary",
        use_container_width=True
    ):

        st.session_state[
            "authenticated"
        ] = False

        st.session_state[
            "user_data"
        ] = None

        st.rerun()


# ==============================================================================
# 11. ROUTER
# ==============================================================================


# ------------------------------------------------------------------------------
# CREATE SPK
# ------------------------------------------------------------------------------

if selected_menu == "Create SPK":

    if show_spk_page:

        show_spk_page()

    else:

        st.error(
            "🚨 Gagal Memuat Halaman SPK"
        )

        if spk_error_msg:

            st.code(
                spk_error_msg,
                language="python"
            )


# ------------------------------------------------------------------------------
# CREATE BOQ
# ------------------------------------------------------------------------------

elif selected_menu == "Create BOQ":

    if show_boq_page:

        show_boq_page()

    else:

        st.error(
            "🚨 Gagal Memuat Halaman BOQ"
        )

        if boq_error_msg:

            st.code(
                boq_error_msg,
                language="python"
            )


# ------------------------------------------------------------------------------
# OPERATION STATUS
# ------------------------------------------------------------------------------

elif selected_menu == "Operation Status":

    st.title(
        "🚧 Operation Div"
    )

    st.info(
        "Modul Operation masih dalam "
        "tahap pengembangan."
    )


# ------------------------------------------------------------------------------
# CREATE DO
# ------------------------------------------------------------------------------

elif selected_menu == "Create DO":

    if show_scm_page:

        show_scm_page()

    else:

        st.error(
            "🚨 Gagal Memuat Halaman "
            "SCM Delivery Order"
        )

        if scm_error_msg:

            st.code(
                scm_error_msg,
                language="python"
            )


# ==============================================================================
# CREATE PO
# ==============================================================================

elif selected_menu == "Create PO":

    if show_po_page:

        show_po_page()

    else:

        st.error(
            "🚨 Gagal Memuat Halaman "
            "Purchase Order"
        )

        if po_error_msg:

            st.code(
                po_error_msg,
                language="python"
            )

        else:

            st.warning(
                "Modul PO belum berhasil "
                "ditemukan."
            )


# ------------------------------------------------------------------------------
# CREATE INVOICE
# ------------------------------------------------------------------------------

elif selected_menu == "Create Invoice":

    if show_invoice_page:

        show_invoice_page()

    else:

        st.error(
            "🚨 Gagal Memuat Halaman "
            "Create Invoice"
        )

        if invoice_error_msg:

            st.code(
                invoice_error_msg,
                language="python"
            )


# ------------------------------------------------------------------------------
# FORM REIMBURSEMENT
# ------------------------------------------------------------------------------

elif selected_menu == "Form Reimbursement":

    if show_reimburse_page:

        show_reimburse_page()

    else:

        st.error(
            "🚨 Gagal Memuat Halaman "
            "Reimbursement"
        )

        if reimburse_error_msg:

            st.code(
                reimburse_error_msg,
                language="python"
            )


# ==============================================================================
# CONFIGURATION
# ==============================================================================

elif selected_menu == "Master Data":

    # --------------------------------------------------------------------------
    # EXTRA SECURITY
    # --------------------------------------------------------------------------

    if current_role not in [
        "admin",
        "administrator",
    ]:

        st.error(
            "🚫 Anda tidak memiliki "
            "hak akses ke Configuration."
        )

    elif show_configuration_page:

        show_configuration_page()

    else:

        st.error(
            "🚨 Gagal Memuat Halaman "
            "Configuration"
        )

        if configuration_error_msg:

            st.code(
                configuration_error_msg,
                language="python"
            )

        else:

            st.warning(
                "Modul Configuration belum "
                "berhasil ditemukan."
            )
