import datetime
import io
import os
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    Image,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# Import Koneksi Database sesuai Struktur ERP
from core.database import get_google_sheet_connection

# PIN Rahasia untuk Login COO (Bisa kamu ubah)
COO_PIN_SECRET = "1234"

# ==============================================================================
# CSS OVERRIDE ULTIMATE (Menebas Mode Gelap pada Selectbox & Data Editor)
# ==============================================================================
PASTEL_ORANGE_CSS = """
<style>
    /* 1. Paksa Warna Utama Aplikasi Ke Mode Terang */
    .stApp, [data-testid="stAppViewContainer"] {
        background-color: #FAFAFA !important;
        color: #2D3748 !important;
    }

    /* 2. Fix Dropdown / Selectbox */
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

    /* 3. Text Input */
    div[data-baseweb="input"] > div,
    div[data-baseweb="input"] input {
        background-color: #FFFFFF !important;
        color: #1A202C !important;
        border-color: #FFCAD4 !important;
    }

    /* 4. Highlight Kode WO */
    code {
        background-color: #FFE5D9 !important;
        color: #C05621 !important;
        border: 1px solid #FFCAD4 !important;
        font-weight: bold !important;
        padding: 3px 8px !important;
        border-radius: 6px !important;
    }

    /* 5. FIX TABEL */
    div[data-testid="stDataEditor"],
    div[data-testid="stDataEditor"] > div,
    .dgb-grid, 
    canvas {
        background-color: #FFFFFF !important;
    }
</style>
"""


# ==============================================================================
# 1. HELPER: GENERATE NOMOR URUT SPK & DOKUMEN PDF
# ==============================================================================


def generate_spk_number(sow_type="GENERAL", sequence_num=1):
    """Menghasilkan Nomor SPK berurutan (misal: 0001/CLX/SPK/CONS/VII/2026)"""
    now = datetime.datetime.now()
    roman_months = [
        "I",
        "II",
        "III",
        "IV",
        "V",
        "VI",
        "VII",
        "VIII",
        "IX",
        "X",
        "XI",
        "XII",
    ]
    month_roman = roman_months[now.month - 1]

    sow_code = "SURVEY" if "survey" in str(sow_type).lower() else "CONS"
    seq_str = f"{sequence_num:04d}"
    return f"{seq_str}/CLX/SPK/{sow_code}/{month_roman}/{now.year}"


def generate_spk_pdf_bytes(
    selected_wo, selected_sites, spk_metadata, matched_sow_df
):
    """Membuat file PDF SPK dalam bentuk bytes"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=portrait(letter),
        leftMargin=0.4 * inch,
        rightMargin=0.4 * inch,
        topMargin=0.4 * inch,
        bottomMargin=0.4 * inch,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Heading1"],
        fontSize=12,
        leading=14,
        alignment=1,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#000000"),
        spaceAfter=2,
    )

    subtitle_style = ParagraphStyle(
        "DocSubtitle",
        parent=styles["Normal"],
        fontSize=9,
        leading=11,
        alignment=1,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#000000"),
        spaceAfter=10,
    )

    header_left = ParagraphStyle(
        "HeaderLeft",
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#1B365D"),
    )

    header_right = ParagraphStyle(
        "HeaderRight",
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#333333"),
    )

    meta_label = ParagraphStyle(
        "MetaLabel", fontName="Helvetica-Bold", fontSize=8, leading=10
    )
    meta_val = ParagraphStyle(
        "MetaVal", fontName="Helvetica", fontSize=8, leading=10
    )

    cell_head = ParagraphStyle(
        "CellHead",
        fontName="Helvetica-Bold",
        fontSize=6.5,
        leading=8,
        alignment=1,
        textColor=colors.white,
    )
    cell_body = ParagraphStyle(
        "CellBody", fontName="Helvetica", fontSize=6, leading=7.5
    )

    elements = []

    # --- 1. KOP SURAT ---
    logo_path = "assets/logo.png"
    if os.path.exists(logo_path):
        company_logo = Image(logo_path, width=1.5 * inch, height=1.0 * inch)
    else:
        company_logo = Paragraph(
            "<b>PT. Connectivity Leads eXcellence</b>", header_left
        )

    header_data = [
        [
            company_logo,
            Paragraph(
                "<b>PT. Connectivity Leads eXcellence</b><br/>"
                "<b>Jakarta Office:</b> Jl. M. Ali 2 No. 19 RT 007 RW 004 Tanah"
                " Baru Beji Kota Depok Jawa Barat<br/><b>E:</b>"
                " clx.central@gmail.com",
                header_right,
            ),
        ]
    ]
    t_header = Table(header_data, colWidths=[2.2 * inch, 5.1 * inch])
    t_header.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LINEBELOW", (0, 0), (-1, -1), 1, colors.HexColor("#1B365D")),
        ])
    )
    elements.append(t_header)
    elements.append(Spacer(1, 8))

    # --- 2. JUDUL SPK & TANGGAL ---
    no_spk = spk_metadata.get("no_spk", "0001/CLX/SPK/CONS/VII/2026")
    date_str = datetime.datetime.now().strftime("%d %B %Y")

    elements.append(Paragraph("SURAT PERINTAH KERJA (SPK)", title_style))
    elements.append(
        Paragraph(f"No. {no_spk}<br/>Jakarta, {date_str}", subtitle_style)
    )

    # --- 3. METADATA SPK ---
    meta_table_data = [
        [
            Paragraph("Proyek:", meta_label),
            Paragraph(spk_metadata.get("proyek", "-"), meta_val),
            Paragraph("No. WO:", meta_label),
            Paragraph(str(selected_wo), meta_val),
        ],
        [
            Paragraph("Pekerjaan:", meta_label),
            Paragraph(spk_metadata.get("pekerjaan", "-"), meta_val),
            Paragraph("Penanggung Jawab:", meta_label),
            Paragraph(
                f"{spk_metadata.get('pic_name','-')}"
                f" ({spk_metadata.get('pic_phone','-')})",
                meta_val,
            ),
        ],
        [
            Paragraph("Lokasi:", meta_label),
            Paragraph(spk_metadata.get("lokasi", "-"), meta_val),
            Paragraph("Penerbit (PT. CLX):", meta_label),
            Paragraph(
                "Wikantiyoso S. (0878-8855-0300)<br/>Hartono (0818-0690-9317)",
                meta_val,
            ),
        ],
        [
            Paragraph("Detail Site:", meta_label),
            Paragraph("Terlampir pada tabel di bawah", meta_val),
            Paragraph("SOW:", meta_label),
            Paragraph(spk_metadata.get("sow_type", "-"), meta_val),
        ],
    ]

    t_meta = Table(
        meta_table_data,
        colWidths=[1.0 * inch, 2.7 * inch, 1.2 * inch, 2.4 * inch],
    )
    t_meta.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8F9FA")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#E5E7EB")),
        ])
    )
    elements.append(t_meta)
    elements.append(Spacer(1, 10))

    # --- 4. TABEL URAIAN SOW ---
    sow_headers = [
        Paragraph("<b>No.</b>", cell_head),
        Paragraph("<b>Uraian Pekerjaan (SoW)</b>", cell_head),
        Paragraph("<b>Target Penyelesaian</b>", cell_head),
    ]
    sow_rows = [sow_headers]

    if not matched_sow_df.empty:
        for idx, (_, r) in enumerate(matched_sow_df.iterrows(), 1):
            deskripsi = str(r.iloc[2]) if r.shape[0] > 2 else "-"
            target = str(r.iloc[3]) if r.shape[0] > 3 else "1 Hari"

            sow_rows.append([
                Paragraph(str(idx), cell_body),
                Paragraph(deskripsi, cell_body),
                Paragraph(target, cell_body),
            ])
    else:
        sow_rows.append([
            Paragraph("1", cell_body),
            Paragraph("SOW Pekerjaan Standar", cell_body),
            Paragraph("1 Hari", cell_body),
        ])

    t_sow = Table(sow_rows, colWidths=[0.4 * inch, 5.0 * inch, 1.9 * inch])
    t_sow.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B365D")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BDC3C7")),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ])
    )
    elements.append(t_sow)
    elements.append(Spacer(1, 10))

    # --- 5. DETAIL SITE LIST ---
    elements.append(
        Paragraph(
            "<b>Detail Site List:</b>",
            ParagraphStyle(
                "SubHeader",
                fontName="Helvetica-Bold",
                fontSize=9,
                spaceAfter=4,
            ),
        )
    )

    site_headers = [
        Paragraph("<b>No</b>", cell_head),
        Paragraph("<b>Site Name</b>", cell_head),
        Paragraph("<b>Charging Type</b>", cell_head),
        Paragraph("<b>WO Number</b>", cell_head),
        Paragraph("<b>Province</b>", cell_head),
        Paragraph("<b>PIC + Contact</b>", cell_head),
        Paragraph("<b>Gmaps</b>", cell_head),
    ]

    site_rows = [site_headers]
    for idx, (_, row) in enumerate(selected_sites.iterrows(), 1):
        site_rows.append([
            Paragraph(str(idx), cell_body),
            Paragraph(str(row.get("col_site", "-")), cell_body),
            Paragraph(str(row.get("col_charge", "-")), cell_body),
            Paragraph(str(selected_wo), cell_body),
            Paragraph(str(row.get("col_province", "-")), cell_body),
            Paragraph(str(row.get("col_pic", "-")), cell_body),
            Paragraph(str(row.get("col_gmaps", "-")), cell_body),
        ])

    col_widths = [
        0.3 * inch,
        1.6 * inch,
        1.0 * inch,
        1.3 * inch,
        0.9 * inch,
        1.1 * inch,
        1.1 * inch,
    ]
    t_site = Table(site_rows, colWidths=col_widths)
    t_site.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B365D")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BDC3C7")),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ])
    )
    elements.append(t_site)
    elements.append(Spacer(1, 15))

    # --- 6. PENGESAHAN ---
    sign_style_bold = ParagraphStyle(
        "SignBold",
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
        alignment=1,
    )
    sign_style_norm = ParagraphStyle(
        "SignNorm", fontName="Helvetica", fontSize=8, leading=10, alignment=1
    )

    sign_data = [
        [
            Paragraph("<b>PELAKSANA PEKERJAAN</b>", sign_style_bold),
            Paragraph(
                "<b>PEMBERI PERINTAH KERJA</b><br/><b>PT. Connectivity Leads"
                " eXcellence</b>",
                sign_style_bold,
            ),
        ],
        [Spacer(1, 35), Spacer(1, 35)],
        [
            Paragraph(
                f"<b>{spk_metadata.get('pic_name', 'Edy')}</b><br/>Contact:"
                f" {spk_metadata.get('pic_phone', '-')}",
                sign_style_norm,
            ),
            Paragraph(
                "<b>Wikantiyoso Suyono</b><br/>Chief Operating Officer",
                sign_style_norm,
            ),
        ],
    ]

    t_sign = Table(sign_data, colWidths=[3.6 * inch, 3.7 * inch])
    t_sign.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ])
    )

    elements.append(KeepTogether([t_sign]))

    doc.build(elements)
    buffer.seek(0)
    return no_spk, buffer.getvalue()


# ==============================================================================
# 2. TAMPILAN STREAMLIT MENU SPK GENERATOR, TAKE OVER, & COO APPROVAL
# ==============================================================================


def show_spk_page():
    st.markdown(PASTEL_ORANGE_CSS, unsafe_allow_html=True)

    sh = get_google_sheet_connection()
    if not sh:
        st.error("❌ Koneksi ke Database tidak tersedia.")
        return

    # 📌 Navigasi Menu Utama
    main_menu = st.radio(
        "📌 **PILIH MENU OPERASIONAL:**",
        options=[
            "📄 Generate SPK Baru",
            "🔄 Take Over Site",
            "🔒 COO Approval Dashboard",
        ],
        horizontal=True,
        key="main_spk_menu",
    )

    st.markdown("---")

    # --------------------------------------------------------------------------
    # MENU 1: GENERATE SPK BARU
    # --------------------------------------------------------------------------
    if main_menu == "📄 Generate SPK Baru":
        st.subheader("📄 Generator Surat Perintah Kerja (SPK) & PDF")
        st.markdown(
            "Pilih Jenis SOW & Nomor WO dari **Sheet Query**, centang site yang"
            " diinginkan, lalu klik **Generate SPK**."
        )

        try:
            query_sheet = sh.worksheet("Query")
            query_rows = query_sheet.get_all_values()

            try:
                dropdown_sheet = sh.worksheet("Master Dropdown")
                dropdown_data = dropdown_sheet.get_all_values()
                df_dropdown = (
                    pd.DataFrame(dropdown_data[1:], columns=dropdown_data[0])
                    if len(dropdown_data) > 1
                    else pd.DataFrame()
                )
                df_dropdown = df_dropdown.loc[
                    :, ~df_dropdown.columns.duplicated()
                ].copy()
            except Exception:
                df_dropdown = pd.DataFrame()

            try:
                sow_sheet = sh.worksheet("Master SOW")
                sow_data = sow_sheet.get_all_values()
                df_master_sow = (
                    pd.DataFrame(sow_data[1:], columns=sow_data[0])
                    if len(sow_data) > 1
                    else pd.DataFrame()
                )
                df_master_sow = df_master_sow.loc[
                    :, ~df_master_sow.columns.duplicated()
                ].copy()
                df_master_sow.columns = df_master_sow.columns.str.strip()
            except Exception:
                df_master_sow = pd.DataFrame()

            if not query_rows or len(query_rows) <= 1:
                st.warning("⚠️ Belum ada data pada sheet 'Query'.")
            else:
                header = [str(h).strip() for h in query_rows[0]]
                df_query = pd.DataFrame(query_rows[1:], columns=header)

                sow_dropdown_list = (
                    df_dropdown["Master SOW"].dropna().unique().tolist()
                    if "Master SOW" in df_dropdown.columns
                    else []
                )
                sow_dropdown_list = [
                    str(s).strip()
                    for s in sow_dropdown_list
                    if str(s).strip() != ""
                ]
                if not sow_dropdown_list:
                    sow_dropdown_list = [
                        "Survey BSS",
                        "Instalasi BSS",
                        "Instalasi EVC",
                    ]

                col1, col2, col3 = st.columns(3)

                with col1:
                    selected_sow_type = st.selectbox(
                        "Pilih Jenis SOW",
                        options=sow_dropdown_list,
                        key="sow_type_select",
                    )

                    is_survey = "survey" in str(selected_sow_type).lower()
                    auto_pekerjaan = (
                        "Survey Location" if is_survey else "Construction"
                    )
                    target_wo_col_idx = 11 if is_survey else 22

                    if df_query.shape[1] > target_wo_col_idx:
                        raw_wos = pd.Series(
                            df_query.iloc[:, target_wo_col_idx].values.ravel()
                        ).dropna().unique()
                        wo_list = [
                            str(wo).strip()
                            for wo in raw_wos
                            if str(wo).strip() != ""
                            and str(wo).strip().lower() != "nan"
                        ]
                    else:
                        wo_list = []

                    wo_label = (
                        f"Pilih Nomor WO ({'Kolom L - Survey' if is_survey else 'Kolom W - Cons'})"
                    )
                    selected_wo = st.selectbox(
                        wo_label,
                        options=wo_list if wo_list else ["- Tidak ada WO -"],
                        key=f"wo_select_{selected_sow_type}",
                    )

                    proyek_input = st.text_input("Proyek", value="V-Green")
                    pekerjaan_input = st.text_input(
                        "Pekerjaan",
                        value=auto_pekerjaan,
                        key=f"pekerjaan_input_{selected_sow_type}",
                    )

                with col2:
                    mitra_col_name = (
                        df_dropdown.columns[0]
                        if len(df_dropdown.columns) > 0
                        else "Mitra"
                    )
                    mitra_list = (
                        df_dropdown[mitra_col_name].dropna().unique().tolist()
                        if not df_dropdown.empty
                        and mitra_col_name in df_dropdown.columns
                        else []
                    )
                    mitra_list = [
                        str(m).strip()
                        for m in mitra_list
                        if str(m).strip() != ""
                    ]

                    selected_pic = st.selectbox(
                        "Penanggung Jawab (Mitra)",
                        options=mitra_list if mitra_list else ["Edy"],
                    )

                    default_phone = "0851-8259-6296"
                    if (
                        not df_dropdown.empty
                        and len(df_dropdown.columns) > 1
                        and selected_pic
                    ):
                        phone_col_name = df_dropdown.columns[1]
                        matched_row = df_dropdown[
                            df_dropdown[mitra_col_name].astype(str).str.strip()
                            == selected_pic
                        ]
                        if not matched_row.empty:
                            default_phone = str(
                                matched_row.iloc[0][phone_col_name]
                            ).strip()

                    pic_phone = st.text_input(
                        "No. Telepon Penanggung Jawab", value=default_phone
                    )

                with col3:
                    st.markdown("**Kontak Penerbit SPK (PT. CLX):**")
                    st.markdown("1. Wikantiyoso Suyono (0878-8855-0300)")
                    st.markdown("2. Hartono (0818-0690-9317)")

                if selected_wo and selected_wo != "- Tidak ada WO -":
                    filtered_df = df_query[
                        df_query.iloc[:, target_wo_col_idx]
                        .astype(str)
                        .str.strip()
                        == selected_wo
                    ].copy()

                    filtered_df["col_charge"] = (
                        filtered_df.iloc[:, 2]
                        if filtered_df.shape[1] > 2
                        else "-"
                    )
                    filtered_df["col_site"] = (
                        filtered_df.iloc[:, 5]
                        if filtered_df.shape[1] > 5
                        else "-"
                    )
                    filtered_df["col_gmaps"] = (
                        filtered_df.iloc[:, 7]
                        if filtered_df.shape[1] > 7
                        else "-"
                    )
                    filtered_df["col_province"] = (
                        filtered_df.iloc[:, 8]
                        if filtered_df.shape[1] > 8
                        else "-"
                    )
                    filtered_df["col_pic"] = (
                        filtered_df.iloc[:, 10]
                        if filtered_df.shape[1] > 10
                        else "-"
                    )

                    st.markdown("---")
                    st.markdown(
                        f"### 📍 Daftar Site untuk WO: `{selected_wo}` (Sheet"
                        " Query)"
                    )

                    if "Pilih" not in filtered_df.columns:
                        filtered_df.insert(0, "Pilih", True)

                    edited_df = st.data_editor(
                        filtered_df,
                        use_container_width=True,
                        hide_index=True,
                        key=f"site_editor_{selected_wo}",
                    )

                    st.markdown("---")

                    if st.button(
                        "🚀 Generate SPK & Save to Database Sheet",
                        type="primary",
                    ):
                        selected_sites = (
                            edited_df[edited_df["Pilih"] == True]
                            if "Pilih" in edited_df.columns
                            else pd.DataFrame()
                        )

                        if len(selected_sites) == 0:
                            st.error(
                                "❌ Silakan centang minimal satu site terlebih"
                                " dahulu."
                            )
                        else:
                            with st.spinner(
                                "Memproses dokumen PDF SPK & memperbarui"
                                " Database..."
                            ):
                                matched_sow_df = pd.DataFrame()
                                if not df_master_sow.empty:
                                    kode_col = df_master_sow.columns[0]
                                    matched_sow_df = df_master_sow[
                                        df_master_sow[kode_col]
                                        .astype(str)
                                        .str.strip()
                                        .str.lower()
                                        == str(selected_sow_type)
                                        .strip()
                                        .lower()
                                    ]

                                provinces = (
                                    selected_sites["col_province"]
                                    .dropna()
                                    .astype(str)
                                    .str.strip()
                                    .tolist()
                                )
                                unique_provinces = sorted(
                                    list(
                                        set([
                                            p
                                            for p in provinces
                                            if p != "" and p != "-"
                                        ])
                                    )
                                )
                                auto_lokasi = (
                                    ", ".join(unique_provinces)
                                    if unique_provinces
                                    else "Jawa Tengah"
                                )

                                target_db_name = (
                                    "DB SPK Survey"
                                    if is_survey
                                    else "DB SPK Cons"
                                )

                                try:
                                    target_sheet = sh.worksheet(target_db_name)
                                except Exception:
                                    target_sheet = sh.add_worksheet(
                                        title=target_db_name, rows=1000, cols=20
                                    )
                                    target_sheet.append_row([
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
                                    ])

                                existing_rows = target_sheet.get_all_values()
                                if len(existing_rows) > 1:
                                    spk_ids = [
                                        row[1].strip()
                                        for row in existing_rows[1:]
                                        if len(row) > 1 and row[1].strip()
                                    ]
                                    unique_spk_count = len(set(spk_ids))
                                    seq_number = unique_spk_count + 1
                                else:
                                    seq_number = 1

                                no_spk = generate_spk_number(
                                    selected_sow_type, sequence_num=seq_number
                                )

                                spk_metadata = {
                                    "no_spk": no_spk,
                                    "proyek": proyek_input,
                                    "pekerjaan": pekerjaan_input,
                                    "lokasi": auto_lokasi,
                                    "pic_name": selected_pic,
                                    "pic_phone": pic_phone,
                                    "sow_type": selected_sow_type,
                                }

                                no_spk, pdf_bytes = generate_spk_pdf_bytes(
                                    selected_wo,
                                    selected_sites,
                                    spk_metadata,
                                    matched_sow_df,
                                )
                                safe_filename = (
                                    f"SPK_{no_spk.replace('/', '_')}.pdf"
                                )

                                current_date_str = (
                                    datetime.datetime.now().strftime("%d/%m/%Y")
                                )
                                current_time_str = (
                                    datetime.datetime.now().strftime(
                                        "%Y-%m-%d %H:%M:%S"
                                    )
                                )
                                new_rows = []

                                for _, row in selected_sites.iterrows():
                                    site_name_val = row.get("col_site", "-")
                                    charger_type_val = row.get(
                                        "col_charge", "-"
                                    )

                                    # Tambahkan Status default: Pending COO Approval (Kolom 14)
                                    new_rows.append([
                                        current_date_str,
                                        no_spk,
                                        str(selected_wo),
                                        pekerjaan_input,
                                        "PT CLX",
                                        site_name_val,
                                        charger_type_val,
                                        selected_sow_type,
                                        current_time_str,
                                        selected_pic,
                                        "",
                                        "",
                                        "",
                                        "Pending COO Approval",
                                    ])

                                if new_rows:
                                    target_sheet.append_rows(new_rows)

                                st.success(
                                    f"✅ SPK `{no_spk}` berhasil dibuat &"
                                    " dikirim ke COO Dashboard untuk"
                                    " Approval!"
                                )

                                st.download_button(
                                    label="📥 Download File PDF SPK (Draft)",
                                    data=pdf_bytes,
                                    file_name=safe_filename,
                                    mime="application/pdf",
                                    type="primary",
                                )

        except Exception as e:
            st.error(f"❌ Terjadi kesalahan pada Menu Generate SPK: {e}")

    # --------------------------------------------------------------------------
    # MENU 2: TAKE OVER SITE
    # --------------------------------------------------------------------------
    elif main_menu == "🔄 Take Over Site":
        st.subheader("🔄 Menu Take Over Site")
        st.markdown(
            "Pilih **Sheet Target** (Survey / Cons) dan **Site Name** yang akan"
            " di-*Take Over*. Sistem akan otomatis membuat No. SPK Baru dan"
            " mencatat riwayat *Take Over* pada kolom K:M."
        )

        try:
            db_target_type = st.radio(
                "Pilih Kategori Database SPK",
                options=["DB SPK Cons", "DB SPK Survey"],
                horizontal=True,
            )

            target_sheet_to = sh.worksheet(db_target_type)
            to_rows = target_sheet_to.get_all_values()

            if len(to_rows) <= 1:
                st.warning(f"⚠️ Sheet `{db_target_type}` belum memiliki data.")
            else:
                headers = to_rows[0]
                df_to = pd.DataFrame(to_rows[1:], columns=headers)

                col_site_name = next(
                    (c for c in df_to.columns if "site" in c.lower()),
                    "Site Name",
                )
                col_spk_num = next(
                    (
                        c
                        for c in df_to.columns
                        if "no. spk" in c.lower() or "spk" in c.lower()
                    ),
                    "No. SPK",
                )
                col_mitra_name = next(
                    (c for c in df_to.columns if "mitra" in c.lower()), "Mitra"
                )

                site_options = df_to[col_site_name].dropna().unique().tolist()
                site_options = [
                    s for s in site_options if str(s).strip() not in ["", "-"]
                ]

                col_to1, col_to2 = st.columns(2)

                with col_to1:
                    selected_site_to = st.selectbox(
                        "Pilih Site Name yang akan di Take Over",
                        options=(
                            site_options
                            if site_options
                            else ["- Tidak Ada Site -"]
                        ),
                    )

                if selected_site_to and selected_site_to != "- Tidak Ada Site -":
                    matched_to_row = df_to[
                        df_to[col_site_name] == selected_site_to
                    ]

                    old_spk = (
                        matched_to_row.iloc[0][col_spk_num]
                        if not matched_to_row.empty
                        and col_spk_num in matched_to_row.columns
                        else "-"
                    )
                    old_mitra = (
                        matched_to_row.iloc[0][col_mitra_name]
                        if not matched_to_row.empty
                        and col_mitra_name in matched_to_row.columns
                        else "-"
                    )

                    with col_to2:
                        st.info(f"📌 **No SPK Lama (di Take Over):** {old_spk}")
                        st.info(f"👤 **Mitra Lama:** {old_mitra}")

                    st.markdown("### 📝 Informasi Take Over Baru")
                    col_to3, col_to4 = st.columns(2)

                    try:
                        dropdown_sheet = sh.worksheet("Master Dropdown")
                        dropdown_data = dropdown_sheet.get_all_values()
                        df_dropdown = (
                            pd.DataFrame(
                                dropdown_data[1:], columns=dropdown_data[0]
                            )
                            if len(dropdown_data) > 1
                            else pd.DataFrame()
                        )
                        mitra_col = df_dropdown.columns[0]
                        new_mitra_list = (
                            df_dropdown[mitra_col].dropna().unique().tolist()
                        )
                    except Exception:
                        new_mitra_list = []

                    with col_to3:
                        new_mitra = st.selectbox(
                            "Pilih Mitra Baru (Pengambil Alih)",
                            options=(
                                new_mitra_list
                                if new_mitra_list
                                else ["Mitra Baru"]
                            ),
                        )
                        sow_type_to = (
                            "Survey BSS"
                            if "Survey" in db_target_type
                            else "Instalasi BSS"
                        )

                    with col_to4:
                        today_date_str = datetime.datetime.now().strftime(
                            "%d/%m/%Y"
                        )
                        st.text_input(
                            "Tanggal Take Over (Otomatis)",
                            value=today_date_str,
                            disabled=True,
                        )

                    if st.button(
                        "🔥 Proses & Simpan Take Over Site", type="primary"
                    ):
                        with st.spinner(
                            "Memproses Take Over & memperbarui sheet..."
                        ):
                            spk_ids = [
                                row[1].strip()
                                for row in to_rows[1:]
                                if len(row) > 1 and row[1].strip()
                            ]
                            seq_number = len(set(spk_ids)) + 1

                            new_spk_no = generate_spk_number(
                                sow_type_to, sequence_num=seq_number
                            )

                            for idx, row in enumerate(to_rows[1:], start=2):
                                if len(row) > 5 and row[5] == selected_site_to:
                                    target_sheet_to.update_cell(
                                        idx, 11, today_date_str
                                    )
                                    target_sheet_to.update_cell(
                                        idx, 12, old_spk
                                    )
                                    target_sheet_to.update_cell(
                                        idx, 13, new_mitra
                                    )
                                    break

                            st.success(
                                f"🎉 Take Over Berhasil! \n\n• **Site**:"
                                f" {selected_site_to}\n• **No. SPK Baru**:"
                                f" `{new_spk_no}`\n• **Nomor SPK Lama"
                                f" Ter-record**: `{old_spk}`\n• **Mitra"
                                f" Baru**: `{new_mitra}`"
                            )

        except Exception as e:
            st.error(f"❌ Terjadi kesalahan pada Menu Take Over: {e}")

    # --------------------------------------------------------------------------
    # MENU 3: COO APPROVAL DASHBOARD (2-WAY APPROVAL)
    # --------------------------------------------------------------------------
    elif main_menu == "🔒 COO Approval Dashboard":
        st.subheader("🔒 COO Approval Dashboard (Otorisasi SPK)")
        st.markdown(
            "Halaman khusus Manajemen/COO untuk menyetujui (**Approve**) atau"
            " menolak (**Reject**) pengajuan SPK."
        )

        # Form PIN Login COO
        pin_input = st.text_input(
            "Masukkan PIN Khusus COO:",
            type="password",
            placeholder="Masukkan PIN...",
        )

        if pin_input == COO_PIN_SECRET:
            st.success("🔓 Akses Diterima! Selamat datang, COO.")

            db_approval_type = st.radio(
                "Pilih Kategori Database SPK",
                options=["DB SPK Cons", "DB SPK Survey"],
                horizontal=True,
                key="approval_db_choice",
            )

            try:
                sheet_appr = sh.worksheet(db_approval_type)
                appr_rows = sheet_appr.get_all_values()

                if len(appr_rows) <= 1:
                    st.info(f"ℹ️ Belum ada data pada sheet `{db_approval_type}`.")
                else:
                    headers = appr_rows[0]
                    df_appr = pd.DataFrame(appr_rows[1:], columns=headers)

                    # Tampilkan Kolom Status jika belum ada
                    if "Status Approval" not in df_appr.columns:
                        df_appr["Status Approval"] = "Pending COO Approval"

                    # Filter SPK yang membutuhkan persetujuan
                    pending_df = df_appr[
                        df_appr["Status Approval"] == "Pending COO Approval"
                    ]

                    st.markdown("---")
                    st.markdown(
                        f"### 📋 Daftar SPK Menunggu Persetujuan ({len(pending_df)} Pending)"
                    )

                    if pending_df.empty:
                        st.balloons()
                        st.success(
                            "🎉 Semua pengajuan SPK telah diproses! Tidak ada"
                            " antrean pending."
                        )
                    else:
                        # Dapatkan No SPK Unik yang pending
                        pending_spks = (
                            pending_df["No. SPK"].dropna().unique().tolist()
                        )

                        col_app1, col_app2 = st.columns([2, 1])

                        with col_app1:
                            selected_appr_spk = st.selectbox(
                                "Pilih SPK yang akan ditinjau:",
                                options=pending_spks,
                            )

                        if selected_appr_spk:
                            spk_details = pending_df[
                                pending_df["No. SPK"] == selected_appr_spk
                            ]

                            st.markdown("#### 🔍 Detail Informasi SPK:")
                            st.write(
                                spk_details[[
                                    "Date SPK",
                                    "No. SPK",
                                    "No. WO",
                                    "Pekerjaan",
                                    "Site Name",
                                    "Mitra",
                                ]]
                            )

                            col_btn1, col_btn2, _ = st.columns([1, 1, 2])

                            with col_btn1:
                                if st.button(
                                    "✅ APPROVE SPK", type="primary"
                                ):
                                    with st.spinner("Memproses Approval..."):
                                        # Update status di Google Sheet
                                        for idx, row in enumerate(
                                            appr_rows[1:], start=2
                                        ):
                                            if (
                                                len(row) > 1
                                                and row[1] == selected_appr_spk
                                            ):
                                                sheet_appr.update_cell(
                                                    idx, 14, "Approved by COO"
                                                )

                                        st.success(
                                            f"✅ SPK `{selected_appr_spk}`"
                                            " BERHASIL DI-APPROVE!"
                                        )
                                        st.rerun()

                            with col_btn2:
                                if st.button("❌ REJECT SPK"):
                                    with st.spinner("Memproses Penolakan..."):
                                        for idx, row in enumerate(
                                            appr_rows[1:], start=2
                                        ):
                                            if (
                                                len(row) > 1
                                                and row[1] == selected_appr_spk
                                            ):
                                                sheet_appr.update_cell(
                                                    idx, 14, "Rejected by COO"
                                                )

                                        st.warning(
                                            f"❌ SPK `{selected_appr_spk}`"
                                            " DITOLAK!"
                                        )
                                        st.rerun()

            except Exception as e:
                st.error(
                    f"❌ Terjadi kesalahan pada Dashboard Approval COO: {e}"
                )

        elif pin_input != "":
            st.error("❌ PIN Salah! Akses ditolak.")


if __name__ == "__main__":
    show_spk_page()