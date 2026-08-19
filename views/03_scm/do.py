import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os

# Import ReportLab untuk PDF Generator A5
try:
    from reportlab.lib.pagesizes import A5, portrait
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# ==============================================================================
# IMPORT DARI CORE DATABASE (SINKRONISASI DATABASE GOOGLE SHEETS)
# ==============================================================================
try:
    from core.database import (
        generate_do_number,
        save_do_to_db_material_out,
        get_all_do_numbers,
        get_do_by_number,
        update_do_in_db_material_out,
        get_used_sites_from_db_material_out,
        get_query_sheet_data  # Fungsi untuk mengambil data Sheet "Query"
    )
except ImportError:
    # Handler bawaan jika modul core.database belum terhubung
    def generate_do_number(is_reloc=False): 
        return "0002/CLX/DO-RELOC/VIII/2026" if is_reloc else "0001/CLX/DO/VIII/2026"
    def save_do_to_db_material_out(data): return True
    def get_all_do_numbers(): return []
    def get_do_by_number(no): return None
    def update_do_in_db_material_out(no, data): return True
    def get_used_sites_from_db_material_out(): return []
    def get_query_sheet_data(): return pd.DataFrame()


# ==============================================================================
# FUNGSI HELPER & LOGIC DATA SHEET "QUERY" (TANPA FALLBACK DATA)
# ==============================================================================

def load_master_dropdown():
    """Membaca data dari Sheet 'Master Dropdown'"""
    charging_types = ["6S1P", "12S1P", "DC20", "DC30", "DC60", "DC120"]
    expeditions = ["BCE", "Lalamove", "Self Pick Up", "JNE", "TIKI"]
    return charging_types, expeditions

def fetch_raw_query_data():
    """
    Mengambil data Sheet 'Query' secara langsung dari Google Sheets.
    Fallback data dummy telah dihapus total agar jika ada error koneksi,
    pesan peringatan muncul secara transparan.
    """
    try:
        df = get_query_sheet_data()
        if isinstance(df, pd.DataFrame) and not df.empty:
            return df
    except Exception as e:
        st.error(f"⚠️ Gagal mengambil data dari Google Sheets: {e}")
    
    return pd.DataFrame()

def load_epc_list():
    """Membaca daftar unik EPC Name dari Sheet 'Query' Kolom B"""
    df_query = fetch_raw_query_data()
    if not df_query.empty:
        col_epc = 'EPC Name' if 'EPC Name' in df_query.columns else 'epc'
        if col_epc in df_query.columns:
            return sorted(df_query[col_epc].dropna().unique().tolist())
    return []

def load_filtered_sites(epc, charging_type):
    """
    Membaca & Memfilter Sheet 'Query' dengan kriteria:
    1. Kolom D (Project Status) BUKAN 'Drop' atau 'Cancel'
    2. Kolom B (EPC Name) == epc & Kolom C (Charging Type) == charging_type
    3. Mengeliminasi site di Kolom F yang SUDAH ada di Sheet 'DB Material Out' (Kolom U)
    4. MENAMPILKAN SEMUA SITE HASIL FILTERING (Tanpa dibatasi 15 item di dropdown)
    """
    df_query = fetch_raw_query_data()
    
    # Ambil site yang sudah terpakai di DB Material Out (Kolom U)
    try:
        used_sites = get_used_sites_from_db_material_out()
    except Exception:
        used_sites = []

    if df_query.empty:
        return []

    # Map nama kolom (Mendukung DataFrame asli GSheet)
    col_epc = 'EPC Name' if 'EPC Name' in df_query.columns else 'epc'
    col_charging = 'Charging Type' if 'Charging Type' in df_query.columns else 'charging'
    col_status = 'Project Status' if 'Project Status' in df_query.columns else 'status'
    col_site = 'Project / Location Name' if 'Project / Location Name' in df_query.columns else 'site'

    target_epc = str(epc).strip().lower() if epc else ""
    target_charging = str(charging_type).strip().lower() if charging_type else ""

    valid_sites = []
    for idx, row in df_query.iterrows():
        status_val = str(row.get(col_status, '')).strip().lower()
        epc_val = str(row.get(col_epc, '')).strip().lower()
        charging_val = str(row.get(col_charging, '')).strip().lower()
        site_name = str(row.get(col_site, '')).strip()

        # Rule 1: Status BUKAN mengandung kata 'Drop' atau 'Cancel'
        is_active = "drop" not in status_val and "cancel" not in status_val
        # Rule 2: Filter EPC (Kolom B) & Charging Type (Kolom C)
        is_match_epc = (epc_val == target_epc)
        is_match_charging = (charging_val == target_charging)
        # Rule 3: Filter site yang belum ada di DB Material Out
        is_not_used = site_name not in used_sites

        if is_active and is_match_epc and is_match_charging and is_not_used:
            if site_name and site_name not in valid_sites:
                valid_sites.append(site_name)

    # Mengembalikan SELURUH site unik tanpa batasan 15 item
    return list(dict.fromkeys(valid_sites))

def load_available_relocation_sites():
    """
    Membaca Sheet 'Query' Kolom F (Project / Location Name)
    dengan mengeliminasi status Drop dan Cancel.
    """
    df_query = fetch_raw_query_data()
    if df_query.empty:
        return []

    col_status = 'Project Status' if 'Project Status' in df_query.columns else 'status'
    col_site = 'Project / Location Name' if 'Project / Location Name' in df_query.columns else 'site'

    valid_sites = []
    for idx, row in df_query.iterrows():
        status_val = str(row.get(col_status, '')).strip().lower()
        site_name = str(row.get(col_site, '')).strip()
        
        if "drop" not in status_val and "cancel" not in status_val and site_name:
            if site_name not in valid_sites:
                valid_sites.append(site_name)
            
    return valid_sites

def load_standard_charging_materials(charging_type):
    """Membaca Sheet 'Standart Charging Type'"""
    if not charging_type:
        return []

    raw_data = [
        {"code": "AC0001", "name": "Clamp Conduit", "uom": "Pcs", "qty_map": {"6S1P": 5, "12S1P": 5}},
        {"code": "AC0002", "name": "Kabel Schoen 16", "uom": "Pcs", "qty_map": {"6S1P": 10, "12S1P": 10}},
        {"code": "AC0004", "name": "Kabel Vynil Biru", "uom": "Pcs", "qty_map": {"6S1P": 4, "12S1P": 4}},
        {"code": "AC0005", "name": "Kabel Vynil Hijau", "uom": "Pcs", "qty_map": {"6S1P": 2, "12S1P": 2}},
        {"code": "AC0006", "name": "Kabel Vynil Hitam", "uom": "Pcs", "qty_map": {"6S1P": 2, "12S1P": 2}},
        {"code": "AC0008", "name": "Kabel Vynil Merah", "uom": "Pcs", "qty_map": {"6S1P": 4, "12S1P": 4}},
        {"code": "AC0009", "name": "Kuku Macan 10", "uom": "Pcs", "qty_map": {"6S1P": 1, "12S1P": 1}},
        {"code": "AC0010", "name": "Sok Konektor Grounding 5/8\"", "uom": "Pcs", "qty_map": {"DC20": 1, "DC30": 1, "DC60": 1}},
        {"code": "MM0001", "name": "APAR 3Kg", "uom": "Pcs", "qty_map": {"DC20": 1, "DC30": 1, "DC60": 1}},
        {"code": "MM0002", "name": "Box APAR", "uom": "Pcs", "qty_map": {"DC20": 1, "DC30": 1, "DC60": 1}},
        {"code": "MM0004", "name": "Combiner 63A BSS", "uom": "Pcs", "qty_map": {"12S1P": 1}},
        {"code": "MM0005", "name": "Combiner 40A 3P", "uom": "Pcs", "qty_map": {"DC20": 1}},
        {"code": "MM0006", "name": "Combiner 40A BSS", "uom": "Pcs", "qty_map": {"6S1P": 1}},
        {"code": "MM0007", "name": "Combiner 63A", "uom": "Pcs", "qty_map": {"DC30": 1}},
        {"code": "MM0009", "name": "Conduit Anaconda 1\"", "uom": "Pcs", "qty_map": {"6S1P": 10, "12S1P": 10}},
        {"code": "MM0011", "name": "Kabel Grounding 6", "uom": "Pcs", "qty_map": {"6S1P": 5, "12S1P": 5}},
        {"code": "MM0012", "name": "Kabel Power NYY 4x10", "uom": "Pcs", "qty_map": {"DC20": 12}},
        {"code": "MM0013", "name": "Kabel Power NYY 4x16", "uom": "Pcs", "qty_map": {"DC30": 12}},
        {"code": "MM0014", "name": "Kabel Power NYY 4x25mm", "uom": "Pcs", "qty_map": {"DC60": 12}},
        {"code": "MM0016", "name": "Kabel Power NYYHY 3x10", "uom": "Pcs", "qty_map": {"6S1P": 10}},
        {"code": "MM0017", "name": "NYA 10mm", "uom": "Pcs", "qty_map": {"6S1P": 5, "12S1P": 5, "DC20": 15, "DC30": 15}},
        {"code": "MM0018", "name": "NYA 16mm", "uom": "Pcs", "qty_map": {"DC60": 15}},
        {"code": "MM0020", "name": "Stick Rod 2m", "uom": "Pcs", "qty_map": {"6S1P": 1, "12S1P": 1, "DC20": 1, "DC30": 1, "DC60": 1}},
        {"code": "MM0021", "name": "Stick Rod 1.5m", "uom": "Pcs", "qty_map": {"6S1P": 1, "12S1P": 1, "DC20": 1, "DC30": 1, "DC60": 1}},
        {"code": "MM0022", "name": "Stick Rod 1m", "uom": "Pcs", "qty_map": {"6S1P": 1, "12S1P": 1, "DC20": 1, "DC30": 1, "DC60": 1}},
    ]

    result = []
    for item in raw_data:
        std_qty = item["qty_map"].get(charging_type, 0)
        if std_qty > 0:
            result.append({
                "code": item["code"],
                "name": item["name"],
                "std_qty": std_qty,
                "uom": item["uom"]
            })
    return result


def generate_do_a5_pdf(data):
    """Fungsi Generator PDF A5 Fit to Paper dengan Header Logo CLX"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=portrait(A5),
        rightMargin=15,
        leftMargin=15,
        topMargin=15,
        bottomMargin=15
    )
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('T', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor("#1a365d"))
    subtitle_style = ParagraphStyle('ST', fontName='Helvetica', fontSize=6, textColor=colors.HexColor("#4a5568"), leading=7)
    body_style = ParagraphStyle('B', fontName='Helvetica', fontSize=6.5, leading=8, textColor=colors.HexColor("#2d3748"))
    body_bold = ParagraphStyle('BB', fontName='Helvetica-Bold', fontSize=6.5, leading=8, textColor=colors.HexColor("#1a365d"))
    header_table_style = ParagraphStyle('HT', fontName='Helvetica-Bold', fontSize=6.5, textColor=colors.white, alignment=1)

    logo_path = "assets/logo.png"
    if os.path.exists(logo_path):
        logo_img = RLImage(logo_path, width=90, height=25)
    else:
        logo_img = Paragraph("<b>PT. CLX</b>", title_style)

    company_info = [
        Paragraph("<b>PT. Connectivity Leads excellence</b>", title_style),
        Paragraph("Jl. M Ali 2 No. 19 RT 007 RW 004 Tanah Baru, Beji, Kota Depok, Jawa barat 16426", subtitle_style),
        Paragraph("E: clx.central@gmail.com | T: +62 821-4858-1879", subtitle_style)
    ]

    head_table = Table([[logo_img, company_info]], colWidths=[95, 295])
    head_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LINEBELOW', (0,0), (-1,-1), 1, colors.HexColor("#1a365d")),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(head_table)
    elements.append(Spacer(1, 5))

    to_box = [
        [Paragraph("<b>To</b>", body_bold), ""],
        [Paragraph("Name:", body_style), Paragraph(str(data.get('to', '')), body_bold)],
        [Paragraph("Phone No.:", body_style), Paragraph(str(data.get('contact', '')), body_style)],
        [Paragraph("Address:", body_style), Paragraph(str(data.get('address', '')), body_style)]
    ]
    to_table = Table(to_box, colWidths=[45, 140])
    to_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e0")),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#edf2f7")),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))

    meta_box = [
        [Paragraph("<b>DELIVERY ORDER</b>", ParagraphStyle('DO', fontName='Helvetica-Bold', fontSize=8, alignment=1, textColor=colors.HexColor("#1a365d"))), ""],
        [Paragraph("No. DO:", body_bold), Paragraph(str(data.get('no_do', '')), body_bold)],
        [Paragraph("Date:", body_style), Paragraph(str(data.get('date', '')), body_style)],
        [Paragraph("EPC:", body_style), Paragraph(str(data.get('epc', '')), body_style)],
        [Paragraph("Charging Type:", body_style), Paragraph(str(data.get('charging_type', '-')), body_style)],
        [Paragraph("Expedition:", body_style), Paragraph(str(data.get('expedition', '-')), body_style)]
    ]
    meta_table = Table(meta_box, colWidths=[65, 140])
    meta_table.setStyle(TableStyle([
        ('SPAN', (0,0), (1,0)),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e0")),
        ('BACKGROUND', (0,0), (1,0), colors.HexColor("#e2e8f0")),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))

    top_info_table = Table([[to_table, meta_table]], colWidths=[190, 200])
    top_info_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    elements.append(top_info_table)
    elements.append(Spacer(1, 5))

    mat_headers = [
        Paragraph("No", header_table_style),
        Paragraph("Material Code", header_table_style),
        Paragraph("Material Name", header_table_style),
        Paragraph("Qty", header_table_style),
        Paragraph("UoM", header_table_style),
        Paragraph("Site Allocation", header_table_style),
        Paragraph("Remarks", header_table_style)
    ]
    
    mat_rows = [mat_headers]
    for idx, item in enumerate(data['materials'], start=1):
        code = item.get('Material Code', item.get('code', ''))
        name = item.get('Material Name', item.get('name', ''))
        site = item.get('Site Alocation') or item.get('Site Allocation') or item.get('Remarks') or ''
        uom = item.get('UoM') or item.get('uom') or 'Pcs'
        qty = item.get('Qty', 0)
        
        mat_rows.append([
            Paragraph(str(idx), ParagraphStyle('C', alignment=1, fontSize=6)),
            Paragraph(str(code), body_style),
            Paragraph(str(name), body_style),
            Paragraph(str(qty), ParagraphStyle('C', alignment=1, fontSize=6)),
            Paragraph(str(uom), ParagraphStyle('C', alignment=1, fontSize=6)),
            Paragraph(str(site), body_style),
            Paragraph(str(item.get('Remarks', '')), body_style)
        ])

    site_allocated_count = data.get('site_count', len(set([m.get('Site Alocation') or m.get('Site Allocation') or m.get('Remarks') for m in data['materials'] if m.get('Site Alocation') or m.get('Site Allocation') or m.get('Remarks')])))

    mat_rows.append([
        Paragraph("<b>TOTAL SITE</b>", ParagraphStyle('R', fontName='Helvetica-Bold', fontSize=6.5, alignment=2)),
        "", "", "", "",
        Paragraph(f"<b>{site_allocated_count} Site Allocated</b>", ParagraphStyle('L', fontName='Helvetica-Bold', fontSize=6.5)),
        ""
    ])

    materials_table = Table(mat_rows, colWidths=[18, 50, 95, 25, 22, 120, 60])
    materials_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1a365d")),
        ('GRID', (0,0), (-1,-2), 0.5, colors.HexColor("#cbd5e0")),
        ('SPAN', (0, -1), (4, -1)),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#edf2f7")),
        ('BOX', (0,-1), (-1,-1), 0.5, colors.HexColor("#1a365d")),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    elements.append(materials_table)
    elements.append(Spacer(1, 8))

    sign_title = ParagraphStyle('ST', fontName='Helvetica-Bold', fontSize=6.5, alignment=1)
    sign_data = [
        [Paragraph("Prepared By,", sign_title), Paragraph("Approved By,", sign_title), Paragraph("Received By,", sign_title)],
        ["", "", ""],
        ["( ____________________ )", "( ____________________ )", "( ____________________ )"]
    ]
    sign_table = Table(sign_data, colWidths=[130, 130, 130])
    sign_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('BOTTOMPADDING', (0,0), (-1,0), 20),
        ('TOPPADDING', (0,0), (-1,-1), 1),
    ]))
    elements.append(sign_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


# ==============================================================================
# RENDER MAIN PAGE
# ==============================================================================
def render():
    st.markdown("""
        <style>
        div[data-baseweb="input"],
        div[data-baseweb="input"] > div,
        div[data-baseweb="textarea"],
        div[data-baseweb="textarea"] > textarea,
        div[data-baseweb="select"] > div,
        div[data-testid="stTextInput"] input,
        div[data-testid="stTextArea"] textarea,
        div[data-testid="stDateInput"] input {
            background-color: #FFFFFF !important;
            color: #000000 !important;
            -webkit-text-fill-color: #000000 !important;
            border: 1px solid #94A3B8 !important;
            border-radius: 6px !important;
            font-weight: 600 !important;
            opacity: 1 !important;
        }

        div[data-baseweb="popover"], div[data-baseweb="menu"] {
            background-color: #FFFFFF !important;
            color: #000000 !important;
        }
        li[role="option"] {
            color: #000000 !important;
            background-color: #FFFFFF !important;
        }
        li[role="option"]:hover {
            background-color: #E2E8F0 !important;
        }

        ::placeholder, textarea::placeholder, input::placeholder {
            color: #64748B !important;
            -webkit-text-fill-color: #64748B !important;
            opacity: 1 !important;
        }

        span[data-baseweb="tag"] {
            background-color: #3B82F6 !important;
            color: #FFFFFF !important;
        }

        div[data-testid="stDataFrame"], 
        div[data-testid="stDataEditor"],
        .glideDataEditor {
            background-color: #FFFFFF !important;
            color: #000000 !important;
            border: 1px solid #CBD5E1 !important;
        }

        label, p, h1, h2, h3, h4, .stMarkdown {
            color: #0F172A !important;
        }
        </style>
    """, unsafe_allow_html=True)

    st.title("🚚 Delivery Order (DO) Generator")
    st.caption("Divisi Supply Chain Management (SCM) - Create, Print, Search & Relocation DO")

    if "relocation_history" not in st.session_state:
        st.session_state.relocation_history = []

    charging_list, exp_list = load_master_dropdown()
    epc_list = load_epc_list()

    tab_form, tab_preview, tab_search = st.tabs([
        "📝 Form Create DO",
        "🖨️ Preview & PDF Cetak (A5)",
        "🔍 Search, Edit & Relokasi Site"
    ])

    # --------------------------------------------------------------------------
    # TAB 1: FORM CREATE DO
    # --------------------------------------------------------------------------
    with tab_form:
        st.subheader("Header Delivery Order")

        col1, col2, col3 = st.columns(3)
        with col1:
            no_do_auto = generate_do_number()
            no_do = st.text_input("1. No. DO (Auto)", value=no_do_auto, disabled=True)
            do_date = st.date_input("2. Date", datetime.now())
            epc = st.selectbox("3. EPC (Query Sheet)", epc_list if epc_list else ["Pilih EPC..."], index=None, placeholder="Pilih EPC...")

        with col2:
            charging_type = st.selectbox("4. Charging Type (Master Dropdown)", charging_list, index=None, placeholder="Pilih Charging Type...")
            expedition = st.selectbox("5. Expedition (Master Dropdown)", exp_list, index=None, placeholder="Pilih Ekspedisi...")
            to_name = st.text_input("6. To (Recipient Name)", value="", placeholder="Contoh: Tsubasa Ozora")

        with col3:
            contact = st.text_input("7. Contact (Phone No.)", value="", placeholder="Contoh: 081234567890")
            address = st.text_area("8. Address", value="", placeholder="Contoh: Alamat Tujuan", height=110)

        st.divider()
        st.subheader("Filter Site & Kalkulasi Material Automatic")

        # FILTER SITE REAL DARI SHEET "QUERY" (Menampilkan SEMUA Site hasil kriteria)
        if epc and charging_type:
            available_sites = load_filtered_sites(epc, charging_type)
        else:
            available_sites = []

        # UI Streamlit Multiselect: Menampilkan SEMUA Opsi Site, tapi Membatasi Pemilihan Maksimal 15 Site
        selected_sites = st.multiselect(
            "Alokasi Site (Maksimal 15 Site terpilih):",
            options=available_sites,
            default=[],
            max_selections=15,
            placeholder="Pilih Alokasi Site..." if (epc and charging_type) else "⚠️ Silakan pilih EPC dan Charging Type terlebih dahulu..."
        )
        
        site_count = len(selected_sites)
        st.info(f"📊 Total Site Terpilih: **{site_count} Site Allocated** (Maksimal 15 Site)")

        raw_materials = load_standard_charging_materials(charging_type) if charging_type else []
        
        table_data = []
        for idx, item in enumerate(raw_materials, start=1):
            total_qty = item["std_qty"] * (site_count if site_count > 0 else 1)
            table_data.append({
                "No": idx,
                "Material Code": item["code"],
                "Material Name": item["name"],
                "Qty": total_qty,
                "UoM": item["uom"],
                "Remarks": ""
            })

        df_materials = pd.DataFrame(table_data)

        st.subheader("Detail Material Item (Akan Didistribusikan per Site)")
        edited_df = st.data_editor(
            df_materials,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "No": st.column_config.NumberColumn(width="small", disabled=True),
                "Material Code": st.column_config.TextColumn(disabled=True),
                "Material Name": st.column_config.TextColumn(disabled=True),
                "Qty": st.column_config.NumberColumn("Total Qty (Auto calculated)", help="Qty = Std Qty x Total Site"),
                "UoM": st.column_config.TextColumn(disabled=True),
                "Remarks": st.column_config.TextColumn("Remarks")
            }
        )

        st.divider()

        if st.button("🚀 Simpan & Generate Delivery Order", type="primary"):
            if not epc or not charging_type or not expedition:
                st.error("EPC, Charging Type, dan Expedition wajib dipilih!")
            elif not to_name or not address:
                st.error("Kolom 'To' dan 'Address' wajib diisi!")
            elif site_count == 0:
                st.error("Pilih minimal 1 Site Allocation!")
            else:
                date_str = do_date.strftime("%Y-%m-%d")
                
                generated_db_rows = []
                row_counter = 1
                
                # KONVERSI OTOMATIS BERDASARKAN BANYAKNYA SITE & MASUK KE KOLOM A:N SHEET DB MATERIAL OUT
                for site_name in selected_sites:
                    for mat_item in raw_materials:
                        generated_db_rows.append({
                            # KOLOM A:N (DO BARU)
                            "No": row_counter,
                            "No. DO": no_do,
                            "Delv. Date": date_str,
                            "Material Code": mat_item["code"],
                            "Material Name": mat_item["name"],
                            "Qty": mat_item["std_qty"],
                            "UoM": mat_item["uom"],
                            "Charging Type": charging_type,
                            "Site Alocation": site_name,
                            "Remarks": site_name,
                            "To": to_name,
                            "Phone No.": contact,
                            "Address": address,
                            "EPC": epc,
                            
                            # KOLOM O:T (DO RELOKASI - DIBIARKAN KOSONG PADA SAAT INITIAL CREATE)
                            "Date Reloc.": "",
                            "No. DO Reloc.": "",
                            "Qty Reloc.": "",
                            "Site Reloc.": "",
                            "Mitra Reloc.": "",
                            "Remarks Reloc.": ""
                        })
                        row_counter += 1

                with st.spinner("Menyimpan transaksi ke sheet 'DB Material Out'..."):
                    save_do_to_db_material_out(generated_db_rows)

                st.session_state.current_do = {
                    "no_do": no_do,
                    "date": date_str,
                    "epc": epc,
                    "charging_type": charging_type,
                    "expedition": expedition,
                    "to": to_name,
                    "contact": contact,
                    "address": address,
                    "sites": selected_sites,
                    "site_count": site_count,
                    "materials": generated_db_rows
                }
                
                st.success(f"✅ Delivery Order {no_do} berhasil disimpan ke sheet 'DB Material Out'!")
                st.rerun()

    # --------------------------------------------------------------------------
    # TAB 2: PREVIEW & GENERATE PDF
    # --------------------------------------------------------------------------
    with tab_preview:
        st.subheader("Preview PDF Delivery Order (A5 Format)")
        do_data = st.session_state.get("current_do", None)
        
        if not REPORTLAB_AVAILABLE:
            st.error("Library `reportlab` belum terpasang. Jalankan `pip install reportlab` di terminal.")
        elif not do_data:
            st.warning("Belum ada Delivery Order yang dibuat/dipilih. Silakan isi form atau cari DO terlebih dahulu.")
        else:
            pdf_bytes = generate_do_a5_pdf(do_data)
            st.download_button(
                label=f"🖨️ Download Delivery Order A5 ({do_data['no_do'].replace('/', '_')}.pdf)",
                data=pdf_bytes,
                file_name=f"DO_{do_data['no_do'].replace('/', '_')}_A5.pdf",
                mime="application/pdf",
                type="primary"
            )

    # --------------------------------------------------------------------------
    # TAB 3: SEARCH, EDIT & RELOKASI SITE
    # --------------------------------------------------------------------------
    with tab_search:
        st.subheader("🔍 Cari, Edit & Relokasi Site Delivery Order")
        st.caption("Cari DO berdasarkan Nomor DO untuk mengedit data, merelokasi site material, atau melihat histori relokasi.")

        existing_dos = get_all_do_numbers()
        
        col_s1, col_s2 = st.columns([3, 1])
        with col_s1:
            selected_do_search = st.selectbox("Pilih Nomor DO yang Tersimpan:", options=[""] + existing_dos)
        with col_s2:
            st.write("")
            st.write("")
            btn_search = st.button("🔎 Cari DO", type="primary")

        if btn_search and selected_do_search:
            with st.spinner(f"Mencari data {selected_do_search}..."):
                found_data = get_do_by_number(selected_do_search)
                if found_data:
                    st.session_state.edit_do_data = found_data
                    st.success(f"Data {selected_do_search} ditemukan!")
                else:
                    st.error("Data DO tidak ditemukan di database.")

        # Tampilkan Form Edit & Relokasi jika Data Ditemukan
        if "edit_do_data" in st.session_state and st.session_state.edit_do_data:
            edit_data = st.session_state.edit_do_data
            st.divider()
            st.subheader(f"Edit Data DO: {edit_data['no_do']}")

            ecol1, ecol2, ecol3 = st.columns(3)
            with ecol1:
                e_no_do = st.text_input("No. DO", value=edit_data.get('no_do', ''), disabled=True, key="e_no_do")
                e_date = st.text_input("Delivery Date", value=edit_data.get('date', ''), key="e_date")
            with ecol2:
                e_to = st.text_input("To (Recipient)", value=edit_data.get('to', ''), key="e_to")
                e_contact = st.text_input("Phone No.", value=edit_data.get('contact', ''), key="e_contact")
            with ecol3:
                e_epc = st.text_input("EPC", value=edit_data.get('epc', ''), key="e_epc")
                e_address = st.text_area("Address", value=edit_data.get('address', ''), key="e_address", height=100)

            st.write("**Material Items per Site Allocation:**")
            df_edit_mat = pd.DataFrame(edit_data['materials'])
            
            # Tampilkan Tabel Data (A:T)
            cols_to_show = ["No", "No. DO", "Delv. Date", "Material Code", "Material Name", "Qty", "UoM", "Charging Type", "Site Alocation", "Remarks", 
                            "To", "Phone No.", "Address", "EPC", "Date Reloc.", "No. DO Reloc.", "Qty Reloc.", "Site Reloc.", "Mitra Reloc.", "Remarks Reloc."]
            
            cols_existing = [c for c in cols_to_show if c in df_edit_mat.columns]
            
            edited_mat_df = st.data_editor(
                df_edit_mat[cols_existing],
                num_rows="dynamic",
                use_container_width=True,
                key="editor_search_do"
            )

            # ==================================================================
            # MODUL FITUR RELOKASI SITE MATERIAL (PRESISI DATABASE KOLOM O:T)
            # ==================================================================
            st.markdown("---")
            st.subheader("🔁 Form Eksekusi Relokasi Site Material")
            st.info("Fitur ini akan memperbarui Kolom O:T pada DO Asal dan otomatis membuat baris DO Relokasi Baru di DB.")

            with st.expander("📌 Klik di sini untuk Melakukan Relokasi Site", expanded=True):
                
                # 1. AMBIL SITE ASAL DARI DO INI
                raw_sites_in_do = []
                if "Site Alocation" in edited_mat_df.columns:
                    raw_sites_in_do.extend(edited_mat_df["Site Alocation"].dropna().astype(str).tolist())
                if "Site Allocation" in edited_mat_df.columns:
                    raw_sites_in_do.extend(edited_mat_df["Site Allocation"].dropna().astype(str).tolist())
                if "Remarks" in edited_mat_df.columns:
                    raw_sites_in_do.extend(edited_mat_df["Remarks"].dropna().astype(str).tolist())

                current_do_sites = []
                for s in raw_sites_in_do:
                    clean_s = s.strip()
                    if clean_s and clean_s not in current_do_sites and not clean_s.isdigit() and clean_s != "None":
                        current_do_sites.append(clean_s)

                # 2. AMBIL SITE QUERY & TANGKAP SEMUA SITE YANG SUDAH PERNAH MEMILIKI DO
                all_query_sites = load_available_relocation_sites()
                used_sites_set = set(current_do_sites)
                selectable_new_sites = [s for s in all_query_sites if s not in used_sites_set]

                col_r1, col_r2 = st.columns(2)
                with col_r1:
                    selected_site_old = st.selectbox(
                        "Pilih Site Asal yang Ingin Direlokasi:", 
                        options=current_do_sites if current_do_sites else ["Tidak Ada Site"], 
                        key="reloc_old_site"
                    )
                with col_r2:
                    selected_site_new = st.selectbox(
                        "Nama Site Tujuan Baru (New Site):", 
                        options=selectable_new_sites if selectable_new_sites else ["Tidak ada site baru yang tersedia"], 
                        key="reloc_new_site"
                    )

                reloc_mitra = st.text_input("Mitra Relokasi:", placeholder="Contoh: PT Mitra Jaya", key="reloc_mitra")
                reloc_remarks = st.text_input("Alasan / Catatan Relokasi:", placeholder="Contoh: Perubahan WO Lapangan / Re-alloc Site", key="reloc_reason")

                if st.button("🔀 Eksekusi Relokasi Site", type="secondary"):
                    if selected_site_old == "Tidak Ada Site":
                        st.error("Site asal tidak ditemukan!")
                    elif not selected_site_new or selected_site_new == "Tidak ada site baru yang tersedia":
                        st.error("Silakan pilih Site Tujuan Baru yang valid!")
                    elif selected_site_old == selected_site_new:
                        st.warning("Site Asal dan Site Tujuan Baru tidak boleh sama!")
                    else:
                        reloc_date = datetime.now().strftime("%Y-%m-%d")
                        reloc_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
                        
                        try:
                            new_reloc_do_num = generate_do_number(is_reloc=True)
                        except TypeError:
                            new_reloc_do_num = f"0002/CLX/DO-RELOC/VIII/2026"

                        new_reloc_rows_to_save = []

                        # UPDATE KOLOM O:T PADA DO ASAL
                        for idx, row in edited_mat_df.iterrows():
                            site_in_row = str(row.get("Site Alocation") or row.get("Site Allocation") or row.get("Remarks") or "").strip()
                            
                            if site_in_row == str(selected_site_old).strip():
                                # ISI PRESISI KOLOM O:T Pada Sheet DB Material Out untuk DO Lama
                                edited_mat_df.at[idx, "Date Reloc."] = reloc_date          # Kolom O
                                edited_mat_df.at[idx, "No. DO Reloc."] = new_reloc_do_num # Kolom P
                                edited_mat_df.at[idx, "Qty Reloc."] = row["Qty"]           # Kolom Q
                                edited_mat_df.at[idx, "Site Reloc."] = selected_site_new  # Kolom R
                                edited_mat_df.at[idx, "Mitra Reloc."] = reloc_mitra        # Kolom S
                                edited_mat_df.at[idx, "Remarks Reloc."] = reloc_remarks    # Kolom T

                                # Buat Row Baru Lengkap (Kolom A:N) untuk DO Relokasi Baru
                                new_row = row.copy()
                                new_row["No. DO"] = new_reloc_do_num
                                new_row["Delv. Date"] = reloc_date
                                new_row["Material Code"] = row.get("Material Code", "")
                                new_row["Material Name"] = row.get("Material Name", "")
                                new_row["Qty"] = row.get("Qty", 0)
                                new_row["UoM"] = row.get("UoM") or row.get("uom") or "Pcs"
                                new_row["Charging Type"] = row.get("Charging Type", "")
                                new_row["Site Alocation"] = selected_site_new
                                new_row["Remarks"] = selected_site_new
                                new_row["To"] = e_to
                                new_row["Phone No."] = e_contact
                                new_row["Address"] = e_address
                                new_row["EPC"] = e_epc
                                
                                # Clear Kolom O:T untuk Baris DO Relokasi Baru
                                new_row["Date Reloc."] = ""
                                new_row["No. DO Reloc."] = ""
                                new_row["Qty Reloc."] = ""
                                new_row["Site Reloc."] = ""
                                new_row["Mitra Reloc."] = ""
                                new_row["Remarks Reloc."] = ""
                                
                                new_reloc_rows_to_save.append(new_row.to_dict())

                        # Simpan pembaruan Kolom O:T DO Asal & Simpan Baris DO Relokasi Baru ke DB
                        update_do_in_db_material_out(e_no_do, edited_mat_df.to_dict(orient="records"))
                        save_do_to_db_material_out(new_reloc_rows_to_save)

                        # Simpan Histori Lokal
                        st.session_state.relocation_history.append({
                            "no_do": e_no_do,
                            "timestamp": reloc_timestamp,
                            "old_site": selected_site_old,
                            "new_site": selected_site_new,
                            "reloc_do": new_reloc_do_num,
                            "reason": reloc_remarks
                        })

                        st.success(f"✅ Relokasi Berhasil! Kolom O:T DO Asal telah diisi & Terbuat DO Relokasi Baru: **{new_reloc_do_num}**")
                        st.rerun()

            # ------------------------------------------------------------------
            # HISTORI RELOKASI MATERIAL
            # ------------------------------------------------------------------
            do_hist = [h for h in st.session_state.relocation_history if h["no_do"] == e_no_do]
            if do_hist:
                st.markdown("#### 📜 Audit Trail / Histori Relokasi DO Ini")
                df_hist = pd.DataFrame(do_hist)
                st.dataframe(df_hist[["timestamp", "old_site", "new_site", "reloc_do", "reason"]], use_container_width=True)

            st.markdown("---")

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("💾 Simpan Perubahan Edit Manual DO", type="primary"):
                    updated_materials = edited_mat_df.to_dict(orient="records")
                    with st.spinner("Memperbarui database Google Sheets..."):
                        if update_do_in_db_material_out(e_no_do, updated_materials):
                            st.success(f"Berhasil memperbarui {e_no_do} di database!")
                            st.session_state.current_do = {
                                "no_do": e_no_do,
                                "date": e_date,
                                "epc": e_epc,
                                "to": e_to,
                                "contact": e_contact,
                                "address": e_address,
                                "materials": updated_materials
                            }
                            st.rerun()

            with col_btn2:
                if st.button("🖨️ Set Ke Preview & Cetak PDF Baru"):
                    st.session_state.current_do = {
                        "no_do": edit_data.get('no_do', ''),
                        "date": edit_data.get('date', ''),
                        "epc": edit_data.get('epc', ''),
                        "to": edit_data.get('to', ''),
                        "contact": edit_data.get('contact', ''),
                        "address": edit_data.get('address', ''),
                        "materials": edited_mat_df.to_dict(orient="records")
                    }
                    st.info("Data DO ini telah diset untuk preview! Buka tab **🖨️ Preview & PDF Cetak (A5)** untuk mengunduh PDF-nya.")

# Alias pemanggilan
show = render
