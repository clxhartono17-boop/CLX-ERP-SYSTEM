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
        update_do_in_db_material_out
    )
except ImportError:
    # Fallback dummy jika terputus
    def generate_do_number(): return "0001/CLX/DO/VII/2026"
    def save_do_to_db_material_out(data): return True
    def get_all_do_numbers(): return []
    def get_do_by_number(no): return None
    def update_do_in_db_material_out(no, data): return True


# ==============================================================================
# FUNGSI HELPER & LOGIC DATA
# ==============================================================================

def load_master_dropdown():
    """Membaca data dari Sheet 'Master Dropdown'"""
    charging_types = ["6S1P", "12S1P", "DC20", "DC30", "DC60", "DC120"]
    expeditions = ["BCE", "Lalamove", "Self Pick Up", "JNE", "TIKI"]
    return charging_types, expeditions

def load_epc_list():
    """Membaca data EPC dari Sheet 'Query' Kolom B"""
    return ["PT Sunrise Internusa", "CLX", "PT Energi Nusantara"]

def load_filtered_sites(epc, charging_type):
    """Filter data site pada Sheet 'Query' Kolom F (Maksimal 15 Site)"""
    all_sites = [
        "Stasiun Madiun", "Stasiun Ngawi", "Stasiun Magetan", "Stasiun Caruban",
        "Stasiun Lamongan", "Stasiun Bojonegoro", "[T1D3] SWADAYA 2",
        "[TAQ5] GOTONG ROYONG", "[TQY8] RAYA BAYUR", "Stasiun Kediri",
        "Stasiun Blitar", "Stasiun Jombang", "Stasiun Kertosono", "Stasiun Tulungagung", "Stasiun Nganjuk"
    ]
    return all_sites[:15]

def load_standard_charging_materials(charging_type):
    """Membaca Sheet 'Standart Charging Type'"""
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
        [Paragraph("Name:", body_style), Paragraph(str(data['to']), body_bold)],
        [Paragraph("Phone No.:", body_style), Paragraph(str(data['contact']), body_style)],
        [Paragraph("Address:", body_style), Paragraph(str(data['address']), body_style)]
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
        [Paragraph("No. DO:", body_bold), Paragraph(str(data['no_do']), body_bold)],
        [Paragraph("Date:", body_style), Paragraph(str(data['date']), body_style)],
        [Paragraph("EPC:", body_style), Paragraph(str(data['epc']), body_style)],
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
        Paragraph("Site Allocation (Max. 15)", header_table_style),
        Paragraph("Remarks", header_table_style)
    ]
    
    mat_rows = [mat_headers]
    for item in data['materials']:
        code = item.get('Material Code', item.get('code', ''))
        name = item.get('Material Name', item.get('name', ''))
        site = item.get('Site Allocation', item.get('Site Alocation', ''))
        uom = item.get('UoM', item.get('uom', 'Pcs'))
        
        mat_rows.append([
            Paragraph(str(item['No']), ParagraphStyle('C', alignment=1, fontSize=6)),
            Paragraph(str(code), body_style),
            Paragraph(str(name), body_style),
            Paragraph(str(item['Qty']), ParagraphStyle('C', alignment=1, fontSize=6)),
            Paragraph(str(uom), ParagraphStyle('C', alignment=1, fontSize=6)),
            Paragraph(str(site), body_style),
            Paragraph(str(item.get('Remarks', '')), body_style)
        ])

    site_allocated_count = data.get('site_count', len([m for m in data['materials'] if m.get('Site Allocation') or m.get('Site Alocation')]))

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

    # Inisialisasi State Histori Relokasi jika belum ada
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
            epc = st.selectbox("3. EPC (Query Sheet)", epc_list)

        with col2:
            charging_type = st.selectbox("4. Charging Type (Master Dropdown)", charging_list)
            expedition = st.selectbox("5. Expedition (Master Dropdown)", exp_list)
            to_name = st.text_input("6. To (Recipient Name)", placeholder="Contoh: sasuke")

        with col3:
            contact = st.text_input("7. Contact (Phone No.)", placeholder="Contoh: 8964254")
            address = st.text_area("8. Address", placeholder="Contoh: konoha", height=110)

        st.divider()
        st.subheader("Filter Site & Kalkulasi Material Automatic")

        available_sites = load_filtered_sites(epc, charging_type)
        selected_sites = st.multiselect(
            "Alokasi Site (Maksimal 15 Site terfilter):",
            options=available_sites,
            default=available_sites[:9],
            max_selections=15
        )
        
        site_count = len(selected_sites)
        st.info(f"📊 Total Site Terpilih: **{site_count} Site Allocated**")

        raw_materials = load_standard_charging_materials(charging_type)
        table_data = []
        for idx, item in enumerate(raw_materials, start=1):
            total_qty = item["std_qty"] * site_count
            allocated_site = selected_sites[idx-1] if idx <= len(selected_sites) else ""
            table_data.append({
                "No": idx,
                "Material Code": item["code"],
                "Material Name": item["name"],
                "Qty": total_qty,
                "UoM": item["uom"],
                "Site Allocation": allocated_site,
                "Remarks": ""
            })

        df_materials = pd.DataFrame(table_data)

        st.subheader("Detail Material & Site Allocation")
        edited_df = st.data_editor(
            df_materials,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "No": st.column_config.NumberColumn(width="small", disabled=True),
                "Material Code": st.column_config.TextColumn(disabled=True),
                "Material Name": st.column_config.TextColumn(disabled=True),
                "Qty": st.column_config.NumberColumn("Qty (Auto calculated)", help="Qty = Std Qty x Total Site"),
                "UoM": st.column_config.TextColumn(disabled=True),
                "Site Allocation": st.column_config.TextColumn("Site Allocation (Max 15)"),
                "Remarks": st.column_config.TextColumn("Remarks")
            }
        )

        st.divider()

        if st.button("🚀 Simpan & Generate Delivery Order", type="primary"):
            if not to_name or not address:
                st.error("Kolom 'To' dan 'Address' wajib diisi!")
            else:
                date_str = do_date.strftime("%Y-%m-%d")
                materials_list = edited_df.to_dict(orient="records")
                
                db_rows = []
                for row in materials_list:
                    db_rows.append({
                        "No": row["No"],
                        "No. DO": no_do,
                        "Delv. Date": date_str,
                        "Material Code": row["Material Code"],
                        "Material Name": row["Material Name"],
                        "Qty": row["Qty"],
                        "uom": row["UoM"],
                        "Site Alocation": row["Site Allocation"],
                        "Remarks": row.get("Remarks", ""),
                        "To": to_name,
                        "Phone No.": contact,
                        "Address": address,
                        "EPC": epc
                    })

                with st.spinner("Menyimpan transaksi ke sheet 'DB Material Out'..."):
                    save_do_to_db_material_out(db_rows)

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
                    "materials": materials_list
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
                e_no_do = st.text_input("No. DO", value=edit_data['no_do'], disabled=True, key="e_no_do")
                e_date = st.text_input("Delivery Date", value=edit_data['date'], key="e_date")
            with ecol2:
                e_to = st.text_input("To (Recipient)", value=edit_data['to'], key="e_to")
                e_contact = st.text_input("Phone No.", value=edit_data['contact'], key="e_contact")
            with ecol3:
                e_epc = st.text_input("EPC", value=edit_data['epc'], key="e_epc")
                e_address = st.text_area("Address", value=edit_data['address'], key="e_address", height=100)

            st.write("**Material Items:**")
            df_edit_mat = pd.DataFrame(edit_data['materials'])
            
            cols_to_show = ["No", "Material Code", "Material Name", "Qty", "uom", "Site Allocation", "Remarks"]
            cols_existing = [c for c in cols_to_show if c in df_edit_mat.columns]
            
            edited_mat_df = st.data_editor(
                df_edit_mat[cols_existing],
                num_rows="dynamic",
                use_container_width=True,
                key="editor_search_do"
            )

            # ==================================================================
            # 🔄 MODUL FITUR RELOKASI SITE MATERIAL
            # ==================================================================
            st.markdown("---")
            st.subheader("🔁 Form Relokasi Material Site")
            st.info("Gunakan modul ini jika ada material yang salah alokasi dan harus dipindahkan ke site lain tanpa menghilangkan histori.")

            with st.expander("📌 Klik di sini untuk Melakukan Relokasi Material", expanded=False):
                # Ambil daftar material yang tersedia
                mat_options = edited_mat_df["Material Name"].tolist() if "Material Name" in edited_mat_df.columns else []
                
                col_r1, col_r2, col_r3 = st.columns(3)
                with col_r1:
                    target_mat_name = st.selectbox("Pilih Material yang Direlokasi:", options=mat_options, key="reloc_mat")
                
                # Cari site asal dari material terpilih
                current_site_val = ""
                if target_mat_name and not edited_mat_df.empty:
                    match_row = edited_mat_df[edited_mat_df["Material Name"] == target_mat_name]
                    if not match_row.empty:
                        col_site = "Site Allocation" if "Site Allocation" in match_row.columns else "Site Alocation"
                        current_site_val = match_row.iloc[0][col_site]

                with col_r2:
                    old_site = st.text_input("Site Asal (Old Site)", value=current_site_val, disabled=True, key="reloc_old_site")
                with col_r3:
                    new_site = st.text_input("Site Tujuan Baru (New Site)", placeholder="Contoh: Stasiun Kediri", key="reloc_new_site")

                reloc_reason = st.text_input("Alasan Relokasi / CatatanTambahan:", placeholder="Contoh: Salah kirim tim lapangan / Revisi WO", key="reloc_reason")

                if st.button("🔀 Eksekusi Relokasi Site", type="secondary"):
                    if not new_site:
                        st.error("Site Tujuan Baru wajib diisi!")
                    elif old_site == new_site:
                        st.warning("Site Asal dan Site Tujuan Baru tidak boleh sama!")
                    else:
                        now_stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
                        
                        # Update dataframe & tambahkan histori di remarks
                        col_site_key = "Site Allocation" if "Site Allocation" in edited_mat_df.columns else "Site Alocation"
                        for idx, row in edited_mat_df.iterrows():
                            if row["Material Name"] == target_mat_name:
                                edited_mat_df.at[idx, col_site_key] = new_site
                                old_rem = str(row.get("Remarks", ""))
                                new_rem = f"[Relokasi {now_stamp}: dari '{old_site}' ke '{new_site}'. Ket: {reloc_reason}] {old_rem}".strip()
                                edited_mat_df.at[idx, "Remarks"] = new_rem

                        # Simpan ke histori relokasi session
                        st.session_state.relocation_history.append({
                            "no_do": e_no_do,
                            "timestamp": now_stamp,
                            "material": target_mat_name,
                            "old_site": old_site,
                            "new_site": new_site,
                            "reason": reloc_reason
                        })

                        st.success(f"✅ Berhasil merelokasi **{target_mat_name}** dari **{old_site}** ke **{new_site}**!")
                        st.rerun()

            # ==================================================================
            # 📜 HISTORI RELOKASI MATERIAL
            # ==================================================================
            # Filter histori relokasi khusus untuk No DO yang sedang dibuka
            do_hist = [h for h in st.session_state.relocation_history if h["no_do"] == e_no_do]
            if do_hist:
                st.markdown("#### 📜 Audit Trail / Histori Relokasi DO Ini")
                df_hist = pd.DataFrame(do_hist)
                st.dataframe(df_hist[["timestamp", "material", "old_site", "new_site", "reason"]], use_container_width=True)

            st.markdown("---")

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("💾 Simpan Perubahan DO & Relokasi ke Database", type="primary"):
                    updated_materials = edited_mat_df.to_dict(orient="records")
                    payload_update = []
                    
                    for row in updated_materials:
                        payload_update.append({
                            "No": row.get("No"),
                            "No. DO": e_no_do,
                            "Delv. Date": e_date,
                            "Material Code": row.get("Material Code"),
                            "Material Name": row.get("Material Name"),
                            "Qty": row.get("Qty"),
                            "uom": row.get("uom", row.get("UoM")),
                            "Site Allocation": row.get("Site Allocation", row.get("Site Alocation")),
                            "Remarks": row.get("Remarks", ""),
                            "To": e_to,
                            "Phone No.": e_contact,
                            "Address": e_address,
                            "EPC": e_epc
                        })

                    with st.spinner("Memperbarui database Google Sheets..."):
                        if update_do_in_db_material_out(e_no_do, payload_update):
                            st.success(f"Berhasil memperbarui {e_no_do} di database!")
                            st.session_state.current_do = {
                                "no_do": e_no_do,
                                "date": e_date,
                                "epc": e_epc,
                                "to": e_to,
                                "contact": e_contact,
                                "address": e_address,
                                "materials": payload_update
                            }
                            st.rerun()

            with col_btn2:
                if st.button("🖨️ Set Ke Preview & Cetak PDF Baru"):
                    st.session_state.current_do = {
                        "no_do": edit_data['no_do'],
                        "date": edit_data['date'],
                        "epc": edit_data['epc'],
                        "to": edit_data['to'],
                        "contact": edit_data['contact'],
                        "address": edit_data['address'],
                        "materials": edited_mat_df.to_dict(orient="records")
                    }
                    st.info("Data DO ini telah diset untuk preview! Buka tab **🖨️ Preview & PDF Cetak (A5)** untuk mengunduh PDF-nya.")

# Alias pemanggilan
show = render
