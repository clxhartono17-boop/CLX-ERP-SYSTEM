import os
import io
import base64
import requests
import streamlit as st
from datetime import datetime
from PIL import Image

# ReportLab Imports
from reportlab.lib.pagesizes import A5, portrait
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Core Database Integration
from core.database import (
    generate_reimbursement_no,
    save_reimbursement_to_sheet,
    get_all_reimbursements,
    update_reimbursement_status,
    upload_image_to_gdrive,
)


# ==============================================================================
# UTILITY HELPER
# ==============================================================================

def compress_image(upload_file, max_size=(600, 600), quality=50):
    """Mengompresi foto nota/struk agar ukuran file optimal dan kecil."""
    img = Image.open(upload_file)
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    img.thumbnail(max_size, Image.Resampling.LANCZOS)
    
    output = io.BytesIO()
    img.save(output, format="JPEG", quality=quality, optimize=True)
    return output.getvalue()

def fix_gdrive_url(url):
    """Membersihkan dan memperbaiki format URL Google Drive."""
    if not url:
        return ""
    str_url = str(url).strip().replace(" ", "")
    return str_url

def render_secure_image(url, caption=None, width=None, use_container_width=False):
    """Fungsi helper untuk mengunduh dan merender gambar secara aman di Streamlit."""
    clean_url = fix_gdrive_url(url)
    if not clean_url:
        return
    
    try:
        response = requests.get(clean_url, timeout=5)
        if response.status_code == 200:
            img_bytes = io.BytesIO(response.content)
            if width:
                st.image(img_bytes, caption=caption, width=width)
            else:
                st.image(img_bytes, caption=caption, use_container_width=use_container_width)
        else:
            st.image(clean_url, caption=caption, use_container_width=use_container_width)
    except Exception:
        try:
            st.image(clean_url, caption=caption, use_container_width=use_container_width)
        except Exception:
            st.caption(f"*(Gagal memuat gambar)*")


# ==============================================================================
# PDF GENERATOR FUNCTION (SINGLE PAGE A5)
# ==============================================================================

def generate_a5_reimbursement_pdf(d, raw_imgs_bytes=None):
    """Generator PDF A5 Form Reimbursement (1 Halaman Penuh Termasuk Lampiran)"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=portrait(A5),
        rightMargin=12,
        leftMargin=12,
        topMargin=12,
        bottomMargin=12
    )
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('DocTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9.5, textColor=colors.HexColor("#1a365d"), alignment=2)
    subtitle_style = ParagraphStyle('SubTitle', parent=styles['Normal'], fontName='Helvetica', fontSize=5.5, textColor=colors.HexColor("#718096"))
    cell_bold = ParagraphStyle('CB', fontName='Helvetica-Bold', fontSize=7, textColor=colors.HexColor("#2d3748"))
    cell_norm = ParagraphStyle('CN', fontName='Helvetica', fontSize=7, textColor=colors.HexColor("#2d3748"))
    cell_header = ParagraphStyle('CH', fontName='Helvetica-Bold', fontSize=7, textColor=colors.white, alignment=1)
    cell_right = ParagraphStyle('CR', fontName='Helvetica', fontSize=7, alignment=2)
    cell_right_bold = ParagraphStyle('CRB', fontName='Helvetica-Bold', fontSize=7, alignment=2)
    
    logo_path = "assets/logo.png"
    if os.path.exists(logo_path):
        logo_cell = RLImage(logo_path, width=110, height=35)
    else:
        logo_cell = Paragraph("<b>PT. CLX</b>", ParagraphStyle('L', fontName='Helvetica-Bold', fontSize=11, textColor=colors.HexColor("#1a365d")))
    
    header_text = [
        Paragraph("<b>FORM REIMBURSEMENT</b>", title_style),
        Paragraph("Internal ERP System - Commercial Division", ParagraphStyle('Sub', parent=subtitle_style, alignment=2))
    ]
    
    header_table = Table([[logo_cell, header_text]], colWidths=[115, 260])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LINEBELOW', (0,0), (-1,-1), 1.2, colors.HexColor("#1a365d")),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 3))

    meta_data = [
        [Paragraph("<b>No. Form</b>", cell_bold), Paragraph(f"<font color='#2b6cb0'><b>{d['form_no']}</b></font>", cell_norm), Paragraph("<b>Tanggal</b>", cell_bold), Paragraph(d['date'], cell_norm)],
        [Paragraph("<b>Nama Pemohon</b>", cell_bold), Paragraph(d['pic'], cell_norm), Paragraph("<b>Status</b>", cell_bold), Paragraph(d['status'], cell_norm)],
    ]
    meta_table = Table(meta_data, colWidths=[65, 155, 45, 110])
    meta_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8fafc")),
        ('BOX', (0,0), (-1,-1), 0.4, colors.HexColor("#e2e8f0")),
        ('INNERGRID', (0,0), (-1,-1), 0.4, colors.HexColor("#e2e8f0")),
        ('TOPPADDING', (0,0), (-1,-1), 1.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
    ]))
    elements.append(meta_table)
    elements.append(Spacer(1, 3))

    table_content = [
        [Paragraph("No", cell_header), Paragraph("Deskripsi Keperluan", cell_header), Paragraph("Qty", cell_header), Paragraph("Biaya (Rp)", cell_header), Paragraph("Total (Rp)", cell_header)]
    ]

    for item in d.get("items", []):
        table_content.append([
            Paragraph(str(item['no']), ParagraphStyle('C', alignment=1, fontSize=7)),
            Paragraph(item['description'], cell_norm),
            Paragraph(str(item['qty']), ParagraphStyle('C', alignment=1, fontSize=7)),
            Paragraph(f"{item['amount']:,.0f}", cell_right),
            Paragraph(f"<b>{item['total']:,.0f}</b>", cell_right)
        ])

    table_content.append([
        Paragraph("<b>TOTAL REIMBURSEMENT:</b>", ParagraphStyle('GT', fontName='Helvetica-Bold', fontSize=7, alignment=2)), "", "", "", 
        Paragraph(f"<b>Rp {d['grand_total']:,.0f}</b>", cell_right_bold)
    ])

    data_table = Table(table_content, colWidths=[18, 185, 25, 75, 72])
    data_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2b6cb0")),
        ('GRID', (0,0), (-1,-2), 0.4, colors.HexColor("#cbd5e0")),
        ('SPAN', (0, -1), (3, -1)),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#ebf8ff")),
        ('BOX', (0, -1), (-1, -1), 0.4, colors.HexColor("#2b6cb0")),
        ('TOPPADDING', (0,0), (-1,-1), 1.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
    ]))
    elements.append(data_table)
    elements.append(Spacer(1, 3))

    remarks_text = d['remarks'] if d['remarks'] else '-'
    remarks_p = Paragraph(f"<b>Catatan:</b> {remarks_text}", ParagraphStyle('R', fontName='Helvetica', fontSize=6.5, leading=8))
    remarks_table = Table([[remarks_p]], colWidths=[375])
    remarks_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#fffaf0")),
        ('BOX', (0,0), (-1,-1), 0.4, colors.HexColor("#cbd5e0")),
        ('TOPPADDING', (0,0), (-1,-1), 1.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
    ]))
    elements.append(remarks_table)
    elements.append(Spacer(1, 4))

    sign_style_title = ParagraphStyle('ST', fontName='Helvetica-Bold', fontSize=7, alignment=1)
    sign_style_name = ParagraphStyle('SN', fontName='Helvetica-Bold', fontSize=7, alignment=1)
    sign_style_sub = ParagraphStyle('SS', fontName='Helvetica', fontSize=6, textColor=colors.HexColor("#718096"), alignment=1)

    coo_sig_path = "assets/Approved COO.png"
    cfo_sig_path = "assets/Approved CFO.png"

    col0_sig = ""
    col1_sig = ""
    col2_sig = ""

    if str(d.get("status_coo", "")).strip().lower() == "approved" and os.path.exists(coo_sig_path):
        col1_sig = RLImage(coo_sig_path, width=45, height=22)

    if str(d.get("status_cfo", "")).strip().lower() == "approved" and os.path.exists(cfo_sig_path):
        col2_sig = RLImage(cfo_sig_path, width=45, height=22)

    approval_data = [
        [Paragraph("Diajukan Oleh,", sign_style_title), Paragraph("Disetujui Oleh,", sign_style_title), Paragraph("Disetujui Oleh,", sign_style_title)],
        [col0_sig, col1_sig, col2_sig],  
        [Paragraph(d['pic'], sign_style_name), Paragraph("Chief Operating Officer", sign_style_name), Paragraph("Chief Finance Officer", sign_style_name)],
        [Paragraph("Pemohon", sign_style_sub), Paragraph("COO", sign_style_sub), Paragraph("CFO", sign_style_sub)]
    ]
    
    approval_table = Table(approval_data, colWidths=[125, 125, 125])
    approval_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,1), (2,1), 'CENTER'),
        ('LINEABOVE', (0,2), (0,2), 0.4, colors.HexColor("#718096")),
        ('LINEABOVE', (1,2), (1,2), 0.4, colors.HexColor("#718096")),
        ('LINEABOVE', (2,2), (2,2), 0.4, colors.HexColor("#718096")),
        ('BOTTOMPADDING', (0,0), (-1,0), 1.5),
        ('TOPPADDING', (0,1), (-1,1), 1),
        ('BOTTOMPADDING', (0,1), (-1,1), 1),
        ('TOPPADDING', (0,2), (-1,-1), 1),
        ('BOTTOMPADDING', (0,2), (-1,-1), 1),
    ]))
    elements.append(approval_table)

    imgs_source = raw_imgs_bytes or []
    if not imgs_source:
        for item_data in d.get("items", []):
            img_data = item_data.get("evident", "")
            if not img_data:
                continue
            try:
                clean_img_url = fix_gdrive_url(img_data)
                if clean_img_url.startswith("data:image"):
                    header, encoded = clean_img_url.split(",", 1)
                    imgs_source.append(base64.b64decode(encoded))
                elif clean_img_url.startswith("http"):
                    res = requests.get(clean_img_url, timeout=5)
                    if res.status_code == 200:
                        imgs_source.append(res.content)
            except Exception:
                continue

    if imgs_source:
        elements.append(Spacer(1, 4))
        elements.append(Paragraph("<b>LAMPIRAN BUKTI NOTA / STRUK:</b>", ParagraphStyle('ImgHeader', fontName='Helvetica-Bold', fontSize=8, textColor=colors.HexColor("#1a365d"))))
        elements.append(Spacer(1, 2))

        img_cells = []
        for raw_img in imgs_source:
            try:
                img_stream = io.BytesIO(raw_img)
                pil_img = Image.open(img_stream)
                orig_w, orig_h = pil_img.size
                
                max_w = 175.0
                max_h = 130.0
                ratio = min(max_w / orig_w, max_h / orig_h)
                
                new_w = orig_w * ratio
                new_h = orig_h * ratio

                img_stream.seek(0)
                rl_img = RLImage(img_stream, width=new_w, height=new_h)
                img_cells.append(rl_img)
            except Exception:
                continue

        if img_cells:
            grid_data = []
            for i in range(0, len(img_cells), 2):
                row = img_cells[i:i+2]
                if len(row) == 1:
                    row.append("")
                grid_data.append(row)

            img_table = Table(grid_data, colWidths=[185, 185])
            img_table.setStyle(TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 2),
                ('TOPPADDING', (0,0), (-1,-1), 2),
            ]))
            elements.append(img_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


# ==============================================================================
# MAIN WORKFLOW APP
# ==============================================================================

def render_reimbursement_page():
    st.title("💸 Form Reimbursement")
    st.caption("Commercial Division ERP - EV Charging Infrastructure Platform")

    # Ambil data user yang sedang login dari session state aplikasi utama (app.py)
    current_user = st.session_state.get("user_data", {})
    current_user_email = st.session_state.get("user_email", "-")
    user_name_label = current_user.get("name", "User")
    user_role_label = current_user.get("role", "User")

    # Mengecek hak approval COO & CFO berdasarkan baris data di sheet Otorisasi (via app.py)
    # Catatan: Kolom J (Approval COO) dan K (Approval CFO) jika ingin ditambahkan di load_users_from_sheets() app.py.
    # Atau kita baca langsung dari row jika sudah di-mapping. Sebagai alternatif aman, kita cek dari role atau variabel tambahan:
    is_super_admin = user_role_label in ["Super Admin", "Admin"]
    is_coo_role = user_role_label == "COO" or current_user_email == "clx.wikan18@gmail.com"
    
    # Hak akses approval langsung merujuk pada role/otorisasi akun login
    can_coo = is_super_admin or is_coo_role or user_role_label.lower() == "coo"
    can_cfo = is_super_admin or user_role_label.lower() == "cfo"

    # Sidebar Informasi Sesi Login Aktif
    with st.sidebar:
        st.markdown("---")
        st.subheader("🔐 Status Otorisasi Sesi")
        st.write(f"**Email:** `{current_user_email}`")
        st.write(f"**Nama:** {user_name_label}")
        st.write(f"**Role:** `{user_role_label}`")
        st.markdown(f"- Hak Approval COO: `{'✅ Ya' if can_coo else '❌ Tidak'}`")
        st.markdown(f"- Hak Approval CFO: `{'✅ Ya' if can_cfo else '❌ Tidak'}`")

    if "reimb_items" not in st.session_state:
        st.session_state.reimb_items = []

    tab1, tab2 = st.tabs(["📝 Form Pengajuan", "📑 Riwayat & Approval Workflow"])

    # --------------------------------------------------------------------------
    # TAB 1: FORM PENGAJUAN
    # --------------------------------------------------------------------------
    with tab1:
        st.subheader("Buat Pengajuan Reimbursement Baru")
        
        c1, c2, c3 = st.columns([2, 2, 2])
        with c1:
            form_no_auto = generate_reimbursement_no()
            st.text_input("Nomor Form (Auto-Generated)", value=form_no_auto, disabled=True, key="form_no_display")
        with c2:
            default_name = user_name_label if user_name_label else "Hartono"
            pic_name = st.text_input("Nama Pemohon (PIC)", value=default_name, key="reimb_pic_input")
        with c3:
            reimb_date = st.date_input("Tanggal Pengajuan", value=datetime.now(), key="reimb_date_picker")

        st.markdown("---")
        st.write("##### 🛒 Detail Item Reimbursement & Bukti Nota Mandiri")

        with st.form("reimbursement_main_form"):
            if not st.session_state.reimb_items:
                st.session_state.reimb_items.append({"description": "", "qty": 1, "amount": 0.0})

            updated_items = []
            uploaded_files_dict = {}

            for idx, item in enumerate(st.session_state.reimb_items):
                st.markdown(f"**Item Ke-{idx+1}**")
                col_a, col_b, col_c = st.columns([3, 1.2, 2])
                with col_a:
                    desc = st.text_input(
                        f"Deskripsi Item #{idx+1}",
                        value=item.get("description", ""),
                        key=f"reimb_desc_{idx}",
                        placeholder="Contoh: Bensin, Tiket Tol, Material Kabel"
                    )
                with col_b:
                    qty = st.number_input(
                        f"Qty #{idx+1}",
                        min_value=1,
                        value=int(item.get("qty", 1)),
                        step=1,
                        key=f"reimb_qty_{idx}"
                    )
                with col_c:
                    amount = st.number_input(
                        f"Biaya Satuan #{idx+1}",
                        min_value=0.0,
                        value=float(item.get("amount", 0.0)),
                        step=1000.0,
                        format="%.0f",
                        key=f"reimb_amt_{idx}"
                    )

                up_file = st.file_uploader(
                    f"📷 Upload Bukti Nota / Evident untuk Item #{idx+1} (Format: JPG, PNG)",
                    type=["jpg", "jpeg", "png"],
                    key=f"reimb_file_widget_{idx}"
                )
                if up_file is not None:
                    uploaded_files_dict[idx] = up_file

                updated_items.append({"description": desc, "qty": qty, "amount": amount})
                st.markdown("---")

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                add_clicked = st.form_submit_button("➕ Tambah Item Keperluan", use_container_width=True)
            with col_btn2:
                del_clicked = st.form_submit_button("🗑️ Hapus Item Terakhir", use_container_width=True)

            temp_grand_total = sum(float(item["qty"]) * float(item["amount"]) for item in updated_items)

            col_bottom_1, col_bottom_2 = st.columns([2, 1])
            with col_bottom_1:
                remarks = st.text_area("Catatan / Keterangan Tambahan", key="reimb_remarks", placeholder="Tambahkan catatan jika ada...")
            with col_bottom_2:
                st.markdown("**Total Keseluruhan (Rp):**")
                st.markdown(f"### `Rp {temp_grand_total:,.0f}`")

            submitted = st.form_submit_button("🚀 Submit Pengajuan Reimbursement", use_container_width=True)

        if add_clicked:
            st.session_state.reimb_items = updated_items
            st.session_state.reimb_items.append({"description": "", "qty": 1, "amount": 0.0})
            st.rerun()

        if del_clicked:
            if len(st.session_state.reimb_items) > 1:
                st.session_state.reimb_items = updated_items
                st.session_state.reimb_items.pop()
                st.rerun()

        if submitted:
            st.session_state.reimb_items = updated_items
            grand_total = 0.0
            calculated_items = []
            all_compressed_bytes = []
            image_web_links = []

            has_error = False
            if not pic_name.strip():
                st.error("⚠️ Nama Pemohon (PIC) wajib diisi!")
                has_error = True

            for i, item in enumerate(st.session_state.reimb_items):
                if not item["description"].strip():
                    st.error(f"⚠️ Deskripsi pada Item #{i+1} wajib diisi!")
                    has_error = True

            if not has_error:
                for i, item in enumerate(st.session_state.reimb_items):
                    tot = float(item["qty"]) * float(item["amount"])
                    grand_total += tot
                    
                    evident_link = ""
                    uploaded_file = uploaded_files_dict.get(i)
                    
                    if uploaded_file is not None:
                        try:
                            uploaded_file.seek(0)
                            comp_bytes = compress_image(uploaded_file)
                            all_compressed_bytes.append(comp_bytes)
                            
                            date_str_folder = reimb_date.strftime("%Y-%m-%d")
                            uploaded_data = upload_image_to_gdrive(
                                file_bytes=comp_bytes,
                                file_name=f"{form_no_auto.replace('/', '_')}_item{i+1}_{uploaded_file.name}",
                                pic_name=pic_name,
                                date_str=date_str_folder
                            )
                            
                            if uploaded_data:
                                evident_link = fix_gdrive_url(uploaded_data)
                                image_web_links.append(evident_link)
                        except Exception as e:
                            st.warning(f"⚠️ Gagal mengunggah gambar pada Item #{i+1}: {e}")

                    calculated_items.append({
                        "no": i + 1,
                        "description": item["description"],
                        "qty": item["qty"],
                        "amount": item["amount"],
                        "total": tot,
                        "evident": evident_link
                    })

                if grand_total <= 0:
                    st.error("⚠️ Total biaya reimbursement tidak boleh 0!")
                else:
                    payload_db = {
                        "form_no": form_no_auto,
                        "pic": pic_name,
                        "date": reimb_date.strftime("%d-%b-%Y").upper(),
                        "remarks": remarks if remarks else "-",
                        "status_coo": "Pending",
                        "status_cfo": "Pending",
                        "status": "Pending COO",
                        "items": calculated_items,
                        "grand_total": grand_total,
                        "image_links": image_web_links
                    }

                    if save_reimbursement_to_sheet(payload_db):
                        st.success(f"✅ Pengajuan **{form_no_auto}** berhasil disimpan!")
                        
                        pdf_bytes = generate_a5_reimbursement_pdf(payload_db, raw_imgs_bytes=all_compressed_bytes)
                        st.download_button(
                            label="📄 Download Form Reimbursement + Lampiran Struk (PDF)",
                            data=pdf_bytes,
                            file_name=f"Reimbursement_{form_no_auto.replace('/', '_')}.pdf",
                            mime="application/pdf",
                            key="dl_btn_submit"
                        )
                        
                        st.session_state.reimb_items = [{"description": "", "qty": 1, "amount": 0.0}]
                        for key in list(st.session_state.keys()):
                            if key.startswith("reimb_") or key.startswith("form_no_"):
                                del st.session_state[key]
                        st.rerun()

    # --------------------------------------------------------------------------
    # TAB 2: RIWAYAT & APPROVAL WORKFLOW BERBASIS SESI LOGIN
    # --------------------------------------------------------------------------
    with tab2:
        st.subheader("Daftar Riwayat & Persetujuan Berdasarkan Sesi Login")
        
        db_reimbursements = get_all_reimbursements()
        
        if not db_reimbursements:
            st.info("Belum ada data pengajuan reimbursement yang tersimpan.")
        else:
            for reimb in reversed(db_reimbursements):
                with st.expander(f"📌 **{reimb['form_no']}** | {reimb['pic']} | Total: **Rp {reimb['grand_total']:,.0f}** | Status: `{reimb['status']}`"):
                    c1, c2 = st.columns([3, 2])
                    with c1:
                        st.write(f"**Tanggal:** {reimb['date']}")
                        st.write(f"**Pemohon:** {reimb['pic']}")
                        st.write(f"**Keterangan:** {reimb['remarks'] if reimb['remarks'] else '-'}")
                        
                        st.write("**Detail Rincian Item Keperluan:**")
                        for itm in reimb["items"]:
                            st.write(f"- {itm['description']} ({itm['qty']}x) : **Rp {itm['total']:,.0f}**")
                            
                    with c2:
                        st.write(f"**Status Approval COO:** `{reimb['status_coo']}`")
                        st.write(f"**Status Approval CFO:** `{reimb['status_cfo']}`")
                        
                        pdf_data = generate_a5_reimbursement_pdf(reimb)
                        st.download_button(
                            label="📥 Download PDF (1 Page + Struk)",
                            data=pdf_data,
                            file_name=f"Reimbursement_{reimb['form_no'].replace('/', '_')}.pdf",
                            mime="application/pdf",
                            key=f"dl_hist_{reimb['form_no']}"
                        )

                    st.markdown("---")
                    st.write("📸 **Bukti Nota / Struk Terlampir Per Baris Item:**")
                    
                    has_any_img = False
                    img_cols = st.columns(3)
                    col_idx = 0
                    
                    for itm in reimb.get("items", []):
                        img_link = str(itm.get("evident", "")).strip()
                        if img_link and img_link not in ["0", "None", ""]:
                            has_any_img = True
                            with img_cols[col_idx % 3]:
                                render_secure_image(img_link, caption=f"Item : {itm['description']}", use_container_width=True)
                            col_idx += 1

                    if not has_any_img and "image_links" in reimb and reimb["image_links"]:
                        for img_link in reimb["image_links"]:
                            img_link_str = str(img_link).strip()
                            if img_link_str and img_link_str not in ["0", "None", ""]:
                                has_any_img = True
                                with img_cols[col_idx % 3]:
                                    render_secure_image(img_link_str, caption="Bukti Nota Struk", use_container_width=True)
                                col_idx += 1
                        
                    if not has_any_img:
                        st.caption("*(Tidak ada lampiran foto nota/struk online)*")

                    st.markdown("---")
                    st.write("**Aksi Approval (Berdasarkan Hak Akses Akun Login):**")
                    
                    col_app1, col_app2, col_app3, col_app4 = st.columns(4)
                    
                    # Hak Akses Tombol COO
                    if can_coo:
                        with col_app1:
                            if st.button("✅ Approve (COO)", key=f"app_coo_{reimb['form_no']}"):
                                if update_reimbursement_status(reimb['form_no'], "Pending CFO", "coo"):
                                    st.success("Persetujuan COO Berhasil!")
                                    st.rerun()

                        with col_app2:
                            if st.button("❌ Reject (COO)", key=f"rej_coo_{reimb['form_no']}"):
                                if update_reimbursement_status(reimb['form_no'], "Rejected", "coo"):
                                    st.error("Pengajuan Ditolak COO!")
                                    st.rerun()
                    else:
                        with col_app1:
                            st.caption("🔒 *Terkunci (Bukan Akses COO)*")
                        with col_app2:
                            st.caption("🔒 *Terkunci (Bukan Akses COO)*")

                    # Hak Akses Tombol CFO
                    if can_cfo:
                        with col_app3:
                            if st.button("✅ Approve (CFO)", key=f"app_cfo_{reimb['form_no']}"):
                                if update_reimbursement_status(reimb['form_no'], "Approved", "cfo"):
                                    st.success("Persetujuan Final CFO Berhasil!")
                                    st.rerun()

                        with col_app4:
                            if st.button("❌ Reject (CFO)", key=f"rej_cfo_{reimb['form_no']}"):
                                if update_reimbursement_status(reimb['form_no'], "Rejected", "cfo"):
                                    st.error("Pengajuan Ditolak CFO!")
                                    st.rerun()
                    else:
                        with col_app3:
                            st.caption("🔒 *Terkunci (Bukan Akses CFO)*")
                        with col_app4:
                            st.caption("🔒 *Terkunci (Bukan Akses CFO)*")


# ==============================================================================
# ALIAS & ENTRY POINT
# ==============================================================================

def render():
    render_reimbursement_page()

if __name__ == "__main__":
    render_reimbursement_page()
