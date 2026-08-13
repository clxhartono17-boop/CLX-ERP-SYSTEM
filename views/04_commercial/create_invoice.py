import os
import io
import streamlit as st
import pandas as pd
from datetime import datetime
import gspread

# Import ReportLab untuk PDF
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PIL import Image as PILImage

# ==============================================================================
# 🎯 SPREADSHEET ID
# ==============================================================================
SPREADSHEET_ID = "1FU1lL3ls3jP_hAxBdx_Fu35Z9Ap4ICdHmOpMvCyA3gY"

MONTH_ROMAN = {
    1: "I", 2: "II", 3: "III", 4: "IV", 5: "V", 6: "VI",
    7: "VII", 8: "VIII", 9: "IX", 10: "X", 11: "XI", 12: "XII"
}

@st.cache_resource
def init_gspread():
    cred_path = "credentials.json"
    if not os.path.exists(cred_path):
        st.error(f"🚨 File '{cred_path}' tidak ditemukan di root folder project!")
        return None, None, None, None

    try:
        gc = gspread.service_account(filename=cred_path)
        sh = gc.open_by_key(SPREADSHEET_ID)

        sheet_query = sh.worksheet("Query")
        sheet_dropdown = sh.worksheet("Master Dropdown")
        sheet_erp = sh.worksheet("ERP Project")
        sheet_db = sh.worksheet("DB Invoice")
        
        return sheet_query, sheet_dropdown, sheet_erp, sheet_db
    except Exception as e:
        st.error(f"🚨 Gagal terhubung ke Google Sheets: {e}")
        return None, None, None, None


def get_raw_matrix(sheet):
    """Membaca seluruh isi sheet menjadi 2D List"""
    try:
        return sheet.get_all_values()
    except Exception:
        return []


def generate_auto_invoice_no(sheet_db):
    """Generate No. Invoice Otomatis -> Format: 0000/INV/CLX/VIII/2026"""
    now = datetime.now()
    roman_month = MONTH_ROMAN.get(now.month, "VIII")
    year = now.year

    if sheet_db:
        try:
            records = sheet_db.get_all_records()
            next_seq = len(records) + 1
        except Exception:
            next_seq = 1
    else:
        next_seq = 1

    formatted_seq = f"{next_seq:04d}"
    return f"{formatted_seq}/INV/CLX/{roman_month}/{year}"


def create_invoice_pdf(data):
    """Fungsi Pembuat File PDF Invoice dengan Header.png & Footer.png"""
    buffer = io.BytesIO()
    
    # Margin atas & bawah disesuaikan agar isi PDF tidak bertabrakan dengan gambar Header/Footer
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=letter, 
        rightMargin=36, 
        leftMargin=36, 
        topMargin=90,   # Ruang aman untuk Header
        bottomMargin=80 # Ruang aman untuk Footer
    )
    elements = []
    styles = getSampleStyleSheet()

    # Style Custom
    title_style = ParagraphStyle(
        name="TitleStyle", parent=styles['Heading1'], fontSize=22, leading=26, 
        textColor=colors.HexColor("#1A365D"), alignment=2
    )
    normal_bold = ParagraphStyle(name="NormalBold", parent=styles['Normal'], fontSize=8.5, leading=11, fontName="Helvetica-Bold")
    normal_text = ParagraphStyle(name="NormalText", parent=styles['Normal'], fontSize=8.5, leading=11)
    pay_title_style = ParagraphStyle(name="PayTitle", parent=styles['Normal'], fontSize=9, leading=11, fontName="Helvetica-Bold", textColor=colors.HexColor("#1A365D"))

    # --- PATH GAMBAR HEADER & FOOTER ---
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    header_path = os.path.join(base_dir, "assets", "header.png")
    footer_path = os.path.join(base_dir, "assets", "Footer.png")

    if not os.path.exists(header_path):
        header_path = os.path.abspath("assets/header.png")
    if not os.path.exists(footer_path):
        footer_path = os.path.abspath("assets/Footer.png")

    # Fungsi Callback untuk Menggambar Header & Footer Tepat di Kertas (Canvas)
    def draw_header_footer(canvas, doc):
        canvas.saveState()
        page_w, page_h = letter # Lebar & Tinggi Halaman Letter
        
        # Draw Header Image (Bagian Atas)
        if os.path.exists(header_path):
            try:
                # Menempelkan header persis di paling atas halaman
                canvas.drawImage(header_path, 0, page_h - 75, width=page_w, height=75, mask='auto')
            except Exception:
                pass

        # Draw Footer Image (Bagian Bawah)
        if os.path.exists(footer_path):
            try:
                # Menempelkan footer persis di paling bawah halaman
                canvas.drawImage(footer_path, 0, 0, width=page_w, height=65, mask='auto')
            except Exception:
                pass

        canvas.restoreState()

    # --- 1. INVOICE TITLE ---
    inv_title = Paragraph("<b>INVOICE</b>", title_style)
    elements.append(inv_title)
    elements.append(Spacer(1, 10))

    # --- 2. BILL TO & INVOICE META DATA ---
    bill_to_text = Paragraph("""
    <b>Bill To :</b><br/>
    <b>PT Vgreen Global Charging Station Investment Indonesia</b><br/>
    Graha Binakarsa Lt.7, Jl. H.R Rasuna Said Kav C-18 RT 02 RW 005,<br/>
    Karet Kuningan, Kec. Setiabudi Jakarta Selatan
    """, normal_text)

    issuer_and_meta = [
        [Paragraph("<b>Invoice No.</b>", normal_bold), Paragraph(f": {data['inv_no']}", normal_text)],
        [Paragraph("<b>Invoice Date</b>", normal_bold), Paragraph(f": {data['inv_date']}", normal_text)],
        [Paragraph("<b>Project Name</b>", normal_bold), Paragraph(f": {data['project_name']}", normal_text)],
        [Paragraph("<b>WO No.</b>", normal_bold), Paragraph(f": {data['wo_no']}", normal_text)],
        [Paragraph("<b>No. Efaktur</b>", normal_bold), Paragraph(f": {data['efaktur']}", normal_text)],
    ]
    meta_table = Table(issuer_and_meta, colWidths=[80, 180])
    meta_table.setStyle(TableStyle([
        ('PADDING', (0,0), (-1,-1), 2),
        ('VALIGN', (0,0), (-1,-1), 'TOP')
    ]))

    info_grid = Table([[bill_to_text, meta_table]], colWidths=[280, 260])
    info_grid.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 0)
    ]))
    elements.append(info_grid)
    elements.append(Spacer(1, 15))

    # --- 3. ITEM TABLE ---
    table_data = [
        [
            Paragraph("<b>NO.</b>", normal_bold), 
            Paragraph("<b>ITEM DESCRIPTION</b>", normal_bold), 
            Paragraph("<b>CHARGING TYPE</b>", normal_bold), 
            Paragraph("<b>TERMIN</b>", normal_bold), 
            Paragraph("<b>AMOUNT (IDR)</b>", normal_bold)
        ],
        [
            "1", 
            f"{data['site_name']}", 
            f"{data['charging_type']}", 
            f"{data['termin']}", 
            f"{data['subtotal']:,.2f}"
        ]
    ]
    item_table = Table(table_data, colWidths=[30, 230, 90, 60, 130])
    item_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#9FA5AD")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (0,-1), 'CENTER'),
        ('ALIGN', (3,0), (3,-1), 'CENTER'),
        ('ALIGN', (4,0), (4,-1), 'RIGHT'),
        ('PADDING', (0,0), (-1,-1), 6),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    elements.append(item_table)
    elements.append(Spacer(1, 10))

    # --- 4. SUMMARY TOTALS ---
    summary_data = [
        ["Subtotal (Excl. Tax) :", f"Rp {data['subtotal']:,.2f}"],
        [f"PPN ({data['tax_rate']}%) :", f"Rp {data['tax_amount']:,.2f}"],
        ["TOTAL AMOUNT :", f"Rp {data['grand_total']:,.2f}"]
    ]
    summary_table = Table(summary_data, colWidths=[380, 160])
    summary_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('LINEABOVE', (0,-1), (-1,-1), 1, colors.HexColor("#2B6CB0")),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 15))

    # --- 5. PAYMENT INFO TABLE (KOTAK TABEL LENGKAP & RAPI) ---
    pay_table_data = [
        [Paragraph("<b>PAYMENT INFO</b>", pay_title_style), "", ""],
        [Paragraph("Bank Name", normal_bold), ":", Paragraph("BANK CENTRAL ASIA (BCA)", normal_text)],
        [Paragraph("Bank Account", normal_bold), ":", Paragraph("540-5282841", normal_text)],
        [Paragraph("Account Name", normal_bold), ":", Paragraph("PT. CONNECTIVITY LEADS EXCELLENCE", normal_text)],
    ]
    
    payment_info_table = Table(pay_table_data, colWidths=[85, 10, 205])
    payment_info_table.setStyle(TableStyle([
        ('SPAN', (0,0), (2,0)),                                     # Merge Header
        ('BACKGROUND', (0,0), (2,0), colors.HexColor("#E2E8F0")),   # Background Header Abu-abu
        ('PADDING', (0,0), (-1,-1), 4),                             # Cell Padding
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#A0AEC0")),  # Garis Sel Tabel
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#2B6CB0")),     # Border Luar Utama
    ]))

    # Signature Block
    now_str = datetime.now().strftime("%d %B %Y")
    sig_block = Paragraph(f"""
    Jakarta, {now_str}<br/><br/><br/><br/>
    <b><u>Christian</u></b><br/>
    President Director
    """, ParagraphStyle(name="SigStyle", parent=styles['Normal'], fontSize=8.5, leading=12, alignment=1))

    # Penempatan Sejajar Bagian Bawah
    bottom_grid = Table([[payment_info_table, sig_block]], colWidths=[300, 240])
    bottom_grid.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
        ('ALIGN', (1,0), (1,0), 'CENTER'),
        ('PADDING', (0,0), (-1,-1), 0)
    ]))
    elements.append(bottom_grid)

    # Build PDF dengan memanggil callback draw_header_footer
    doc.build(elements, onFirstPage=draw_header_footer, onLaterPages=draw_header_footer)
    buffer.seek(0)
    return buffer.getvalue()


def render():
    st.title("📄 Create Invoice")
    st.caption("Modul Pembuatan Invoice - CLX ERP System")
    st.markdown("---")

    sheet_query, sheet_dropdown, sheet_erp, sheet_db = init_gspread()

    if not sheet_query:
        return

    # Load All Sheets
    raw_query = get_raw_matrix(sheet_query)
    raw_dropdown = get_raw_matrix(sheet_dropdown)
    raw_erp = get_raw_matrix(sheet_erp)
    raw_db = get_raw_matrix(sheet_db)

    auto_inv_no = generate_auto_invoice_no(sheet_db)

    # --- PROJECT OPTIONS ---
    project_options = []
    if len(raw_dropdown) > 1:
        for row in raw_dropdown[1:]:
            if len(row) > 6 and row[6].strip():
                val = row[6].strip()
                if val not in project_options:
                    project_options.append(val)
    
    if not project_options:
        project_options = ["VGreen - Project", "VGreen - Operation", "SIP", "Charge Core"]

    # --- FORM UTAMA ---
    with st.form("form_create_invoice", clear_on_submit=False):
        col1, col2 = st.columns(2)

        with col1:
            selected_project = st.selectbox("Project Name", options=project_options)
            st.text_input("No. Invoice (Auto Generated)", value=auto_inv_no, disabled=True)
            inv_date = st.date_input("Invoice Date", datetime.now())

            # --- SITE NAME OPTIONS ---
            site_name = ""
            if selected_project == "VGreen - Project":
                site_options = []
                if len(raw_query) > 1:
                    for row in raw_query[1:]:
                        if len(row) > 5 and row[5].strip():
                            s_name = row[5].strip()
                            if s_name not in site_options:
                                site_options.append(s_name)
                                
                site_name = st.selectbox("Item Description / Site Name", options=site_options)
            else:
                site_name = st.text_input("Item Description / Site Name", placeholder="Ketik Site Name secara manual...")

        # --- AUTO FETCH CHARGING TYPE & WO ---
        charging_type = ""
        wo_number = ""

        if selected_project == "VGreen - Project" and site_name and len(raw_query) > 1:
            for row in raw_query[1:]:
                if len(row) > 5 and row[5].strip().lower() == site_name.strip().lower():
                    if len(row) > 2:
                        charging_type = row[2].strip()
                    if len(row) > 22:
                        wo_number = row[22].strip()
                    break

        with col2:
            st.text_input("Charging Type (Auto)", value=charging_type, disabled=True)
            st.text_input("WO No. (Auto)", value=wo_number, disabled=True)
            efaktur_no = st.text_input("No. Efaktur", placeholder="Contoh: 04002600290355647")

        st.markdown("### 💳 Detail Pembayaran & Termin")
        
        # --- 🔒 DETEKSI & ELIMINASI TERMIN YANG SUDAH DIPAKAI DARI DB INVOICE ---
        all_possible_termins = ["35%", "60%", "5%"]
        used_termins = []

        if selected_project == "VGreen - Project" and len(raw_db) > 1:
            for r in raw_db[1:]:
                if len(r) > 6:
                    db_site = r[4].strip().lower()
                    db_term = r[6].strip()
                    if db_site == site_name.strip().lower():
                        used_termins.append(db_term)

        # Filter opsi termin yang BELUM pernah dipakai
        available_termins = [t for t in all_possible_termins if t not in used_termins]

        if selected_project == "VGreen - Project":
            if used_termins:
                st.warning(f"ℹ️ Termin yang sudah pernah dibuat untuk site **{site_name}**: {', '.join(set(used_termins))}")
            
            if not available_termins:
                st.error(f"⛔ Semua termin (35%, 60%, 5%) untuk site **{site_name}** sudah terpakai di DB Invoice! Tidak bisa membuat invoice lagi.")

        c1, c2, c3 = st.columns(3)

        with c1:
            if available_termins:
                selected_termin = st.selectbox("Termin Pembayaran", options=available_termins)
                try:
                    pct_val = float(selected_termin.replace("%", "").strip()) / 100.0
                except ValueError:
                    pct_val = 1.0
            else:
                selected_termin = "N/A"
                st.selectbox("Termin Pembayaran", options=["Penuh / Lunas"], disabled=True)
                pct_val = 0.0

        # --- FETCH BOQ AMOUNT (ERP Project Kolom E = Indeks 4, Kolom M = Indeks 12) ---
        boq_amount = 0.0
        if selected_project == "VGreen - Project" and site_name and len(raw_erp) > 1:
            for row in raw_erp[1:]:
                if len(row) > 4:
                    erp_site = row[4].strip().lower()
                    erp_charge = row[2].strip().lower() if len(row) > 2 else ""

                    if erp_site == site_name.strip().lower() and (not charging_type or erp_charge == charging_type.strip().lower()):
                        if len(row) > 12:
                            raw_boq = row[12].replace("Rp", "").replace(".", "").replace(",", ".").replace(" ", "").strip()
                            try:
                                boq_amount = float(raw_boq)
                            except ValueError:
                                boq_amount = 0.0
                        break

        calculated_unit_price = boq_amount * pct_val

        with c2:
            if selected_project == "VGreen - Project":
                unit_price = st.number_input("Unit Price (IDR - Auto Calculated)", value=calculated_unit_price, disabled=True, format="%.2f")
            else:
                unit_price = st.number_input("Unit Price (IDR)", min_value=0.0, value=0.0, step=1000.0)

        with c3:
            tax_rate = st.number_input("PPN / Tax (%)", min_value=0.0, value=11.0, step=0.5)

        # Totals
        termin_amount = unit_price
        tax_amount = termin_amount * (tax_rate / 100)
        grand_total = termin_amount + tax_amount

        st.info(f"""
        **Ringkasan Perhitungan:**
        * **BOQ Amount Base (Sheet ERP):** Rp {boq_amount:,.2f}
        * **Nilai Invoice ({selected_termin}):** Rp {termin_amount:,.2f}
        * **PPN ({tax_rate}%):** Rp {tax_amount:,.2f}
        * **GRAND TOTAL:** Rp {grand_total:,.2f}
        """)

        is_disabled = True if (selected_project == "VGreen - Project" and not available_termins) else False
        submit_btn = st.form_submit_button("💾 Save to DB Invoice", type="primary", disabled=is_disabled)

    # --- SAVE TO DB & GENERATE PDF ---
    if submit_btn:
        if not efaktur_no:
            st.error("⚠️ No. Efaktur wajib diisi!")
        elif not site_name:
            st.error("⚠️ Site Name / Item Description wajib diisi!")
        else:
            new_row = [
                selected_project,
                auto_inv_no,
                inv_date.strftime("%d %B %Y"),
                charging_type,
                site_name,
                wo_number,
                selected_termin,
                efaktur_no,
                f"{termin_amount:.2f}",
                f"{grand_total:.2f}"
            ]

            if sheet_db:
                try:
                    sheet_db.append_row(new_row)
                    st.success(f"✅ Invoice **{auto_inv_no}** berhasil disimpan ke Sheet 'DB Invoice'!")
                    st.balloons()

                    # Render PDF Bytes
                    pdf_payload = {
                        "project_name": selected_project,
                        "inv_no": auto_inv_no,
                        "inv_date": inv_date.strftime("%d %B %Y"),
                        "charging_type": charging_type,
                        "site_name": site_name,
                        "wo_no": wo_number,
                        "termin": selected_termin,
                        "efaktur": efaktur_no,
                        "subtotal": termin_amount,
                        "tax_rate": tax_rate,
                        "tax_amount": tax_amount,
                        "grand_total": grand_total
                    }

                    pdf_file_bytes = create_invoice_pdf(pdf_payload)
                    
                    st.session_state["pdf_ready"] = pdf_file_bytes
                    st.session_state["pdf_name"] = f"Invoice_{auto_inv_no.replace('/', '_')}.pdf"

                except Exception as e:
                    st.error(f"🚨 Gagal menyimpan ke Google Sheets: {e}")
            else:
                st.error("🚨 Tidak terhubung ke Google Sheets.")

    # --- Area Download PDF ---
    if "pdf_ready" in st.session_state and st.session_state["pdf_ready"]:
        st.markdown("---")
        st.subheader("📥 Download Berkas Invoice")
        st.download_button(
            label="📄 Download Invoice PDF",
            data=st.session_state["pdf_ready"],
            file_name=st.session_state["pdf_name"],
            mime="application/pdf",
            type="primary",
            key="dl_btn_invoice"
        )