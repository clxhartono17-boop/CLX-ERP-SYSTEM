import datetime
import io
import os
import re
import sys
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
import streamlit as st

# ==============================================================================
# SAFE IMPORT FOR SERVICES & ROOT DIR CONFIGURATION
# ==============================================================================
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, "../../"))

if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

CWD = os.getcwd()
if CWD not in sys.path:
    sys.path.insert(0, CWD)

try:
    from services.gsheet import get_google_sheet_connection
except ModuleNotFoundError:

    def get_google_sheet_connection():
        return None


# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================
def parse_price(val):
    if (
        pd.isna(val)
        or val == "-"
        or str(val).strip() == "-"
        or str(val).strip().lower() == "nan"
    ):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = (
        str(val)
        .replace("Rp.", "")
        .replace("Rp", "")
        .replace(",", "")
        .replace(".", "")
        .strip()
    )
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_qty_num(val):
    if pd.isna(val) or str(val).strip() in ["-", "", "nan", "None"]:
        return 0.0
    s = str(val).strip()
    try:
        return float(s.replace(".", "").replace(",", "."))
    except Exception:
        match = re.search(r"^([\d\.]+)", s)
        if match:
            try:
                return float(match.group(1).replace(".", ""))
            except Exception:
                return 1.0
    return 1.0


def map_to_standard_province(raw_province):
    if not raw_province or pd.isna(raw_province):
        return "JAVA"

    p = str(raw_province).strip().upper()

    java_keywords = [
        "JAVA", "JAWA", "BANTEN", "DKI", "JAKARTA", "JABODETABEK",
        "YOGYAKARTA", "DIY", "WEST JAVA", "EAST JAVA", "CENTRAL JAVA"
    ]
    sumatera_keywords = [
        "SUMATERA", "SUMATRA", "ACEH", "MEDAN", "RIU", "RIAU", "JAMBI",
        "LAMPUNG", "BENGKULU", "PALEMBANG", "PADANG", "BABEL", "BANGKA"
    ]
    kalimantan_keywords = [
        "KALIMANTAN", "BORNEO", "PONTIANAK", "BANJARMASIN", "BALIKPAPAN", "SAMARINDA"
    ]
    sulawesi_keywords = [
        "SULAWESI", "CELEBES", "MAKASSAR", "MANADO", "PALU", "GORONTALO", "MAMUJU", "POLEWALI"
    ]
    bali_nusa_keywords = [
        "BALI", "NUSA", "NTB", "NTT", "LOMBOK", "DENPASAR", "SUMBAWA", "FLORES"
    ]

    if any(k in p for k in java_keywords):
        return "JAVA"
    elif any(k in p for k in sumatera_keywords):
        return "SUMATERA"
    elif any(k in p for k in kalimantan_keywords):
        return "KALIMANTAN"
    elif any(k in p for k in sulawesi_keywords):
        return "SULAWESI"
    elif any(k in p for k in bali_nusa_keywords):
        return "BALI NUSATENGGARA"

    return "JAVA"


def generate_boq_number(sequence_num=1):
    now = datetime.datetime.now()
    roman_months = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
    month_roman = roman_months[now.month - 1]
    return f"{sequence_num:04d}/CLX/BOQ/{month_roman}/{now.year}"


def get_all_saved_boq():
    """Mengambil seluruh data BOQ tersimpan dari sheet DB BOQ"""
    try:
        sh = get_google_sheet_connection()
        if not sh:
            return []
        worksheet = sh.worksheet("DB BOQ")
        rows = worksheet.get_all_values()
        if len(rows) <= 1:
            return []
        
        headers = [h.strip() for h in rows[0]]
        data = []
        for idx, r in enumerate(rows[1:], start=2):
            if any(r):
                item = {"row_idx": idx}
                for h_idx, h in enumerate(headers):
                    item[h] = r[h_idx] if h_idx < len(r) else ""
                data.append(item)
        return data
    except Exception as e:
        st.error(f"Gagal mengambil data DB BOQ: {e}")
        return []


def get_existing_saved_site_charger_pairs():
    """
    Mengambil set tuple (site_name_lowercase, charger_type_lowercase) 
    yang sudah tersimpan di DB BOQ agar pengecekan duplikasi presisi.
    """
    saved_boqs = get_all_saved_boq()
    saved_pairs = set()
    for item in saved_boqs:
        site_name = item.get("Site Name", "").strip().lower()
        charger_type = item.get("Charger Type", "").strip().lower()
        if site_name and charger_type:
            saved_pairs.add((site_name, charger_type))
    return saved_pairs


def fetch_query_site_options(exclude_saved=True):
    """
    Mengambil daftar site aktif dari sheet Query.
    Memfilter status DROP & CANCEL serta (Site Name + Charger Type) yang sudah tersimpan di DB BOQ.
    """
    site_options = []
    site_data_map = {}

    existing_saved_pairs = get_existing_saved_site_charger_pairs() if exclude_saved else set()

    try:
        sh = get_google_sheet_connection()
        if sh:
            worksheet = sh.worksheet("Query")
            data_query = worksheet.get_all_values()

            if len(data_query) > 1:
                for row in data_query[1:]:
                    if len(row) > 5:
                        epc_val = str(row[1]).strip() if len(row) > 1 else "-"
                        charger_val = str(row[2]).strip() if len(row) > 2 else "DC20"
                        status_val = str(row[3]).strip() if len(row) > 3 else ""
                        site_val = str(row[5]).strip() if len(row) > 5 else ""
                        address_val = str(row[6]).strip() if len(row) > 6 else "-"
                        raw_prov_val = str(row[8]).strip() if len(row) > 8 else ""

                        # 🔍 FILTER 1: Eliminasi status DROP atau CANCEL
                        status_upper = status_val.upper()
                        if "DROP" in status_upper or "CANCEL" in status_upper:
                            continue

                        # 🔍 FILTER 2: Presisi (Site Name + Charger Type) yang sudah tersimpan
                        pair_key = (site_val.strip().lower(), charger_val.strip().lower())
                        if exclude_saved and pair_key in existing_saved_pairs:
                            continue

                        if site_val and site_val.lower() not in ["nan", "", "none", "project / location name"]:
                            std_province = map_to_standard_province(raw_prov_val)
                            display_label = f"{site_val} ({charger_val})"

                            if display_label not in site_data_map:
                                site_options.append(display_label)

                            site_data_map[display_label] = {
                                "site_name": site_val,
                                "epc": epc_val if epc_val else "-",
                                "address": address_val if address_val else "-",
                                "charger": charger_val if charger_val else "DC20",
                                "province": std_province,
                                "raw_province": raw_prov_val if raw_prov_val else "-",
                                "status": status_val
                            }
    except Exception as e:
        st.sidebar.warning(f"⚠️ Gagal membaca sheet 'Query': {e}")

    if not site_options:
        site_options = ["(Semua kombinasi Site & Charger telah dibuatkan BOQ)"]
        site_data_map = {
            "(Semua kombinasi Site & Charger telah dibuatkan BOQ)": {
                "site_name": "-",
                "epc": "-",
                "address": "-",
                "charger": "-",
                "province": "JAVA",
                "raw_province": "-",
                "status": "-"
            }
        }

    return site_options, site_data_map


def load_boq_dataframe(charging_type, province_str):
    """Membaca template Excel berdasarkan jenis charger dan wilayah/provinsi"""
    region_normalized = map_to_standard_province(province_str)
    raw_charging_key = charging_type.upper().strip()

    sheet_map = {
        "20 KW": "DC20", "20KW": "DC20", "DC20": "DC20",
        "30 KW": "DC30", "30KW": "DC30", "DC30": "DC30",
        "60 KW": "DC60", "60KW": "DC60", "DC60": "DC60",
        "120 KW": "DC120", "120KW": "DC120", "DC120": "DC120",
        "6S1P": "6S1P", "12S1P": "12S1P", "7KW": "7KW", "7 KW": "7KW", "22KW": "22KW", "22 KW": "22KW",
    }
    target_sheet = sheet_map.get(raw_charging_key, raw_charging_key)
    
    # 🔍 INDEKS KOLOM UNTUK TIAP WILAYAH
    region_col_indices = {
        "JAVA": 0, 
        "SUMATERA": 8, 
        "BALI NUSATENGGARA": 16, 
        "BALI NUSA TENGGARA": 16,
        "KALIMANTAN": 25, 
        "SULAWESI": 33
    }
    start_idx = region_col_indices.get(region_normalized, 0)

    possible_paths = [
        os.path.join(ROOT_DIR, "assets", "templates", "Template BOQ Vgreen.xlsx"),
        os.path.join(CWD, "assets", "templates", "Template BOQ Vgreen.xlsx"),
        os.path.join(os.path.dirname(CURRENT_DIR), "assets", "templates", "Template BOQ Vgreen.xlsx"),
        os.path.join(CURRENT_DIR, "assets", "templates", "Template BOQ Vgreen.xlsx"),
    ]

    excel_path = next((p for p in possible_paths if os.path.exists(p)), None)
    if not excel_path:
        st.error("⚠️ File template Excel tidak ditemukan!")
        return None, target_sheet, region_normalized

    try:
        xls = pd.ExcelFile(excel_path)
        sheet_found = next((s for s in xls.sheet_names if s.strip().lower() == target_sheet.strip().lower()), None)
        
        if sheet_found:
            df_raw = pd.read_excel(xls, sheet_name=sheet_found, header=None)
            
            if start_idx + 6 >= df_raw.shape[1]:
                start_idx = 0

            col_indices = [start_idx + i for i in range(7)]

            df_boq = df_raw.iloc[6:55, col_indices].copy()
            df_boq.columns = ["NO", "Item", "Unit/Volume", "Satuan/Uom", "MERK", "UNIT PRICE", "TOTAL PRICE"]
            df_boq = df_boq.dropna(subset=["Item"])
            df_boq["Item"] = df_boq["Item"].astype(str).str.strip()
            df_boq = df_boq[(df_boq["Item"] != "") & (df_boq["Item"].str.lower() != "nan")].reset_index(drop=True)

            for c in ["UNIT PRICE", "TOTAL PRICE"]:
                df_boq[c] = df_boq[c].apply(parse_price)
            for c in ["NO", "Unit/Volume", "Satuan/Uom", "MERK"]:
                df_boq[c] = df_boq[c].fillna("-").astype(str).str.replace("nan", "-", case=False)
                
            return df_boq, target_sheet, region_normalized
        else:
            st.error(f"⚠️ Sheet '{target_sheet}' tidak ditemukan di file Excel Template.")
    except Exception as e:
        st.error(f"Gagal membaca Excel Template: {e}")
        
    return None, target_sheet, region_normalized


def save_to_db_boq(site_name, charger_capacity, sub_total, grand_total, epc_name="-"):
    try:
        sh = get_google_sheet_connection()
        if not sh:
            return None
        try:
            worksheet = sh.worksheet("DB BOQ")
        except Exception:
            worksheet = sh.add_worksheet(title="DB BOQ", rows=1000, cols=10)
            worksheet.append_row([
                "No", "BOQ No.", "Site Name", "Charger Type", 
                "BOQ Amount Exc. PPN", "BOQ Amount inc. PPN", "EPC Name"
            ])

        existing_rows = worksheet.get_all_values()
        no_urut = (
            len([r for r in existing_rows[1:] if any(r)]) + 1
            if len(existing_rows) > 1
            else 1
        )
        boq_no = generate_boq_number(sequence_num=no_urut)

        new_row = [
            no_urut, boq_no, str(site_name), str(charger_capacity),
            sub_total, grand_total, str(epc_name)
        ]
        worksheet.append_row(new_row)
        return boq_no
    except Exception as e:
        st.error(f"❌ Gagal menyimpan ke DB BOQ: {e}")
        return None


def update_db_boq_row(row_idx, old_site_name, new_site_name, charger_capacity, sub_total, grand_total, epc_name):
    """Mengubah data baris tertentu di sheet DB BOQ dan menyinkronkan ke Sum Project"""
    try:
        sh = get_google_sheet_connection()
        if not sh:
            return False
        worksheet = sh.worksheet("DB BOQ")
        
        worksheet.update_cell(row_idx, 3, new_site_name)
        worksheet.update_cell(row_idx, 4, charger_capacity)
        worksheet.update_cell(row_idx, 5, sub_total)
        worksheet.update_cell(row_idx, 6, grand_total)
        worksheet.update_cell(row_idx, 7, epc_name)

        update_google_sheet_summary(old_site_name, new_site_name, sub_total, grand_total)
        return True
    except Exception as e:
        st.error(f"❌ Gagal memperbarui DB BOQ: {e}")
        return False


def update_google_sheet_summary(old_site_name, new_site_name, sub_total, grand_total):
    """Update otomatis ke sheet Summary (Sum Project)"""
    try:
        sh = get_google_sheet_connection()
        if not sh:
            return False

        sheet_names = [ws.title for ws in sh.worksheets()]
        if "Sum Project" not in sheet_names:
            return False

        worksheet = sh.worksheet("Sum Project")
        data_sum = worksheet.get_all_values()

        for idx, row in enumerate(data_sum[1:], start=2):
            if len(row) > 2 and str(row[2]).strip().lower() == str(old_site_name).strip().lower():
                worksheet.update_cell(idx, 3, new_site_name)
                if len(row) >= 19:
                    worksheet.update_cell(idx, 18, sub_total)
                    worksheet.update_cell(idx, 19, grand_total)
                return True
    except Exception as e:
        st.caption(f"ℹ️ Info: Sheet 'Sum Project' belum ter-update ({e})")
    return False


# ==============================================================================
# PDF GENERATOR
# ==============================================================================
def generate_boq_pdf(site_name, site_location, charger_capacity, region, df_boq, sub_total, vat, grand_total):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=10, leftMargin=10, topMargin=10, bottomMargin=10)
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle("DocTitle", parent=styles["Heading1"], fontSize=9, leading=10, fontName="Helvetica-Bold", textColor=colors.HexColor("#111111"), spaceAfter=1)
    sub_style = ParagraphStyle("DocSub", parent=styles["Normal"], fontSize=7, leading=8, fontName="Helvetica", textColor=colors.HexColor("#333333"), spaceAfter=0)
    
    table_text = ParagraphStyle("TableText", parent=styles["Normal"], fontSize=5.5, leading=6.5, fontName="Helvetica")
    table_text_bold = ParagraphStyle("TableTextBold", parent=table_text, fontName="Helvetica-Bold")
    table_header = ParagraphStyle("TableHeader", parent=styles["Normal"], fontSize=6, leading=7.5, fontName="Helvetica-Bold", textColor=colors.white)

    elements.append(Paragraph(f"CHARGING WORK {charger_capacity}", title_style))
    elements.append(Paragraph(f"<b>Site Name:</b> {site_name}", sub_style))
    elements.append(Paragraph(f"<b>Site Location:</b> {site_location}", sub_style))
    elements.append(Paragraph(f"<b>NEW PLAN BOQ VGREEN - {region} ISLAND</b>", ParagraphStyle("SubHeader", parent=title_style, fontSize=7.5, leading=8.5, spaceAfter=2, spaceBefore=1)))

    table_data = [[
        Paragraph("NO", table_header), Paragraph("Item", table_header), Paragraph("Unit/Vol", table_header),
        Paragraph("Satuan", table_header), Paragraph("MERK", table_header), Paragraph("UNIT PRICE", table_header), Paragraph("TOTAL PRICE", table_header),
    ]]

    parent_headers = ["A", "B", "C", "D", "E", "F"]

    for _, row in df_boq.iterrows():
        no_str = str(row.get("NO", "")).strip().upper()
        is_parent = no_str in parent_headers

        up_val = parse_price(row.get("UNIT PRICE", 0))
        tp_val = parse_price(row.get("TOTAL PRICE", 0))

        up_str = f"Rp. {up_val:,.0f}".replace(",", ".") if up_val > 0 else ("-" if not is_parent else "")
        tp_str = f"Rp. {tp_val:,.0f}".replace(",", ".") if tp_val > 0 else "-"

        style_to_use = table_text_bold if is_parent else table_text

        table_data.append([
            Paragraph(no_str, style_to_use), Paragraph(str(row.get("Item", "")), style_to_use),
            Paragraph(str(row.get("Unit/Volume", "")), style_to_use), Paragraph(str(row.get("Satuan/Uom", "")), style_to_use),
            Paragraph(str(row.get("MERK", "")), style_to_use), Paragraph(up_str, style_to_use), Paragraph(tp_str, style_to_use),
        ])

    table_data.append(["", "", Paragraph("<b>Sub Total:</b>", table_text_bold), "", "", "", Paragraph(f"<b>Rp. {sub_total:,.0f}</b>".replace(",", "."), table_text_bold)])
    table_data.append(["", "", Paragraph("<b>VAT 11%</b>", table_text_bold), "", "", "", Paragraph(f"<b>Rp. {vat:,.0f}</b>".replace(",", "."), table_text_bold)])
    table_data.append(["", "", Paragraph("<b>Total Contractor Price</b>", table_text_bold), "", "", "", Paragraph(f"<b>Rp. {grand_total:,.0f}</b>".replace(",", "."), table_text_bold)])

    col_widths = [18, 260, 45, 45, 67, 70, 70]
    t = Table(table_data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -4), 0.3, colors.HexColor("#CCCCCC")),
        ("BACKGROUND", (0, -3), (-1, -1), colors.HexColor("#F8F9F9")),
        ("LINEABOVE", (0, -3), (-1, -3), 0.8, colors.HexColor("#2C3E50")),
        ("TOPPADDING", (0, 0), (-1, -1), 0.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
    ]))

    elements.append(t)
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


# ==============================================================================
# MAIN RENDER FUNCTION
# ==============================================================================
def render():
    st.title("📝 Quotation & BOQ Manager")

    tab_create, tab_edit = st.tabs(["➕ Buat BOQ Baru", "✏️ Edit / Reuse / Re-Download BOQ"])

    # Fetch site options (Excluding Site + Charger Type combinations that are already saved)
    site_options, site_data_map = fetch_query_site_options(exclude_saved=True)

    # ==========================================================================
    # TAB 1: BUAT BOQ BARU
    # ==========================================================================
    with tab_create:
        st.subheader("1. Informasi Site & Spesifikasi")
        col_s1, col_s2 = st.columns(2)

        with col_s1:
            selected_label = st.selectbox(
                "Pilih Site Name & Charger (Filtered: Drop/Cancel & Saved BOQ Pair Excluded)", 
                sorted(list(site_options))
            )

        current_meta = site_data_map.get(selected_label, {
            "site_name": selected_label, "epc": "-", "address": "-", "charger": "-", "province": "JAVA", "raw_province": "-"
        })

        selected_site = current_meta.get("site_name", selected_label)

        with col_s2:
            site_address = st.text_input("Address (Kolom G)", value=current_meta["address"])

        col_s3, col_s4, col_s5 = st.columns([1.5, 1.5, 1])
        with col_s3:
            charging_type = st.text_input("Charging Type (Kolom C)", value=current_meta["charger"])
        with col_s4:
            province = st.text_input("Province Standar (Kolom I Mapped)", value=current_meta["province"])
            st.caption(f"📍 Raw Province Sheet: `{current_meta.get('raw_province', '-')}`")
        with col_s5:
            epc_name = st.text_input("EPC Name (Kolom B)", value=current_meta.get("epc", "-"))

        # 🔍 CEK INTEGRITAS PRESISI BERSAMAAN: (SITE NAME + CHARGER TYPE)
        saved_pairs = get_existing_saved_site_charger_pairs()
        current_pair_key = (selected_site.strip().lower(), charging_type.strip().lower())
        is_already_saved = current_pair_key in saved_pairs or selected_site == "-"

        if is_already_saved and selected_site != "-":
            st.warning(f"⚠️ **Peringatan:** Kombinasi Site `{selected_site}` dengan Charger `{charging_type}` sudah pernah dibuatkan BOQ! Silakan gunakan Tab **'Edit / Reuse / Re-Download BOQ'** jika ingin memperbarui data.")
        elif selected_site == "-":
            st.info("ℹ️ Semua kombinasi Site & Charger aktif di sheet Query telah dibuatkan BOQ.")

        df_boq, target_sheet, region_normalized = load_boq_dataframe(charging_type, province)

        if df_boq is not None and not df_boq.empty:
            st.subheader(f"2. Table BOQ ({target_sheet} - {region_normalized}) Auto Load")
            st.info("💡 Khusus **CABLING AND ACCESSORIES INSTALLATION** (Poin 1-3), **Qty/Volume** dapat diubah secara manual:")

            e_start = False
            editable_indices = []
            for idx, row in df_boq.iterrows():
                no_val = str(row.get("NO", "")).strip().upper()
                if "CABLING AND ACCESSORIES" in str(row.get("Item", "")).upper() or no_val == "E":
                    e_start = True
                    continue
                elif no_val in ["A", "B", "C", "D", "F"]:
                    e_start = False

                if e_start and no_val in ["1", "2", "3"]:
                    editable_indices.append(idx)

            col_e1, col_e2, col_e3 = st.columns(3)
            for idx in editable_indices:
                no_val = str(df_boq.loc[idx, "NO"])
                item_name = str(df_boq.loc[idx, "Item"])
                current_qty = parse_qty_num(df_boq.loc[idx, "Unit/Volume"])

                col_target = col_e1 if no_val == "1" else (col_e2 if no_val == "2" else col_e3)
                with col_target:
                    new_qty = st.number_input(
                        f"Qty Poin {no_val}: {item_name[:25]}...",
                        min_value=0.0,
                        value=float(current_qty),
                        step=1.0,
                        key=f"qty_e_{no_val}_{selected_label}",
                    )
                    df_boq.loc[idx, "Unit/Volume"] = str(int(new_qty) if new_qty.is_integer() else new_qty)

            parent_headers = ["A", "B", "C", "D", "E", "F"]
            current_parent_idx = None
            current_parent_sum = 0.0

            for idx, row in df_boq.iterrows():
                no_val = str(row.get("NO", "")).strip().upper()
                if no_val in parent_headers:
                    if current_parent_idx is not None and current_parent_sum > 0:
                        df_boq.loc[current_parent_idx, "TOTAL PRICE"] = current_parent_sum
                    current_parent_idx = idx
                    current_parent_sum = 0.0
                    vol_num = parse_qty_num(row.get("Unit/Volume", 0))
                    up_num = parse_price(row.get("UNIT PRICE", 0))
                    if vol_num > 0 and up_num > 0:
                        df_boq.loc[idx, "TOTAL PRICE"] = vol_num * up_num
                else:
                    vol_num = parse_qty_num(row.get("Unit/Volume", 0))
                    up_num = parse_price(row.get("UNIT PRICE", 0))
                    tp = (vol_num * up_num) if (vol_num > 0 and up_num > 0) else parse_price(row.get("TOTAL PRICE", 0))
                    df_boq.loc[idx, "TOTAL PRICE"] = tp
                    current_parent_sum += tp

            if current_parent_idx is not None and current_parent_sum > 0:
                df_boq.loc[current_parent_idx, "TOTAL PRICE"] = current_parent_sum

            sub_total = sum(parse_price(row.get("TOTAL PRICE", 0)) for _, row in df_boq.iterrows() if str(row.get("NO", "")).strip().upper() in parent_headers)
            vat_amount = sub_total * 0.11
            grand_total = sub_total + vat_amount

            display_df = df_boq.copy()
            for c in ["UNIT PRICE", "TOTAL PRICE"]:
                display_df[c] = display_df[c].apply(lambda x: f"Rp. {x:,.0f}".replace(",", ".") if parse_price(x) > 0 else "-")

            st.dataframe(display_df, use_container_width=True, hide_index=True)

            st.markdown("---")
            c_res1, c_res2 = st.columns([2, 1])
            with c_res1:
                st.success(f"✅ Tabel BOQ Aktif: **{selected_site}** | Tipe Charger: **{charging_type}** | Wilayah: **{region_normalized}**")
            with c_res2:
                st.metric(label="Total Contractor Price (Inc. VAT 11%)", value=f"Rp. {grand_total:,.0f}".replace(",", "."))

            st.subheader("3. Action & Generate PDF")
            col_btn1, col_btn2 = st.columns(2)

            with col_btn1:
                # 🚀 PERLINDUNGAN SIMPAN: Di-disable jika (Site Name + Charger Type) sudah tersimpan
                if st.button("🚀 Simpan ke Database (DB BOQ)", type="primary", disabled=is_already_saved):
                    if is_already_saved:
                        st.error(f"❌ Gagal Simpan! Kombinasi '{selected_site}' ({charging_type}) sudah terdaftar di database.")
                    else:
                        with st.spinner("Memproses penyimpanan ke DB BOQ..."):
                            boq_no = save_to_db_boq(selected_site, charging_type, sub_total, grand_total, epc_name)
                            if boq_no:
                                update_google_sheet_summary(selected_site, selected_site, sub_total, grand_total)
                                st.success(f"✅ BOQ Berhasil Dibuat dengan No: `{boq_no}` untuk **{epc_name}**!")
                                st.rerun()

            with col_btn2:
                pdf_bytes = generate_boq_pdf(
                    selected_site, site_address, charging_type, region_normalized, df_boq, sub_total, vat_amount, grand_total
                )
                st.download_button(
                    label="📥 Download PDF BOQ (1 Page Fit)",
                    data=pdf_bytes,
                    file_name=f"BOQ_{charging_type}_{region_normalized}_{selected_site.replace(' ', '_')}.pdf",
                    mime="application/pdf",
                )

    # ==========================================================================
    # TAB 2: EDIT / REUSE / RE-DOWNLOAD BOQ TERSIMPAN
    # ==========================================================================
    with tab_edit:
        st.subheader("✏️ Edit, Reuse, & Re-Download BOQ Tersimpan")
        
        saved_boq_list = get_all_saved_boq()
        
        if not saved_boq_list:
            st.info("ℹ️ Belum ada data BOQ yang tersimpan di sheet `DB BOQ`.")
        else:
            boq_options = {
                f"{item.get('BOQ No.', '-')}: {item.get('Site Name', '-')} [{item.get('Charger Type', '-')}]": item
                for item in saved_boq_list
            }
            
            selected_boq_key = st.selectbox("Pilih Nomor BOQ yang Ingin Di-edit / Re-assign / Re-Download", sorted(list(boq_options.keys())))
            selected_data = boq_options[selected_boq_key]
            
            st.markdown("---")
            st.caption(f"📌 Editing Row: `{selected_data['row_idx']}` | No. BOQ: `{selected_data.get('BOQ No.', '-')}`")
            
            old_site_name = selected_data.get("Site Name", "")
            old_charger_type = selected_data.get("Charger Type", "")
            
            # Ambil seluruh opsi tanpa filter untuk opsi re-assign
            all_site_options, all_site_map = fetch_query_site_options(exclude_saved=False)

            matched_meta = next(
                (
                    meta for label, meta in all_site_map.items() 
                    if meta["site_name"].strip().lower() == old_site_name.strip().lower()
                    and meta["charger"].strip().lower() == old_charger_type.strip().lower()
                ),
                {"address": "-", "province": "JAVA"}
            )
            
            edit_address = matched_meta["address"]
            edit_province = matched_meta["province"]
            
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                use_existing_query_site = st.checkbox("Ganti/Re-assign dengan Site Aktif dari Sheet Query", value=False)
                
                if use_existing_query_site:
                    selected_reassign_label = st.selectbox("Pilih Site & Charger Pengganti", sorted(list(all_site_options)))
                    reassign_meta = all_site_map[selected_reassign_label]
                    edit_site_name = reassign_meta["site_name"]
                    edit_charger_type = reassign_meta["charger"]
                    edit_epc_name = reassign_meta["epc"]
                    edit_address = reassign_meta["address"]
                    edit_province = reassign_meta["province"]
                    st.info(f"💡 Menimpa BOQ No. `{selected_data.get('BOQ No.', '-')}` ke Site Baru: **{edit_site_name} ({edit_charger_type})**")
                else:
                    edit_site_name = st.text_input("Site Name", value=old_site_name)
                    edit_charger_type = st.text_input("Charger Type", value=old_charger_type)
                    edit_epc_name = st.text_input("EPC Name", value=selected_data.get("EPC Name", "-"))
                
            with col_e2:
                edit_sub_total = st.number_input(
                    "BOQ Amount Exc. PPN (Rp)", 
                    value=parse_price(selected_data.get("BOQ Amount Exc. PPN", 0)), 
                    step=100000.0
                )
                
                auto_inc_ppn = edit_sub_total * 1.11
                edit_grand_total = st.number_input(
                    "BOQ Amount Inc. PPN 11% (Rp)", 
                    value=float(auto_inc_ppn),
                    step=100000.0
                )
            
            st.markdown("---")
            col_act1, col_act2 = st.columns(2)

            with col_act1:
                if st.button("💾 Save & Update BOQ Database", type="primary"):
                    with st.spinner("Memperbarui data di DB BOQ & Sum Project..."):
                        success = update_db_boq_row(
                            row_idx=selected_data["row_idx"],
                            old_site_name=old_site_name,
                            new_site_name=edit_site_name,
                            charger_capacity=edit_charger_type,
                            sub_total=edit_sub_total,
                            grand_total=edit_grand_total,
                            epc_name=edit_epc_name
                        )
                        if success:
                            st.success(f"🎉 BOQ `{selected_data.get('BOQ No.', '-')}` berhasil diperbarui ke Site: `{edit_site_name}` ({edit_charger_type})!")
                            st.rerun()

            # 📥 RE-DOWNLOAD PDF UNTUK SEMUA STATUS BOQ
            with col_act2:
                df_boq_edit, _, reg_norm = load_boq_dataframe(edit_charger_type, edit_province)
                
                if df_boq_edit is not None:
                    vat_edit = edit_sub_total * 0.11
                    pdf_bytes_edit = generate_boq_pdf(
                        edit_site_name, edit_address, edit_charger_type, reg_norm, df_boq_edit, edit_sub_total, vat_edit, edit_grand_total
                    )
                    
                    st.download_button(
                        label=f"📥 Re-Download PDF BOQ ({selected_data.get('BOQ No.', '-')})",
                        data=pdf_bytes_edit,
                        file_name=f"BOQ_{selected_data.get('BOQ No.', '').replace('/', '_')}_{edit_site_name.replace(' ', '_')}_{edit_charger_type}.pdf",
                        mime="application/pdf",
                    )


if __name__ == "__main__":
    render()