# ==============================================================================
# good_receipt.py
# PT. CLX - SCM MODULE
# GOOD RECEIVED / GOODS RECEIPT
#
# Integration:
#   from good_receipt import render_good_receipt_module
#   render_good_receipt_module()
#
# Database:
#   DB Material IN
#   DB General IN
#
# Rules:
#   Material:
#     J:M = Receive 1
#     N:Q = Receive 2
#     R   = Total Qty Received
#     S   = End Status
#
#   General:
#     I:L = Receive 1
#     M:P = Receive 2
#     Q   = Total Qty Received
#     R   = End Status
#
# End Status:
#   Pending -> Partial -> Completed
#
# The supplied database structure supports a maximum of 2 GR stages.
# ==============================================================================

from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from typing import Any, Dict, List, Tuple

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials


# ==============================================================================
# CONFIGURATION
# ==============================================================================

APP_TITLE = "PT. CLX - Goods Receipt"

MATERIAL_DB_SHEET = "DB Material IN"
GENERAL_DB_SHEET = "DB General IN"

MATERIAL_HEADERS = [
    "No",
    "No. PO",
    "PO. Date",
    "Material Code",
    "Material Name",
    "Qty",
    "UoM",
    "Amount",
    "Total Amount",
    "Receive Date 1",
    "Qty 1",
    "Status 1",
    "PIC 1",
    "Receive Date 2",
    "Qty 2",
    "Status 2",
    "PIC 2",
    "Total Qty Received",
    "End Status",
]

GENERAL_HEADERS = [
    "No",
    "No. PO",
    "PO. Date",
    "Material/Item Name",
    "Qty",
    "UoM",
    "Amount",
    "Total Amount",
    "Receive Date 1",
    "Qty 1",
    "Status 1",
    "PIC 1",
    "Receive Date 2",
    "Qty 2",
    "Status 2",
    "PIC 2",
    "Total Qty Received",
    "End Status",
]


# ==============================================================================
# HELPERS
# ==============================================================================

def clean_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def normalize_key(value: Any) -> str:
    return " ".join(
        clean_text(value).lower().split()
    )


def parse_decimal(
    value: Any,
    default: Decimal = Decimal("0"),
) -> Decimal:
    if value is None or value == "":
        return default

    if isinstance(value, Decimal):
        return value

    if isinstance(value, (int, float)):
        try:
            return Decimal(str(value))
        except Exception:
            return default

    s = clean_text(value)

    if not s:
        return default

    s = (
        s.replace("Rp", "")
        .replace("rp", "")
        .replace(" ", "")
    )

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 2 or (
            len(parts) == 2 and len(parts[1]) == 3
        ):
            s = s.replace(".", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) > 2 or (
            len(parts) == 2 and len(parts[1]) == 3
        ):
            s = s.replace(",", "")

    try:
        return Decimal(s)
    except InvalidOperation:
        return default


def format_qty(value: Any) -> str:
    d = parse_decimal(value)

    if d == d.to_integral():
        return f"{int(d):,}"

    return (
        f"{d:,.3f}"
        .rstrip("0")
        .rstrip(".")
    )


def format_rupiah(value: Any) -> str:
    return f"Rp{int(parse_decimal(value)):,.0f}"


# ==============================================================================
# GOOGLE SHEETS
# ==============================================================================

@st.cache_resource
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    sa = None

    try:
        sa = st.secrets.get("SERVICE_ACCOUNT_JSON")
    except Exception:
        pass

    if sa:
        if isinstance(sa, str):
            sa = json.loads(sa)

        return gspread.authorize(
            Credentials.from_service_account_info(
                sa,
                scopes=scopes,
            )
        )

    raw_json = os.getenv(
        "SERVICE_ACCOUNT_JSON",
        "",
    ).strip()

    if raw_json:
        return gspread.authorize(
            Credentials.from_service_account_info(
                json.loads(raw_json),
                scopes=scopes,
            )
        )

    credentials_file = os.getenv(
        "GOOGLE_APPLICATION_CREDENTIALS",
        "",
    ).strip()

    if credentials_file and os.path.exists(
        credentials_file
    ):
        return gspread.authorize(
            Credentials.from_service_account_file(
                credentials_file,
                scopes=scopes,
            )
        )

    raise RuntimeError(
        "Google credential tidak ditemukan."
    )


@st.cache_resource
def get_spreadsheet():
    spreadsheet_id = ""

    try:
        spreadsheet_id = clean_text(
            st.secrets.get("SPREADSHEET_ID")
        )
    except Exception:
        pass

    if not spreadsheet_id:
        spreadsheet_id = os.getenv(
            "SPREADSHEET_ID",
            "",
        ).strip()

    if not spreadsheet_id:
        raise RuntimeError(
            "SPREADSHEET_ID belum dikonfigurasi."
        )

    return get_gspread_client().open_by_key(
        spreadsheet_id
    )


def get_worksheet(name: str):
    try:
        return get_spreadsheet().worksheet(name)
    except gspread.WorksheetNotFound:
        raise RuntimeError(
            f'Sheet "{name}" tidak ditemukan.'
        )


def retry_call(func, *args, retries: int = 5, **kwargs):
    last_error = None

    for attempt in range(retries):
        try:
            return func(*args, **kwargs)
        except Exception as exc:
            last_error = exc
            text = str(exc).lower()

            retryable = any(
                token in text
                for token in (
                    "429",
                    "quota",
                    "rate limit",
                    "503",
                    "500",
                    "timed out",
                )
            )

            if not retryable or attempt == retries - 1:
                raise

            time.sleep(min(8, 2 ** attempt))

    raise last_error


def get_all_records_with_row_numbers(
    sheet_name: str,
) -> List[Tuple[int, Dict[str, Any]]]:
    ws = get_worksheet(sheet_name)

    values = retry_call(
        ws.get_all_values
    )

    if not values:
        return []

    headers = [
        clean_text(x)
        for x in values[0]
    ]

    result = []

    for row_number, row in enumerate(
        values[1:],
        start=2,
    ):
        padded = list(row) + [""] * max(
            0,
            len(headers) - len(row),
        )

        result.append((
            row_number,
            {
                headers[i]: padded[i]
                for i in range(len(headers))
            },
        ))

    return result


def update_cells(
    sheet_name: str,
    row_number: int,
    updates: Dict[str, Any],
):
    if not updates:
        return

    ws = get_worksheet(sheet_name)

    data = []

    for column, value in updates.items():
        data.append({
            "range": f"{column}{row_number}",
            "values": [[value]],
        })

    retry_call(
        ws.batch_update,
        data,
        value_input_option="USER_ENTERED",
    )


# ==============================================================================
# DATA RETRIEVAL
# ==============================================================================

def get_open_gr_rows(
    po_type: str,
) -> List[Tuple[int, Dict[str, Any]]]:
    sheet_name = (
        MATERIAL_DB_SHEET
        if po_type == "Material"
        else GENERAL_DB_SHEET
    )

    rows = get_all_records_with_row_numbers(
        sheet_name
    )

    result = []

    for row_number, record in rows:
        status = normalize_key(
            record.get("End Status", "")
        )

        if status != "completed":
            result.append(
                (row_number, record)
            )

    return result


def get_gr_label(
    po_type: str,
    record: Dict[str, Any],
) -> str:
    po_no = clean_text(
        record.get("No. PO")
    )

    qty_po = format_qty(
        record.get("Qty")
    )

    received = format_qty(
        record.get("Total Qty Received")
    )

    status = (
        clean_text(
            record.get("End Status")
        )
        or "Pending"
    )

    if po_type == "Material":
        code = clean_text(
            record.get("Material Code")
        )

        name = clean_text(
            record.get("Material Name")
        )

        return (
            f"{po_no} | {code} | {name} | "
            f"PO Qty: {qty_po} | "
            f"Received: {received} | "
            f"{status}"
        )

    name = clean_text(
        record.get("Material/Item Name")
    )

    return (
        f"{po_no} | {name} | "
        f"PO Qty: {qty_po} | "
        f"Received: {received} | "
        f"{status}"
    )


# ==============================================================================
# GOODS RECEIPT ENGINE
# ==============================================================================

@dataclass
class GRResult:
    row_number: int
    po_no: str
    item_name: str
    qty_po: Decimal
    qty_received_now: Decimal
    total_received: Decimal
    balance: Decimal
    receipt_stage: int
    receipt_status: str
    end_status: str


def process_goods_received(
    po_type: str,
    row_number: int,
    record: Dict[str, Any],
    receive_date: str,
    qty_received: Decimal,
    pic: str,
) -> GRResult:
    if not clean_text(receive_date):
        raise ValueError(
            "Receive Date wajib diisi."
        )

    if qty_received <= 0:
        raise ValueError(
            "Qty Actual Received harus lebih besar dari 0."
        )

    if not clean_text(pic):
        raise ValueError(
            "PIC Penerima wajib diisi."
        )

    if po_type not in (
        "Material",
        "General",
    ):
        raise ValueError(
            f"PO Type tidak valid: {po_type}"
        )

    sheet_name = (
        MATERIAL_DB_SHEET
        if po_type == "Material"
        else GENERAL_DB_SHEET
    )

    qty_po = parse_decimal(
        record.get("Qty")
    )

    qty1 = parse_decimal(
        record.get("Qty 1")
    )

    qty2 = parse_decimal(
        record.get("Qty 2")
    )

    date1 = clean_text(
        record.get("Receive Date 1")
    )

    date2 = clean_text(
        record.get("Receive Date 2")
    )

    total_before = qty1 + qty2

    if total_before >= qty_po:
        raise ValueError(
            "PO item ini sudah Completed."
        )

    # --------------------------------------------------------------------------
    # Determine receiving stage.
    # --------------------------------------------------------------------------

    if not date1:
        stage = 1

        if po_type == "Material":
            date_col = "J"
            qty_col = "K"
            status_col = "L"
            pic_col = "M"
            total_col = "R"
            end_col = "S"
        else:
            date_col = "I"
            qty_col = "J"
            status_col = "K"
            pic_col = "L"
            total_col = "Q"
            end_col = "R"

    elif not date2:
        stage = 2

        if po_type == "Material":
            date_col = "N"
            qty_col = "O"
            status_col = "P"
            pic_col = "Q"
            total_col = "R"
            end_col = "S"
        else:
            date_col = "M"
            qty_col = "N"
            status_col = "O"
            pic_col = "P"
            total_col = "Q"
            end_col = "R"

    else:
        raise ValueError(
            "PO item ini sudah menggunakan "
            "2 tahap penerimaan. Struktur database "
            "saat ini tidak menyediakan GR tahap ke-3."
        )

    # --------------------------------------------------------------------------
    # Validate over-receiving.
    #
    # The original requirement says >= Qty PO becomes Completed.
    # We allow equal / over-receipt, but explicitly show the over-received
    # quantity in the result.
    # --------------------------------------------------------------------------

    new_total = total_before + qty_received
    balance = qty_po - new_total

    if new_total >= qty_po:
        receipt_status = "Completed"
        end_status = "Completed"
    else:
        receipt_status = "Partial"
        end_status = "Partial"

    update_cells(
        sheet_name,
        row_number,
        {
            date_col: receive_date,
            qty_col: float(qty_received),
            status_col: receipt_status,
            pic_col: pic.strip(),
            total_col: float(new_total),
            end_col: end_status,
        },
    )

    if po_type == "Material":
        item_name = clean_text(
            record.get("Material Name")
        )
    else:
        item_name = clean_text(
            record.get("Material/Item Name")
        )

    return GRResult(
        row_number=row_number,
        po_no=clean_text(
            record.get("No. PO")
        ),
        item_name=item_name,
        qty_po=qty_po,
        qty_received_now=qty_received,
        total_received=new_total,
        balance=balance,
        receipt_stage=stage,
        receipt_status=receipt_status,
        end_status=end_status,
    )


# ==============================================================================
# UI
# ==============================================================================

def _render_gr_history(
    po_type: str,
    record: Dict[str, Any],
):
    if po_type == "Material":
        receive1 = {
            "Date": clean_text(
                record.get("Receive Date 1")
            ),
            "Qty": format_qty(
                record.get("Qty 1")
            ),
            "Status": clean_text(
                record.get("Status 1")
            ),
            "PIC": clean_text(
                record.get("PIC 1")
            ),
        }

        receive2 = {
            "Date": clean_text(
                record.get("Receive Date 2")
            ),
            "Qty": format_qty(
                record.get("Qty 2")
            ),
            "Status": clean_text(
                record.get("Status 2")
            ),
            "PIC": clean_text(
                record.get("PIC 2")
            ),
        }

    else:
        receive1 = {
            "Date": clean_text(
                record.get("Receive Date 1")
            ),
            "Qty": format_qty(
                record.get("Qty 1")
            ),
            "Status": clean_text(
                record.get("Status 1")
            ),
            "PIC": clean_text(
                record.get("PIC 1")
            ),
        }

        receive2 = {
            "Date": clean_text(
                record.get("Receive Date 2")
            ),
            "Qty": format_qty(
                record.get("Qty 2")
            ),
            "Status": clean_text(
                record.get("Status 2")
            ),
            "PIC": clean_text(
                record.get("PIC 2")
            ),
        }

    st.markdown("### History Penerimaan")

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("**Receive 1**")
        st.json(receive1)

    with c2:
        st.markdown("**Receive 2**")
        st.json(receive2)


def _render_one_gr(
    po_type: str,
):
    title = (
        "📥 Goods Received - Material"
        if po_type == "Material"
        else "📥 Goods Received - General"
    )

    st.subheader(title)

    try:
        rows = get_open_gr_rows(
            po_type
        )
    except Exception as exc:
        st.error(
            f"Gagal membaca database: {exc}"
        )
        return

    if not rows:
        st.success(
            "Tidak ada item Pending / Partial."
        )
        return

    labels = [
        get_gr_label(
            po_type,
            record,
        )
        for _, record in rows
    ]

    selected = st.selectbox(
        "Pilih PO / Item",
        labels,
        key=f"gr_select_{po_type}",
    )

    index = labels.index(selected)

    row_number, record = rows[index]

    st.markdown("### Detail PO")

    if po_type == "Material":
        c1, c2, c3, c4, c5, c6 = st.columns(6)

        c1.metric(
            "No. PO",
            clean_text(
                record.get("No. PO")
            ),
        )

        c2.metric(
            "Material Code",
            clean_text(
                record.get("Material Code")
            ),
        )

        c3.metric(
            "Material",
            clean_text(
                record.get("Material Name")
            ),
        )

        c4.metric(
            "PO Qty",
            format_qty(
                record.get("Qty")
            ),
        )

        c5.metric(
            "Received",
            format_qty(
                record.get(
                    "Total Qty Received"
                )
            ),
        )

        c6.metric(
            "Status",
            clean_text(
                record.get("End Status")
            ) or "Pending",
        )

    else:
        c1, c2, c3, c4, c5 = st.columns(5)

        c1.metric(
            "No. PO",
            clean_text(
                record.get("No. PO")
            ),
        )

        c2.metric(
            "Item",
            clean_text(
                record.get(
                    "Material/Item Name"
                )
            ),
        )

        c3.metric(
            "PO Qty",
            format_qty(
                record.get("Qty")
            ),
        )

        c4.metric(
            "Received",
            format_qty(
                record.get(
                    "Total Qty Received"
                )
            ),
        )

        c5.metric(
            "Status",
            clean_text(
                record.get("End Status")
            ) or "Pending",
        )

    qty_po = parse_decimal(
        record.get("Qty")
    )

    total_received = parse_decimal(
        record.get("Total Qty Received")
    )

    balance = qty_po - total_received

    if balance < 0:
        balance_display = (
            f"Over Received "
            f"{format_qty(abs(balance))}"
        )
    else:
        balance_display = format_qty(
            balance
        )

    st.info(
        f"**PO Qty:** {format_qty(qty_po)}  |  "
        f"**Total Received:** {format_qty(total_received)}  |  "
        f"**Remaining:** {balance_display}"
    )

    with st.expander(
        "History Penerimaan",
        expanded=True,
    ):
        _render_gr_history(
            po_type,
            record,
        )

    st.markdown("### Input Goods Receipt")

    c1, c2, c3 = st.columns(3)

    with c1:
        receive_date = st.date_input(
            "Receive Date *",
            value=date.today(),
            key=f"gr_date_{po_type}",
        )

    with c2:
        receive_qty = st.number_input(
            "Qty Actual Received *",
            min_value=0.0,
            value=1.0,
            step=1.0,
            key=f"gr_qty_{po_type}",
        )

    with c3:
        pic = st.text_input(
            "PIC Penerima *",
            key=f"gr_pic_{po_type}",
        )

    st.caption(
        "Status otomatis: "
        "Partial jika total penerimaan < Qty PO; "
        "Completed jika total penerimaan >= Qty PO."
    )

    if st.button(
        "📥 Process Goods Receipt",
        type="primary",
        use_container_width=True,
        key=f"gr_submit_{po_type}",
    ):
        try:
            result = process_goods_received(
                po_type=po_type,
                row_number=row_number,
                record=record,
                receive_date=receive_date.strftime(
                    "%d/%b/%y"
                ),
                qty_received=Decimal(
                    str(receive_qty)
                ),
                pic=pic,
            )

            if result.balance < 0:
                over = format_qty(
                    abs(result.balance)
                )

                st.warning(
                    f"GR berhasil, tetapi terdapat "
                    f"over-receipt sebanyak {over}."
                )

            st.success(
                f"GR berhasil disimpan. "
                f"PO: {result.po_no} | "
                f"Item: {result.item_name} | "
                f"Tahap: {result.receipt_stage} | "
                f"Qty GR: {format_qty(result.qty_received_now)} | "
                f"Total Received: {format_qty(result.total_received)} | "
                f"End Status: {result.end_status}"
            )

            st.rerun()

        except Exception as exc:
            st.error(
                f"Gagal memproses Goods Receipt: {exc}"
            )


# ==============================================================================
# PUBLIC ENTRY POINT
# ==============================================================================

def render_good_receipt_module():
    """
    Main entry point for CLX ERP.

    Recommended app.py routing:

        if selected_menu == "Good Receipt":
            from good_receipt import render_good_receipt_module
            render_good_receipt_module()
    """

    st.title("📦 Goods Receipt")

    tab_material, tab_general = st.tabs([
        "📦 GR Material",
        "🧾 GR General",
    ])

    with tab_material:
        _render_one_gr("Material")

    with tab_general:
        _render_one_gr("General")


# Aliases for flexible router integration.
render = render_good_receipt_module
show_good_receipt_page = render_good_receipt_module
show = render_good_receipt_module


if __name__ == "__main__":
    render_good_receipt_module()
