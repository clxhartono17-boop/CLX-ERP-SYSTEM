# =============================================================================
# manage_service.py
# CLX ERP -> Operation Div. -> Field Service Management
# Streamlit + Google Sheets + optional Google Drive Evidence + WhatsApp API
# =============================================================================

from __future__ import annotations
import json, math, os, re, uuid
from datetime import datetime, date, timedelta
from typing import Any, Dict, List, Optional

import pandas as pd
import requests
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# -----------------------------------------------------------------------------
# CONFIG
# -----------------------------------------------------------------------------

RESPONSE_HOURS = 1
RESOLUTION_HOURS = 12
CUSTOMER_ACCEPTANCE_HOURS = 4
DEFAULT_GEOFENCE_M = 100

SHEETS = {
    "settings": "System Settings",
    "users": "User Master",
    "customers": "Customer Master",
    "sites": "Site List Accepted",
    "with_battery": "Site List With Battery",
    "without_battery": "Site List Without Battery",
    "duplicates": "Site List Duplicate",
    "tt": "TT Database",
    "wo": "WO Database",
    "triage": "TT Triage",
    "activity": "CM Activity",
    "evidence": "CM Evidence",
    "hold": "CM Hold",
    "escalation": "CM Escalation",
    "report": "CM Report",
    "acceptance": "Customer Acceptance",
    "sla": "SLA Log",
    "wa": "WhatsApp Log",
    "audit": "System Audit Log",
    "pm": "PM Schedule",
    "pm_check": "PM Checklist",
    "calendar": "SLA Calendar",
}

HEADERS = {
"System Settings":["Key","Value","Description"],
"User Master":["User ID","Name","Email","Phone","WhatsApp","Role","Region","Active","Created At","Updated At"],
"Customer Master":["Customer ID","Customer Name","PIC Name","PIC Phone","PIC WhatsApp","PIC Email","Active","Created At","Updated At"],
"Site List Accepted":["No","Cabinet S/N","Battery Qty","Site Code","Location Name","Province","Region","Address","Latitude","Longitude","Geofence Radius (m)","PIC","PIC Contact","Status","Status Site","Created At","Updated At"],
"Site List With Battery":["No","Cabinet S/N","Battery Qty","Site Code","Location Name","Province","Region","Address","Latitude","Longitude","Geofence Radius (m)","PIC","PIC Contact","Status","Status Site","Created At","Updated At"],
"Site List Without Battery":["No","Cabinet S/N","Battery Qty","Site Code","Location Name","Province","Region","Address","Latitude","Longitude","Geofence Radius (m)","PIC","PIC Contact","Status","Status Site","Created At","Updated At"],
"Site List Duplicate":["Duplicate Key","Cabinet S/N","Location Name","Count","Rows","Detected At"],
"TT Database":["TT No","TT Date","Customer ID","Customer","Site Code","Cabinet S/N","Location Name","Province","Problem","Priority","Responsibility","Status","Response Due","Resolution Due","Response Status","SLA Status","Assigned Technician","Assigned At","Sent At","Accepted At","On The Way At","ETA","Arrived At","Check In At","Check Out At","Testing At","Report Submitted At","Customer Acceptance","Customer Acceptance Due","Hold Active","Hold Reason","Completed At","Closure Type","Created By","Created At","Updated At","Notes"],
"WO Database":["WO No","TT No","Technician ID","Technician Name","Status","Assignment Attempt","Reject Reason","Assigned At","Accepted At","On The Way At","ETA","Arrived At","Check In At","CM Start At","Testing At","Test Result","Check Out At","Report Submitted At","Revision Requested At","Completed At","Created At","Updated At"],
"TT Triage":["Triage ID","TT No","Duplicate Check","Site Validated","Priority","Responsibility","SLA Target Hours","Target Response Hours","Triage Result","Triaged By","Triaged At","Notes"],
"CM Activity":["Activity ID","TT No","WO No","Technician ID","Technician Name","Activity Type","Description","Status","Latitude","Longitude","GPS Accuracy","Distance From Site (m)","Created At"],
"CM Evidence":["Evidence ID","TT No","WO No","Technician ID","Evidence Type","File Name","File URL","Latitude","Longitude","GPS Accuracy","Distance From Site (m)","Captured At","Drive Folder"],
"CM Hold":["Hold ID","TT No","WO No","Reason Code","Reason Detail","Evidence URL","Requested By","Requested At","Status","Approved By","Approved At","Pause Start","Pause End","Paused Minutes","Resume Reason"],
"CM Escalation":["Escalation ID","TT No","WO No","Escalation Type","Level","Reason","Assigned To","Status","Created At","Acknowledged At","Resolved At","Notes"],
"CM Report":["Report ID","TT No","WO No","Technician ID","RCA","Finding","Action Taken","Material Used","Sparepart Used","Before Condition","After Condition","Recommendation","Report Status","Submitted At","Reviewed By","Reviewed At","Revision Notes","PDF URL"],
"Customer Acceptance":["Acceptance ID","TT No","WO No","Customer","Customer PIC","Status","Remark","Signature URL","Requested At","Due At","Responded At","Closure Type"],
"SLA Log":["Log ID","TT No","Event","Old Status","New Status","At","Actor","Remark"],
"WhatsApp Log":["Log ID","TT No","Recipient","Message Type","Message","Status","Provider Response","Message ID","HTTP Status","Error Code","Error Message","Retry Count","Last Retry At","Sent At"],
"System Audit Log":["Audit ID","User","Role","Action","Entity","Entity ID","Details","Created At"],
"PM Schedule":["PM No","Site Code","Cabinet S/N","Location Name","Due Date","Assigned Technician","Status","Created At","Updated At"],
"PM Checklist":["Checklist ID","PM No","Item","Result","Remark","Evidence URL","Created At"],
"SLA Calendar":["Calendar ID","Name","Working Day","Start Time","End Time","Active"],
}

# -----------------------------------------------------------------------------
# CONNECTION
# -----------------------------------------------------------------------------

@st.cache_resource(show_spinner=False)
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    try:
        if "google_creds" in st.secrets:
            creds = Credentials.from_service_account_info(dict(st.secrets["google_creds"]), scopes=scopes)
            return gspread.authorize(creds)
    except Exception:
        pass
    raw = os.getenv("GOOGLE_APPLICATION_CREDENTIALS","")
    if raw and os.path.exists(raw):
        creds = Credentials.from_service_account_file(raw, scopes=scopes)
        return gspread.authorize(creds)
    raw = str(st.secrets.get("SERVICE_ACCOUNT_JSON",""))
    if raw:
