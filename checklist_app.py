import hmac
import html
import re
import uuid
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

# ─── Config ───────────────────────────────────────────────────────────────────
SHEET_ID    = "1jzQpd3MpCp0kT1djnJo_WfQMsbe-reK1oGtvgn6vE1Q"
SUPERVISORS = ["Ashutosh Singh", "Anas", "Sarosh"]
SCOPES      = ["https://www.googleapis.com/auth/spreadsheets"]
IST         = ZoneInfo("Asia/Kolkata")          # servers run in UTC; we want Noida time

EXC_TAB     = "Exceptions Log"
EXC_HEADER  = ["ID", "Date", "Source", "Issue", "Action Taken", "Escalated To",
               "Follow-up By", "Follow-up Date", "Status", "Resolved On"]
EXC_KEYS    = ["id", "date", "source", "issue", "action", "escalated",
               "fby", "fdate", "status", "resolved_on"]
STATUSES    = ["Open", "In Progress", "Resolved"]

# Samasya portal colours
GREEN, NAVY, AMBER, ORANGE = "#00CC66", "#0A0A7A", "#F0C24B", "#FF9800"
BLUE, DARK, RED, OKGREEN   = "#6FA8DC", "#2B2B2B", "#E53935", "#43A047"
STATUS_COLOUR = {"Open": RED, "In Progress": ORANGE, "Resolved": OKGREEN}


def tab_name(kind, d):
    return f"{kind}({d:%d-%m-%Y})"

def legacy_name(kind, d):
    """Old tabs (imported from the Excel file) had no year; they are all from 2026."""
    return f"{kind}({d:%d-%m})" if d.year == 2026 else None


def now_ist():
    return datetime.now(IST)


# ─── Checklist definitions (single source of truth for UI + sheet output) ─────
def I(label, key, kind="yn", crit=False, ph=""):
    return {"label": label, "key": key, "kind": kind, "crit": crit, "ph": ph}

DS_SECTIONS = [
    ("1. ENVIRONMENT & WAREHOUSE READINESS", [
        I("AC / climate control working",             "ac"),
        I("Humidifier / dehumidifier working",        "humidifier"),
        I("Pest / insect / rodent activity observed", "pest", "ny"),
        I("Aisles / walking paths clear",             "aisles"),
        I("Shelves / work areas tidy",                "shelves"),
        I("Packing tables clean",                     "packing_tables"),
    ]),
    ("2. STOCK & STORAGE READINESS", [
        I("Stock left in temporary / unusual locations", "unusual"),
        I("Stock on floor that should be shelved",       "floor"),
        I("Moves pending (move list entries)",          "moves",     "num", True),
        I("Cancelled orders pending",                   "cancelled", "num", True),
        I("Open stock merged to base",                  "open_merged"),
        I("Damaged stock handled",                      "damaged_handled"),
    ]),
    ("3. SAMASYA / OPERATIONAL QUEUE (at time of check)", [
        I("Orders still in ASSEMBLING", "assembling",    "num", True),
        I("Orders still in PACKING",    "packing",       "num", True),
        I("RTO queue requiring action", "rto",           "num"),
        I("DTO queue requiring action", "dto",           "num"),
        I("Notes pending",              "notes_pending", "num"),
    ]),
    ("4. EQUIPMENT & PACKING READINESS", [
        I("Tablets working",                             "tablets",  crit=True),
        I("Scanner working",                             "scanner",  crit=True),
        I("Printers working",                            "printers", crit=True),
        I("Weighing scales working",                     "scales",   crit=True),
        I("Thermal rolls available",                     "thermal"),
        I("Boxes available",                             "boxes",    crit=True),
        I("Pouches available",                           "pouches",  crit=True),
        I("Bubble wrap / protective material available", "bubble_wrap", crit=True),
        I("Tape available",                              "tape"),
        I("Packing station ready",                       "packing_station"),
        I("Any equipment issue / note",                  "eq_issue", "text", True, "None / describe"),
        I("Cameras working",                             "cameras"),
    ]),
]

DE_SECTIONS = [
    ("1. ORDER CLOSURE / WORKLOAD CARRIED FORWARD  ★ CRITICAL", [
        I("Orders remaining in ASSEMBLING",         "assembling", "num", True),
        I("Orders remaining in PACKING",            "packing",    "num", True),
        I("Cancelled orders pending",               "cancelled",  "num", True),
        I("Courier handovers completed / recorded", "courier"),
    ]),
    ("2. MOVES / CHECKING", [
        I("Moves pending (move list entries)", "moves",    "num", True),
        I("Checking pending",                  "checking", "num"),
    ]),
    ("3. RTO / DTO", [
        I("RTOs left",                           "rto_left", "num"),
        I("DTOs left",                           "dto_left", "num"),
        I("Returned stock requiring restocking", "restocking"),
    ]),
    ("4. WAREHOUSE RESET  ★ CRITICAL (Close Clean)", [
        I("Loose stock returned to appropriate location", "loose_stock",       crit=True),
        I("LT → ST movement completed",                   "lt_st",             crit=True),
        I("Packing materials replenished",                "packing_materials", crit=True),
        I("Picking / assembly area reset",                "assembly_reset",    crit=True),
        I("Packing tables cleared",                       "packing_tables",    crit=True),
        I("Aisles clear",                                 "aisles",            crit=True),
    ]),
    ("5. EQUIPMENT", [
        I("Tablets kept on charge",           "tablets"),
        I("Scanners kept on charge / stored", "scanner"),
        I("Printers left ready for next day", "printers"),
        I("Weighing scales left ready",       "scales"),
        I("Camera working",                   "cameras"),
    ]),
    ("6. ENVIRONMENT", [
        I("AC / climate control set for overnight",      "ac"),
        I("Humidifier / dehumidifier set appropriately", "humidifier"),
        I("Pest / insect / rodent activity observed",    "pest", "ny"),
    ]),
]


# ─── Google Sheets plumbing ───────────────────────────────────────────────────
@st.cache_resource
def get_client():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
    return gspread.authorize(creds)

def get_spreadsheet():
    return get_client().open_by_key(SHEET_ID)

def tab_exists_fresh(name):
    return name in [w.title for w in get_spreadsheet().worksheets()]

@st.cache_data(ttl=20, show_spinner=False)
def existing_tabs():
    try:
        return [w.title for w in get_spreadsheet().worksheets()]
    except Exception:
        return []

@st.cache_data(ttl=20, show_spinner=False)
def load_tab(name):
    try:
        return get_spreadsheet().worksheet(name).get_all_values()
    except gspread.WorksheetNotFound:
        return None

def explain(ex):
    """Readable error text, with the likely fix for the usual Google Sheets problems."""
    name, msg = type(ex).__name__, str(ex).strip()
    low = (name + " " + msg).lower()
    try:
        email = st.secrets["gcp_service_account"]["client_email"]
    except Exception:
        email = "the service account email"
    hint = ""
    if "secret" in low or isinstance(ex, KeyError):
        hint = " The Google credentials are missing from the Streamlit app's Secrets."
    elif "notfound" in low or "404" in low or "not found" in low:
        hint = (f" The app can't see this Google Sheet. Share it with {email} as Editor, "
                "and check SHEET_ID in the code matches the sheet's link.")
    elif "403" in low or "permission" in low:
        hint = (f" Permission problem. Share the sheet with {email} as Editor, and make sure "
                "the Google Sheets API is enabled in the Google Cloud project.")
    elif "not supported for this document" in low or "failed_precondition" in low:
        hint = " This file is still an Excel (.xlsx) file. In Google Sheets use File > Save as Google Sheets, then use the new link's ID."
    return f"{name}: {msg or '(no message)'}.{hint}"


def clear_caches():
    existing_tabs.clear()
    load_tab.clear()
    load_exceptions.clear()


# ─── Building sheet rows ──────────────────────────────────────────────────────
BLANK = [""]

def sections_to_rows(sections, vals):
    rows = []
    for title, items in sections:
        rows.append([title, "", "", ""])
        rows.append(["Check / Field", "Critical?", "Entry", "Notes / Details"])
        for it in items:
            v, note = vals[it["key"]]
            rows.append([it["label"], "CRITICAL" if it["crit"] else "", str(v), note])
        rows.append(BLANK)
    return rows

def exceptions_to_rows(title, exceptions):
    rows = [[title], ["Issue / Exception", "Action Taken", "Escalated To"]]
    logged = [e for e in exceptions if e["issue"] or e["action"] or e["escalated"]]
    for e in logged:
        rows.append([e["issue"], e["action"], e["escalated"]])
    if not logged:
        rows.append(["None"])
    rows.append(["Follow-up for these is tracked in the Exceptions Tracker."])
    rows.append(BLANK)
    return rows

def format_tab(sh, ws, rows):
    """Bold headers / amber section bands / red CRITICAL - cosmetic only."""
    sid, reqs = ws.id, []

    def cell_fmt(r, c0, c1, fmt, fields):
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r + 1,
                      "startColumnIndex": c0, "endColumnIndex": c1},
            "cell": {"userEnteredFormat": fmt}, "fields": fields}})

    amber = {"red": 0.94, "green": 0.76, "blue": 0.29}
    grey  = {"red": 0.92, "green": 0.92, "blue": 0.92}
    for i, r in enumerate(rows):
        first = r[0] if r else ""
        filled = sum(1 for c in r if str(c).strip())
        if i == 0:
            cell_fmt(i, 0, 6, {"textFormat": {"bold": True, "fontSize": 13}}, "userEnteredFormat.textFormat")
        elif filled == 1 and re.match(r"^(\d+\.|EXCEPTIONS)", first):
            cell_fmt(i, 0, 6, {"textFormat": {"bold": True}, "backgroundColor": amber},
                     "userEnteredFormat(textFormat,backgroundColor)")
        elif first in ("Check / Field", "Issue / Exception"):
            cell_fmt(i, 0, 6, {"textFormat": {"bold": True}, "backgroundColor": grey},
                     "userEnteredFormat(textFormat,backgroundColor)")
        elif len(r) > 1 and r[1] == "CRITICAL":
            cell_fmt(i, 1, 2, {"textFormat": {"bold": True, "foregroundColor": {"red": 0.9, "green": 0.1, "blue": 0.1}}},
                     "userEnteredFormat.textFormat")
    for col, px in enumerate([380, 90, 150, 300, 110, 150]):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": col, "endIndex": col + 1},
            "properties": {"pixelSize": px}, "fields": "pixelSize"}})
    sh.batch_update({"requests": reqs})

def write_tab(name, rows, legacy=None):
    sh = get_spreadsheet()
    width = max(6, max(len(r) for r in rows))
    try:
        ws = sh.worksheet(name)
        ws.clear()
        ws.resize(rows=len(rows) + 5, cols=width)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=name, rows=len(rows) + 5, cols=width)
    ws.batch_update([{"range": "A1", "values": rows}], value_input_option="RAW")
    try:
        format_tab(sh, ws, rows)
    except Exception:
        pass                                     # formatting is cosmetic; never fail a save over it
    if legacy and legacy != name:                # replaced an old year-less tab -> drop it
        try:
            sh.del_worksheet(sh.worksheet(legacy))
        except gspread.WorksheetNotFound:
            pass


# ─── Exceptions Log (permanent sheet) ─────────────────────────────────────────
def get_log_ws(sh):
    try:
        return sh.worksheet(EXC_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=EXC_TAB, rows=500, cols=len(EXC_HEADER))
        ws.batch_update([{"range": "A1", "values": [EXC_HEADER]}], value_input_option="RAW")
        try:
            ws.format("A1:J1", {"textFormat": {"bold": True}})
        except Exception:
            pass
        return ws

def sync_exceptions(day, source, exceptions):
    """Replace this day+source's rows in the log with the freshly submitted ones."""
    sh = get_spreadsheet()
    ws = get_log_ws(sh)
    old = ws.get_all_values()
    body = [(r + [""] * 10)[:10] for r in old[1:] if any(r)]
    kept = [r for r in body if not (r[1] == day and r[2] == source)]
    new = [[uuid.uuid4().hex[:8], day, source, e["issue"], e["action"], e["escalated"], "", "", "Open", ""]
           for e in exceptions if e["issue"] or e["action"] or e["escalated"]]
    final = [EXC_HEADER] + kept + new
    if len(final) + 5 > ws.row_count:
        ws.resize(rows=len(final) + 200)
    ws.batch_update([{"range": "A1", "values": final}], value_input_option="RAW")
    if len(old) > len(final):                    # shrunk (overwrite) - blank out the leftover rows
        ws.batch_clear([f"A{len(final) + 1}:J{len(old)}"])
    return len(new)

@st.cache_data(ttl=20, show_spinner=False)
def load_exceptions():
    ws = get_log_ws(get_spreadsheet())
    out = []
    for r in ws.get_all_values()[1:]:
        r = (r + [""] * 10)[:10]
        if any(r):
            e = dict(zip(EXC_KEYS, r))
            e["status"] = e["status"] or "Open"
            out.append(e)
    out.sort(key=lambda e: e["date"])
    return out

def update_exception(exc_id, fby, fdate, status):
    ws = get_log_ws(get_spreadsheet())
    ids = ws.col_values(1)
    if exc_id not in ids:
        raise ValueError("This exception no longer exists (its day may have been re-submitted). Hit Refresh.")
    r = ids.index(exc_id) + 1
    resolved_on = now_ist().date().isoformat() if status == "Resolved" else ""
    ws.batch_update([{"range": f"G{r}:J{r}", "values": [[fby, fdate, status, resolved_on]]}],
                    value_input_option="RAW")


# ─── Summary Log (one row per Day Start / Day End) ────────────────────────────
SUMMARY_TAB = "Summary Log"
SUM_HEADER  = ["Date", "Type", "Time", "Orders in assembly", "Orders in packing", "All pickups done",
               "Cancelled orders pending", "Moves pending", "Items left in checking", "RTOs left",
               "DTOs left", "Tablets on charge", "Scanners on charge", "Pest seen", "AC working",
               "Humidifier working", "Exception count"]
SUM_COLS    = ["date", "type", "time", "assembling", "packing", "pickups", "cancelled", "moves",
               "checking", "rto", "dto", "tablets", "scanners", "pest", "ac", "humidifier", "exc"]

def _val(vals, key):
    return vals[key][0] if key in vals else ""

def summary_row(kind, d, time_str, vals, n_exc):
    """Builds the Summary Log row (a dict keyed by SUM_COLS) from a submitted checklist."""
    row = {k: "" for k in SUM_COLS}
    row.update(date=d.isoformat(), type=kind, time=time_str, exc=n_exc,
               assembling=_val(vals, "assembling"), packing=_val(vals, "packing"),
               cancelled=_val(vals, "cancelled"), moves=_val(vals, "moves"),
               pest=_val(vals, "pest"))
    if kind == "Day Start":
        row.update(rto=_val(vals, "rto"), dto=_val(vals, "dto"),
                   ac=_val(vals, "ac"), humidifier=_val(vals, "humidifier"))
    else:
        row.update(pickups=_val(vals, "courier"), checking=_val(vals, "checking"),
                   rto=_val(vals, "rto_left"), dto=_val(vals, "dto_left"),
                   tablets=_val(vals, "tablets"), scanners=_val(vals, "scanner"))
    return row

def _to_date(v):
    """A date cell may come back as text (2026-11-13, 13-11-2026) or as a Sheets serial number."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return date(1899, 12, 30) + timedelta(days=int(v))
    s = str(v).strip()
    for f in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None

def _clean_cell(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    if hasattr(v, "item"):                       # numpy scalar -> plain python
        v = v.item()
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return v

def format_summary(sh, ws):
    sid = ws.id
    amber = {"red": 0.94, "green": 0.76, "blue": 0.29}
    def col_fmt(c, ntype, pattern):
        return {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 1, "startColumnIndex": c, "endColumnIndex": c + 1},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": ntype, "pattern": pattern}}},
            "fields": "userEnteredFormat.numberFormat"}}
    reqs = [
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True}, "backgroundColor": amber,
                                           "wrapStrategy": "WRAP", "verticalAlignment": "MIDDLE"}},
            "fields": "userEnteredFormat(textFormat,backgroundColor,wrapStrategy,verticalAlignment)"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 1, "frozenColumnCount": 2}},
            "fields": "gridProperties(frozenRowCount,frozenColumnCount)"}},
        col_fmt(0, "DATE", "dd-mm-yyyy"),
        col_fmt(2, "TIME", "HH:mm"),
    ]
    for c, px in enumerate([95, 85, 65] + [105] * (len(SUM_HEADER) - 3)):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": c, "endIndex": c + 1},
            "properties": {"pixelSize": px}, "fields": "pixelSize"}})
    sh.batch_update({"requests": reqs})

def get_summary_ws(sh):
    try:
        return sh.worksheet(SUMMARY_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SUMMARY_TAB, rows=500, cols=len(SUM_HEADER))
        ws.batch_update([{"range": "A1", "values": [SUM_HEADER]}], value_input_option="RAW")
        try:
            format_summary(sh, ws)
        except Exception:
            pass                                 # cosmetic only
        return ws

def upsert_summary(new_rows):
    """Adds the given rows to the Summary Log, replacing any row with the same date + type."""
    sh = get_spreadsheet()
    ws = get_summary_ws(sh)
    try:
        old = ws.get_all_values(value_render_option="UNFORMATTED_VALUE")
    except TypeError:
        old = ws.get_all_values()
    n = len(SUM_HEADER)
    body = [(list(r) + [""] * n)[:n] for r in old[1:] if any(str(c).strip() for c in r)]
    fresh = [[_clean_cell(r[k]) for k in SUM_COLS] for r in new_rows]
    replaced = {(_to_date(r[0]), r[1]) for r in fresh}
    kept = [r for r in body if (_to_date(r[0]), r[1]) not in replaced]

    def sort_key(r):
        return (_to_date(r[0]) or date.max, 0 if r[1] == "Day Start" else 1)

    final = sorted(kept + fresh, key=sort_key)
    for r in final:                              # one date format everywhere (Sheets reads ISO as a date)
        d = _to_date(r[0])
        if d:
            r[0] = d.isoformat()
    values = [SUM_HEADER] + final
    if len(values) + 5 > ws.row_count:
        ws.resize(rows=len(values) + 200)
    ws.batch_update([{"range": "A1", "values": values}], value_input_option="USER_ENTERED")

def summary_keys():
    """(date, type) pairs already in the Summary Log - without creating the tab."""
    try:
        ws = get_spreadsheet().worksheet(SUMMARY_TAB)
    except gspread.WorksheetNotFound:
        return set()
    try:
        rows = ws.get_all_values(value_render_option="UNFORMATTED_VALUE")
    except TypeError:
        rows = ws.get_all_values()
    return {(_to_date(r[0]), r[1]) for r in rows[1:] if r and str(r[0]).strip()}


# ─── One-time fill of the Summary Log from the older sheets ───────────────────
OLD_TAB_RE = re.compile(r"^(Day Start|Day End)\((\d{2})-(\d{2})(?:-(\d{4}))?\)")

def _num(s):
    m = re.match(r"^\s*(\d+)\s*$", str(s))
    return int(m.group(1)) if m else ""

def _yes_no(s):
    t = str(s).strip().lower()
    if re.match(r"(yes|y)\b", t):
        return "Yes"
    if re.match(r"(no|n)\b", t):
        return "No"
    return ""

def _pest(s):
    t = str(s).strip().lower()
    if not t:
        return ""
    r = _yes_no(t)
    if r:
        return r
    if re.match(r"(none|nil|nothing|n/?a|-)", t):
        return "No"
    return "Yes"                                 # something was described

def _humidifier(s):
    t = str(s).strip().lower()
    if not t:
        return ""
    if re.search(r"\b(not|stopped|stop|off|no|down|broken)\b", t):
        return "No"
    if re.search(r"work|\bon\b|yes|ok|running", t):
        return "Yes"
    return ""

def _moves(s):
    t = str(s).strip()
    if not t:
        return ""
    if re.fullmatch(r"\d+", t):
        return int(t)
    if t.lower() in ("none", "nil", "no", "zero", "n/a", "na", "-"):
        return 0
    parts = [p for p in re.split(r"[,;/&\s]+|\band\b", t) if p]
    if parts and all(re.fullmatch(r"[A-Za-z]{1,3}\d{1,3}[A-Za-z]?", p) for p in parts):
        return len(parts)                        # a list of cubicles like "A3, B1"
    return ""                                    # can't tell - left blank for you to fill

def _time_of(s):
    m = re.search(r"(\d{1,2}):(\d{2})(?::\d{2})?\s*([ap]m)?", str(s).lower())
    if not m:
        return ""
    h, mi, ap = int(m.group(1)), int(m.group(2)), m.group(3)
    if ap == "pm" and h < 12:
        h += 12
    if ap == "am" and h == 12:
        h = 0
    return f"{h:02d}:{mi:02d}" if h < 24 and mi < 60 else ""

def parse_old_tab(title, vals):
    """Reads one old-layout day sheet (Field | Entry | Notes | Status | Critical) into a Summary Log row."""
    m = OLD_TAB_RE.match(title.strip())
    if not m:
        return None
    kind, dd, mm, yyyy = m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4) or 2026)
    try:
        day = date(yyyy, mm, dd)
    except ValueError:
        return None
    rows = [[str(c).strip() for c in (list(r) + [""] * 6)[:6]] for r in vals]
    if not any(r[0] == "Check / Field" and r[1].startswith("Entry") for r in rows):
        return None                              # not an old-layout sheet (the app's own sheets are skipped)
    body = []
    for r in rows:
        if r[0] == "Issue / Exception" and r[1] == "Action Taken":
            break
        if r[0] and r[0] != "Check / Field" and not re.match(r"^\d+\.", r[0]) and r[1]:
            body.append(r)
    if len(body) < 3:
        return None                              # empty template that was never filled in

    def get(prefix):
        for r in rows:
            if r[0].lower().startswith(prefix):
                return r[1]
        return ""

    time_str = ""
    for r in rows[:8]:
        for i, c in enumerate(r[:-1]):
            if c in ("Time checked", "Time completed"):
                time_str = _time_of(r[i + 1])
    n_exc = 0
    for i, r in enumerate(rows):
        if r[0] == "Issue / Exception" and r[1] == "Action Taken":
            for x in rows[i + 1:]:
                if not any(x) or x[0].startswith("Completed by"):
                    break
                if x[0].lower() in ("", "none", "n/a", "na", "nil", "-", "no") and not x[1] and not x[2]:
                    continue
                n_exc += 1
    out = {k: "" for k in SUM_COLS}
    out.update(date=day.isoformat(), type=kind, time=time_str, exc=n_exc)
    if kind == "Day Start":
        out.update(assembling=_num(get("orders still in assembling")), packing=_num(get("orders still in packing")),
                   moves=_moves(get("cubicles still")), rto=_num(get("rto queue")), dto=_num(get("dto queue")),
                   pest=_pest(get("pest")), ac=_yes_no(get("ac /")), humidifier=_humidifier(get("humidifier")))
    else:
        out.update(assembling=_num(get("orders remaining in assembling")), packing=_num(get("orders remaining in packing")),
                   pickups=_yes_no(get("courier handovers")), moves=_moves(get("cubicles still")),
                   tablets=_yes_no(get("tablets kept")), scanners=_yes_no(get("scanners kept")),
                   pest=_pest(get("final pest")))
    return out

def scan_backfill():
    """Reads every old day sheet in ONE request; returns rows not yet in the Summary Log."""
    sh = get_spreadsheet()
    titles = [w.title for w in sh.worksheets() if OLD_TAB_RE.match(w.title.strip())]
    titles.sort(key=lambda t: len(t.strip()))    # plain names first, so "(14-10) (2)" never beats "(14-10)"
    parsed = []
    if titles:
        ranges = ["'" + t.replace("'", "''") + "'!A1:F90" for t in titles]
        res = sh.values_batch_get(ranges)
        for t, vr in zip(titles, res.get("valueRanges", [])):
            p = parse_old_tab(t, vr.get("values", []))
            if p:
                parsed.append(p)
    have, seen, rows = summary_keys(), set(), []
    for p in sorted(parsed, key=lambda p: (p["date"], 0 if p["type"] == "Day Start" else 1)):
        key = (_to_date(p["date"]), p["type"])
        if key in have or key in seen:
            continue
        seen.add(key)
        rows.append(p)
    return pd.DataFrame([{"Add": True, **{h: str(r[k]) for h, k in zip(SUM_HEADER, SUM_COLS)}} for r in rows],
                        columns=["Add"] + SUM_HEADER)

def import_backfill(df):
    rows = []
    for _, r in df.iterrows():
        if bool(r["Add"]):
            row = {}
            for h, k in zip(SUM_HEADER, SUM_COLS):
                v = _clean_cell(r[h])
                if isinstance(v, str):
                    v = v.strip()
                    if re.fullmatch(r"\d+", v) and k not in ("date", "time"):
                        v = int(v)
                row[k] = v
            rows.append(row)
    if rows:
        upsert_summary(rows)
    return len(rows)

# ─── Save entry points ────────────────────────────────────────────────────────
def log_summary(kind, d, time_str, vals, n_exc):
    """Adds/replaces this entry's Summary Log row. Returns a warning text if it failed (the day sheet is already saved)."""
    try:
        upsert_summary([summary_row(kind, d, time_str, vals, n_exc)])
        return ""
    except Exception as ex:
        return f"The day sheet was saved, but the Summary Log row could not be updated. {explain(ex)}"

def save_day_start(h, vals, exceptions):
    d = h["date"]
    name = tab_name("Day Start", d)
    rows = [
        ["WAREHOUSE — DAY START CHECKLIST"],
        ["Complete after arrival (approx 10:00). Critical items are highlighted. "
         "Closing Supervisor owns Day End checklist."],
        ["Date", d.isoformat(), "Day", d.strftime("%A"), "Time checked", h["time"]],
        ["Completed by", h["by"], "Supervisor", h["supervisor"], "Overall Status", h["status"]],
        BLANK,
    ]
    rows += sections_to_rows(DS_SECTIONS, vals)
    rows += exceptions_to_rows("EXCEPTIONS / ISSUES LOG", exceptions)
    rows.append(["Completed by (signature/name)", h["by"], "", "Supervisor verification"])
    write_tab(name, rows, legacy_name("Day Start", d))
    n = sync_exceptions(d.isoformat(), "Day Start", exceptions)
    return name, n, log_summary("Day Start", d, h["time"], vals, n)

def save_day_end(h, vals, exceptions):
    d = h["date"]
    name = tab_name("Day End", d)
    rows = [
        ["WAREHOUSE — DAY END CHECKLIST"],
        ["TO BE COMPLETED BY CLOSING SUPERVISOR / SENIOR EXEC (target ~7:00 PM). "
         "Critical items highlighted. Review next morning."],
        ["Date", d.isoformat(), "Day", d.strftime("%A"), "Time completed", h["time"]],
        ["Completed by (Closing Supervisor)", h["by"], "Overall Status", h["status"]],
        BLANK,
    ]
    rows += sections_to_rows(DE_SECTIONS, vals)
    rows += exceptions_to_rows("EXCEPTIONS / ISSUES LOG (Hand-over for next day)", exceptions)
    rows.append(["Completed by (Closing Supervisor)", h["by"], "", "Next-day review by", "", h["review"]])
    write_tab(name, rows, legacy_name("Day End", d))
    n = sync_exceptions(d.isoformat(), "Day End", exceptions)
    return name, n, log_summary("Day End", d, h["time"], vals, n)


# ─── Styling ──────────────────────────────────────────────────────────────────
CSS = f"""
<style>
.stApp {{ background:#fff; color:#111; }}
/* Colours are set explicitly (text AND background) so nothing disappears in a dark browser theme */
[data-testid="stMarkdownContainer"] p, [data-testid="stMarkdownContainer"] li,
[data-testid="stMarkdownContainer"] strong, [data-testid="stMarkdownContainer"] em,
[data-testid="stCheckbox"] label p, [data-testid="stCheckbox"] label span {{ color:#111 !important; }}
[data-testid="stWidgetLabel"] p, [data-testid="stCaptionContainer"],
[data-testid="stCaptionContainer"] p {{ color:#222 !important; }}
[data-testid="stAlertContainer"] {{ background:#F3F4F6 !important; border:1px solid #D1D5DB; }}
[data-testid="stAlertContainer"]:has([data-testid="stAlertContentWarning"]) {{ background:#FFF3CD !important; border-color:#F0C24B; }}
[data-testid="stAlertContainer"]:has([data-testid="stAlertContentError"])   {{ background:#FDE2E1 !important; border-color:#F5A3A0; }}
[data-testid="stAlertContainer"]:has([data-testid="stAlertContentSuccess"]) {{ background:#DDF3E4 !important; border-color:#9ADBB0; }}
[data-testid="stAlertContainer"]:has([data-testid="stAlertContentInfo"])    {{ background:#DCEBFA !important; border-color:#A9CCF0; }}
[data-testid="stAlert"] p, [data-testid="stAlert"] li, [data-testid="stAlert"] strong,
[data-testid="stAlert"] a {{ color:#1f2937 !important; }}
header[data-testid="stHeader"] {{ background:transparent; }}
.block-container {{ padding-top:1.2rem; max-width:1150px; }}
div[data-testid="stVerticalBlock"] {{ gap:.45rem; }}

.banner {{ background:{GREEN}; border-radius:4px; padding:18px 10px 16px; text-align:center; margin-bottom:14px; }}
.banner .bt {{ color:{NAVY}; font-size:2.1rem; font-weight:800; line-height:1.1; }}
.banner .bs {{ color:#0b1f4d; font-size:1.05rem; font-weight:600; margin-top:4px; }}
.banner .pill {{ display:inline-block; background:{DARK}; color:#fff; font-size:.8rem; font-weight:700;
                 padding:3px 12px; border-radius:4px; margin-top:10px; }}

.stTabs [role="tablist"] {{ gap:3px; }}
.stTabs [role="tab"] {{ height:auto; padding:10px 22px; border-radius:0; font-weight:700; border:0;
                        background:{AMBER} !important; color:#3b2a00 !important; opacity:1 !important; }}
.stTabs [role="tab"]:nth-child(3) {{ background:{BLUE} !important; color:#fff !important; }}
.stTabs [role="tab"]:nth-child(4) {{ background:#5CCB5F !important; color:#fff !important; }}
.stTabs [role="tab"][aria-selected="true"] {{ background:{ORANGE} !important; color:#fff !important;
                        box-shadow:0 3px 7px rgba(0,0,0,.35); }}
.stTabs [role="tab"] p, .stTabs [role="tab"] div {{ color:inherit !important; font-weight:700; opacity:1 !important; }}
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {{ display:none; }}

div.stButton > button {{ background:{DARK}; color:#fff; border:0; border-radius:4px; font-weight:700; }}
div.stButton > button:hover {{ background:#000; color:#fff; }}
div.stButton > button, div.stButton > button p,
div.stButton > button [data-testid="stMarkdownContainer"] p {{ color:#fff !important; }}

.sec {{ background:{AMBER}; color:#3b2a00; font-weight:800; padding:7px 12px; margin:18px 0 10px;
        border-radius:3px; letter-spacing:.2px; }}
.colh {{ font-size:.7rem; color:#6b7280; font-weight:700; letter-spacing:.6px; }}
.lbl {{ font-weight:500; color:#111; }}
.hl {{ background:#FFFF00; padding:1px 4px; font-weight:600; color:#111; }}
.crit {{ background:{RED}; color:#fff; font-size:.6rem; font-weight:800; padding:1px 6px;
         border-radius:3px; margin-left:6px; vertical-align:middle; }}

.exc {{ border-left:6px solid #999; padding:4px 10px 6px; margin-bottom:10px; }}
.exc .chip {{ color:#fff; font-size:.7rem; font-weight:800; padding:2px 9px; border-radius:3px; margin-right:6px; }}
.exc .meta {{ font-weight:700; color:#111; }}
.exc div.l {{ margin-top:3px; color:#222; }}

.vt-title {{ font-size:1.15rem; font-weight:800; margin:4px 0 6px; color:#111; }}
.vt-sec {{ background:{AMBER}; font-weight:800; padding:5px 10px; margin-top:12px; color:#3b2a00; }}
.vt table {{ width:100%; border-collapse:collapse; font-size:.86rem; }}
.vt td {{ border-bottom:1px solid #eee; padding:3px 6px; vertical-align:top; color:#111; }}
.vt tr.h td {{ background:#EBEBEB; font-weight:700; }}
.vt td.critc {{ color:{RED}; font-weight:800; }}
.vt .note {{ color:#444; font-size:.85rem; margin:3px 0; }}
</style>
"""


# ─── UI helpers ───────────────────────────────────────────────────────────────
def render_sections(prefix, sections):
    """One row per check: label | entry | notes. Returns {key: (value, note)}."""
    out = {}
    for title, items in sections:
        st.markdown(f'<div class="sec">{html.escape(title)}</div>', unsafe_allow_html=True)
        for it in items:
            c1, c2, c3 = st.columns([4, 2, 4], vertical_alignment="center")
            lab = html.escape(it["label"])
            if it["crit"]:
                c1.markdown(f'<span class="hl">{lab}</span><span class="crit">CRITICAL</span>',
                            unsafe_allow_html=True)
            else:
                c1.markdown(f'<span class="lbl">{lab}</span>', unsafe_allow_html=True)
            k = f"{prefix}_{it['key']}"
            with c2:
                if it["kind"] in ("yn", "ny"):          # "ny" = a problem check, so it starts on No
                    v = st.selectbox(it["label"], ["Yes", "No"] if it["kind"] == "yn" else ["No", "Yes"],
                                     key=k, label_visibility="collapsed")
                elif it["kind"] == "num":
                    v = int(st.number_input(it["label"], min_value=0, step=1, key=k,
                                            label_visibility="collapsed"))
                else:
                    v = st.text_input(it["label"], key=k, placeholder=it["ph"],
                                      label_visibility="collapsed")
            n = c3.text_input(it["label"] + " notes", key=k + "_n", placeholder="Notes",
                              label_visibility="collapsed")
            out[it["key"]] = (v, n)
    return out

def exceptions_input(prefix):
    st.markdown('<div class="sec">EXCEPTIONS / ISSUES</div>', unsafe_allow_html=True)
    st.caption("Log the issue, what was done, and who it was escalated to. "
               "Follow-up is handled later in the Exceptions Tracker tab.")
    nkey = f"{prefix}_exc_n"
    st.session_state.setdefault(nkey, 0)
    n = st.session_state[nkey]
    out = []
    for i in range(n):
        c = st.columns([4, 4, 2])
        out.append({
            "issue":     c[0].text_input("Issue", key=f"{prefix}_ei_{i}", placeholder="Issue / exception", label_visibility="collapsed"),
            "action":    c[1].text_input("Action taken", key=f"{prefix}_ea_{i}", placeholder="Action taken", label_visibility="collapsed"),
            "escalated": c[2].text_input("Escalated to", key=f"{prefix}_ee_{i}", placeholder="Escalated to", label_visibility="collapsed"),
        })
    b = st.columns([1.5, 1.5, 5])
    if b[0].button("＋ Add exception", key=f"{prefix}_add_exc"):
        st.session_state[nkey] += 1
        st.rerun()
    if n and b[1].button("− Remove last", key=f"{prefix}_rm_exc"):
        for k in ("ei", "ea", "ee"):
            st.session_state.pop(f"{prefix}_{k}_{n - 1}", None)
        st.session_state[nkey] -= 1
        st.rerun()
    return out

def dup_notice(kind, d):
    """Shown right under the date, so it is seen before the form is filled."""
    names = [tab_name(kind, d)] + ([legacy_name(kind, d)] if legacy_name(kind, d) else [])
    if any(n in existing_tabs() for n in names):
        st.warning(f"⚠️ A **{kind}** entry for **{d:%d/%m/%Y}** already exists. "
                   "Submitting will **replace it completely**, including the exceptions logged from it.")

def submit_gate(prefix, label, name, legacy=None):
    """Renders the submit button; True only when it's OK to write (handles overwrite confirm)."""
    names = [name] + ([legacy] if legacy else [])
    if not st.button(f"✅  Submit {label} Checklist", type="primary",
                     use_container_width=True, key=f"{prefix}_submit"):
        return False
    try:
        exists = any(tab_exists_fresh(n) for n in names)
    except Exception as ex:
        st.error(f"Could not reach the Google Sheet. {explain(ex)}")
        return False
    if exists and st.session_state.get(f"{prefix}_armed") != name:
        st.session_state[f"{prefix}_armed"] = name
        st.error(f"An entry for **{name}** already exists. Click Submit again to overwrite it.")
        return False
    return True

def pretty_date(iso):
    try:
        return datetime.strptime(iso, "%Y-%m-%d").strftime("%d %b %Y")
    except ValueError:
        return iso

def parse_iso(iso):
    try:
        return datetime.strptime(iso, "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None

def tab_html(vals):
    """Read-only render of a saved day sheet."""
    out, table = ['<div class="vt">'], []
    in_table = False

    def flush():
        if table:
            out.append("<table>" + "".join(table) + "</table>")
            table.clear()

    for i, r in enumerate(vals):
        filled = [c for c in r if c.strip()]
        if not filled:
            flush()
            in_table = False
            continue
        first = r[0].strip()
        if i == 0:
            out.append(f'<div class="vt-title">{html.escape(first)}</div>')
        elif len(filled) == 1 and re.match(r"^(\d+\.|EXCEPTIONS)", first):
            flush()
            in_table = False
            out.append(f'<div class="vt-sec">{html.escape(first)}</div>')
        elif len(filled) == 1 and (not in_table or len(first) > 70):
            flush()
            out.append(f'<div class="note">{html.escape(first)}</div>')
        else:
            head = first in ("Check / Field", "Issue / Exception")
            in_table = in_table or head
            r = list(r)
            while r and not r[-1].strip():
                r.pop()
            cells = "".join(
                f'<td class="{"critc" if c == "CRITICAL" else ""}">{html.escape(c)}</td>' for c in r)
            table.append(f'<tr class="{"h" if head else ""}">{cells}</tr>')
    flush()
    out.append("</div>")
    return "".join(out)


# ─── Page ─────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="SayaCare Checklist", layout="wide", page_icon="🏭")
st.markdown(CSS, unsafe_allow_html=True)
st.markdown(
    '<div class="banner"><div class="bt">Warehouse Checklist</div>'
    '<div class="bs">SayaCare Operations</div></div>',
    unsafe_allow_html=True)

# Optional password gate: add  app_password = "..."  to the Streamlit secrets to switch it on.
try:
    _pw = st.secrets["app_password"]
except Exception:
    _pw = None
if _pw and not st.session_state.get("authed"):
    entered = st.text_input("Password", type="password")
    if entered and hmac.compare_digest(entered, str(_pw)):
        st.session_state["authed"] = True
        st.rerun()
    elif entered:
        st.error("Wrong password")
    st.stop()

tab_ds, tab_de, tab_ex, tab_view = st.tabs(
    ["Day Start", "Day End", "Exceptions Tracker", "View Entry"])

# ═══════════════════════════ DAY START ════════════════════════════════════════
with tab_ds:
    c = st.columns(4)
    ds_date   = c[0].date_input("Date", value=now_ist().date(), key="ds_date", format="DD/MM/YYYY")
    ds_by     = c[1].selectbox("Completed by", SUPERVISORS, key="ds_by")
    ds_sup    = c[2].selectbox("Supervisor", SUPERVISORS, key="ds_sup")
    ds_status = c[3].selectbox("Overall status", ["Green", "Amber", "Red"], key="ds_status")

    dup_notice("Day Start", ds_date)
    ds_vals = render_sections("ds", DS_SECTIONS)
    ds_exc  = exceptions_input("ds")
    st.divider()

    ds_name = tab_name("Day Start", ds_date)
    if submit_gate("ds", "Day Start", ds_name, legacy_name("Day Start", ds_date)):
        try:
            header = {"date": ds_date, "time": now_ist().strftime("%H:%M"),
                      "by": ds_by, "supervisor": ds_sup, "status": ds_status}
            name, n, warn = save_day_start(header, ds_vals, ds_exc)
            st.session_state.pop("ds_armed", None)
            clear_caches()
            st.success(f"✅ Saved as **{name}**" + (f" · {n} exception(s) logged" if n else ""))
            if warn:
                st.warning(warn)
        except Exception as ex:
            st.error(f"Error saving. {explain(ex)}")

# ═══════════════════════════ DAY END ══════════════════════════════════════════
with tab_de:
    c = st.columns(4)
    de_date   = c[0].date_input("Date", value=now_ist().date(), key="de_date", format="DD/MM/YYYY")
    de_by     = c[1].selectbox("Completed by (Closing Supervisor)", SUPERVISORS, key="de_by")
    de_review = c[2].selectbox("Next-day review by", ["—"] + SUPERVISORS, key="de_review")
    de_status = c[3].selectbox("Overall status", ["Green", "Amber", "Red"], key="de_status")

    dup_notice("Day End", de_date)
    de_vals = render_sections("de", DE_SECTIONS)
    de_exc  = exceptions_input("de")
    st.divider()

    de_name = tab_name("Day End", de_date)
    if submit_gate("de", "Day End", de_name, legacy_name("Day End", de_date)):
        try:
            header = {"date": de_date, "time": now_ist().strftime("%H:%M"), "by": de_by,
                      "review": "" if de_review == "—" else de_review, "status": de_status}
            name, n, warn = save_day_end(header, de_vals, de_exc)
            st.session_state.pop("de_armed", None)
            clear_caches()
            st.success(f"✅ Saved as **{name}**" + (f" · {n} exception(s) logged" if n else ""))
            if warn:
                st.warning(warn)
        except Exception as ex:
            st.error(f"Error saving. {explain(ex)}")

# ═══════════════════════════ EXCEPTIONS TRACKER ═══════════════════════════════
with tab_ex:
    top = st.columns([1.3, 1, 5])
    show_resolved = top[0].checkbox("Show resolved", key="ex_show_resolved")
    if top[1].button("↻ Refresh", key="ex_refresh"):
        clear_caches()
        st.rerun()

    try:
        allx = load_exceptions()
    except Exception as ex:
        allx = None
        st.error(f"Could not load exceptions. {explain(ex)}")

    if allx is not None:
        today = now_ist().date()
        open_x = [e for e in allx if e["status"] != "Resolved"]
        overdue = [e for e in open_x if (parse_iso(e["fdate"]) or date.max) < today]
        if open_x:
            st.markdown(f"**{len(open_x)} open**" + (f" · :red[{len(overdue)} overdue]" if overdue else ""))
        else:
            st.success("No open exceptions. 🎉")

        for e in open_x:
            fd = parse_iso(e["fdate"])
            is_over = bool(fd and fd < today)
            col = STATUS_COLOUR.get(e["status"], RED)
            with st.container(border=True):
                chips = f'<span class="chip" style="background:{col}">{html.escape(e["status"].upper())}</span>'
                if is_over:
                    chips += f'<span class="chip" style="background:#7f0000">OVERDUE</span>'
                st.markdown(
                    f'<div class="exc" style="border-left-color:{col}">{chips}'
                    f'<span class="meta">{pretty_date(e["date"])} · {html.escape(e["source"])}</span>'
                    f'<div class="l"><b>Issue:</b> {html.escape(e["issue"])}</div>'
                    f'<div class="l"><b>Action taken:</b> {html.escape(e["action"]) or "—"}</div>'
                    f'<div class="l"><b>Escalated to:</b> {html.escape(e["escalated"]) or "—"}</div></div>',
                    unsafe_allow_html=True)
                f = st.columns([2, 2, 2, 1.2], vertical_alignment="bottom")
                fby    = f[0].text_input("Follow-up by", value=e["fby"], key=f"ex_fby_{e['id']}")
                fdate  = f[1].date_input("Follow-up date", value=fd, key=f"ex_fd_{e['id']}", format="DD/MM/YYYY")
                status = f[2].selectbox("Status", STATUSES, index=STATUSES.index(e["status"]),
                                        key=f"ex_st_{e['id']}")
                if f[3].button("Save", key=f"ex_save_{e['id']}", use_container_width=True):
                    try:
                        update_exception(e["id"], fby, fdate.isoformat() if fdate else "", status)
                        clear_caches()
                        st.toast("Resolved — removed from tracker" if status == "Resolved" else "Saved")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"Could not save. {explain(ex)}")

        if show_resolved:
            done = [e for e in allx if e["status"] == "Resolved"]
            st.markdown('<div class="sec">RESOLVED</div>', unsafe_allow_html=True)
            if not done:
                st.caption("Nothing resolved yet.")
            for e in reversed(done):
                st.markdown(
                    f'<div class="exc" style="border-left-color:{OKGREEN}">'
                    f'<span class="chip" style="background:{OKGREEN}">RESOLVED</span>'
                    f'<span class="meta">{pretty_date(e["date"])} · {html.escape(e["source"])}</span>'
                    f'<div class="l"><b>Issue:</b> {html.escape(e["issue"])} — '
                    f'<b>Action:</b> {html.escape(e["action"]) or "—"} — '
                    f'<b>Escalated to:</b> {html.escape(e["escalated"]) or "—"}</div>'
                    f'<div class="l">Follow-up by {html.escape(e["fby"]) or "—"} · '
                    f'resolved {pretty_date(e["resolved_on"])}</div></div>',
                    unsafe_allow_html=True)

# ═══════════════════════════ VIEW ENTRY ═══════════════════════════════════════
with tab_view:
    vd = st.date_input("Date to view", value=now_ist().date(), key="view_date", format="DD/MM/YYYY")
    view_err = False
    try:
        ds_v = load_tab(tab_name("Day Start", vd))
        de_v = load_tab(tab_name("Day End", vd))
        if ds_v is None and legacy_name("Day Start", vd):
            ds_v = load_tab(legacy_name("Day Start", vd))
        if de_v is None and legacy_name("Day End", vd):
            de_v = load_tab(legacy_name("Day End", vd))
    except Exception as ex:
        ds_v = de_v = None
        view_err = True
        st.error(f"Could not load. {explain(ex)}")
    if view_err:
        pass
    elif ds_v is None and de_v is None:
        st.info(f"No Day Start or Day End entry found for {vd:%d/%m/%Y}.")
    else:
        L, R = st.columns(2)
        with L:
            if ds_v is None:
                st.info("No Day Start entry for this date.")
            else:
                st.markdown(tab_html(ds_v), unsafe_allow_html=True)
        with R:
            if de_v is None:
                st.info("No Day End entry for this date.")
            else:
                st.markdown(tab_html(de_v), unsafe_allow_html=True)


# ═══════════════════ ONE-TIME: fill Summary Log from older sheets ═════════════
# Run once, then delete this whole block (from this comment to the end of the file).
with st.expander("🛠️ One-time: fill the Summary Log from existing sheets"):
    st.caption("Reads every older Day Start / Day End sheet, shows what it found, and adds the ticked rows to the "
               "Summary Log. Dates already in the Summary Log are skipped. Fields the old sheets never recorded "
               "stay blank - you can edit any cell below before importing.")
    if st.button("1. Scan existing sheets", key="bf_scan"):
        try:
            st.session_state["bf_df"] = scan_backfill()
            st.session_state.pop("bf_done", None)
        except Exception as ex:
            st.error(f"Could not read the sheets. {explain(ex)}")
    if st.session_state.get("bf_done") is not None:
        st.success(f"✅ Added {st.session_state['bf_done']} row(s) to the Summary Log.")
    df0 = st.session_state.get("bf_df")
    if df0 is not None:
        if df0.empty:
            st.info("Nothing to add - every existing sheet is already in the Summary Log.")
        else:
            st.write(f"Found **{len(df0)}** entries. Untick any row you do not want.")
            edited = st.data_editor(df0, hide_index=True, use_container_width=True, key="bf_editor",
                                    column_config={"Add": st.column_config.CheckboxColumn("Add")})
            if st.button("2. Add ticked rows to the Summary Log", type="primary", key="bf_go"):
                try:
                    st.session_state["bf_done"] = import_backfill(edited)
                    st.session_state.pop("bf_df", None)
                    st.rerun()
                except Exception as ex:
                    st.error(f"Could not write. {explain(ex)}")
