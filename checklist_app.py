import hmac
import html
import re
import uuid
from datetime import datetime, date
from zoneinfo import ZoneInfo

import gspread
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


def now_ist():
    return datetime.now(IST)


# ─── Checklist definitions (single source of truth for UI + sheet output) ─────
def I(label, key, kind="yn", crit=False, ph=""):
    return {"label": label, "key": key, "kind": kind, "crit": crit, "ph": ph}

DS_SECTIONS = [
    ("1. ENVIRONMENT & WAREHOUSE READINESS", [
        I("AC / climate control working",             "ac"),
        I("Temperature (°C)",                         "temp",       "text", ph="e.g. 24"),
        I("Humidity (%)",                             "humidity",   "text", ph="e.g. 45"),
        I("Humidifier / dehumidifier status",         "humidifier", "text", ph="Working / Stopped"),
        I("Pest / insect / rodent activity observed", "pest"),
        I("Aisles / walking paths clear",             "aisles"),
        I("Shelves / work areas tidy",                "shelves"),
        I("Packing tables clean",                     "packing_tables"),
    ]),
    ("2. STOCK & STORAGE READINESS", [
        I("Stock left in temporary / unusual locations",        "unusual"),
        I("Stock on floor that should be shelved",              "floor"),
        I("LT → ST movement still required?",                   "lt_st", crit=True),
        I("Cubicles still needing movement (e.g. A3, B1, C7)",  "cubicles", "text", True, "A3, B1 / count"),
        I("Damaged / open stock requiring attention",           "damaged", crit=True),
        I("Stock discrepancy carried over from previous day",   "discrepancy", "text", ph="None / describe"),
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
        I("Orders remaining in READY FOR SHIPPING", "rfs",        "text", True, "None / count"),
        I("Issue / Exception (order ID)",           "issue",      "text", True, "Order ID or None"),
        I("Courier handovers completed / recorded", "courier"),
    ]),
    ("2. MOVE LIST / CHECKING", [
        I("Move List reviewed",                      "move_list"),
        I("Cubicles still pending LT → ST",          "cubicles", "text", True, "A3, B1 / count"),
        I("Checking reviewed",                       "checking"),
        I("Outstanding discrepancies",               "discrepancies", "text", ph="None / describe"),
    ]),
    ("3. RTO / DTO", [
        I("RTOs processed today",                "rto_processed"),
        I("Unresolved RTO discrepancies",        "rto_disc", "text", ph="0 / count"),
        I("DTOs processed today",                "dto_processed"),
        I("Unresolved DTO discrepancies",        "dto_disc", "text", ph="0 / count"),
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
        I("Final temperature (°C)",                      "temp",       "text", ph="e.g. 24"),
        I("Final humidity (%)",                          "humidity",   "text", ph="e.g. 45"),
        I("AC / climate control set for overnight",      "ac"),
        I("Humidifier / dehumidifier set appropriately", "humidifier"),
        I("Final pest / insect / rodent observation",    "pest", "text", ph="None / describe"),
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

def clear_caches():
    existing_tabs.clear()
    load_tab.clear()
    load_exceptions.clear()


# ─── Building sheet rows ──────────────────────────────────────────────────────
BLANK = [""]

def sections_to_rows(sections, vals):
    rows = []
    for title, items in sections:
        rows.append([title, "", "", "", ""])
        rows.append(["Check / Field", "Entry", "Notes / Details", "Status", "Critical?"])
        for it in items:
            v, note = vals[it["key"]]
            rows.append([it["label"], str(v), note, "", "CRITICAL" if it["crit"] else ""])
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
        elif len(r) > 4 and r[4] == "CRITICAL":
            cell_fmt(i, 4, 5, {"textFormat": {"bold": True, "foregroundColor": {"red": 0.9, "green": 0.1, "blue": 0.1}}},
                     "userEnteredFormat.textFormat")
    for col, px in enumerate([380, 150, 300, 110, 90, 150]):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": col, "endIndex": col + 1},
            "properties": {"pixelSize": px}, "fields": "pixelSize"}})
    sh.batch_update({"requests": reqs})

def write_tab(name, rows):
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


# ─── Save entry points ────────────────────────────────────────────────────────
def save_day_start(h, vals, exceptions):
    d = h["date"]
    name = f"Day Start({d:%d-%m})"
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
    write_tab(name, rows)
    n = sync_exceptions(d.isoformat(), "Day Start", exceptions)
    return name, n

def save_day_end(h, vals, exceptions):
    d = h["date"]
    name = f"Day End({d:%d-%m})"
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
    write_tab(name, rows)
    n = sync_exceptions(d.isoformat(), "Day End", exceptions)
    return name, n


# ─── Styling ──────────────────────────────────────────────────────────────────
CSS = f"""
<style>
.stApp {{ background:#fff; }}
header[data-testid="stHeader"] {{ background:transparent; }}
.block-container {{ padding-top:1.2rem; max-width:1150px; }}
div[data-testid="stVerticalBlock"] {{ gap:.45rem; }}

.banner {{ background:{GREEN}; border-radius:4px; padding:18px 10px 16px; text-align:center; margin-bottom:14px; }}
.banner .bt {{ color:{NAVY}; font-size:2.1rem; font-weight:800; line-height:1.1; }}
.banner .bs {{ color:#0b1f4d; font-size:1.05rem; font-weight:600; margin-top:4px; }}
.banner .pill {{ display:inline-block; background:{DARK}; color:#fff; font-size:.8rem; font-weight:700;
                 padding:3px 12px; border-radius:4px; margin-top:10px; }}

.stTabs [data-baseweb="tab-list"] {{ gap:3px; }}
.stTabs [data-baseweb="tab"] {{ height:auto; padding:10px 22px; border-radius:0; font-weight:700;
                                background:{AMBER}; color:#3b2a00; }}
.stTabs [data-baseweb="tab"]:nth-child(3) {{ background:{BLUE}; color:#fff; }}
.stTabs [data-baseweb="tab"]:nth-child(4) {{ background:#5CCB5F; color:#fff; }}
.stTabs [data-baseweb="tab"][aria-selected="true"] {{ background:{ORANGE}; color:#fff;
                                box-shadow:0 3px 7px rgba(0,0,0,.35); }}
.stTabs [data-baseweb="tab"] p {{ color:inherit; font-weight:700; }}
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {{ display:none; }}

div.stButton > button {{ background:{DARK}; color:#fff; border:0; border-radius:4px; font-weight:700; }}
div.stButton > button:hover {{ background:#000; color:#fff; }}
div.stButton > button p {{ color:#fff; }}

.sec {{ background:{AMBER}; color:#3b2a00; font-weight:800; padding:7px 12px; margin:18px 0 4px;
        border-radius:3px; letter-spacing:.2px; }}
.colh {{ font-size:.7rem; color:#6b7280; font-weight:700; letter-spacing:.6px; }}
.lbl {{ font-weight:500; color:#111; }}
.hl {{ background:#FFFF00; padding:1px 4px; font-weight:600; color:#111; }}
.crit {{ background:{RED}; color:#fff; font-size:.6rem; font-weight:800; padding:1px 6px;
         border-radius:3px; margin-left:6px; vertical-align:middle; }}

.exc {{ border-left:6px solid #999; padding:4px 10px 6px; }}
.exc .chip {{ color:#fff; font-size:.7rem; font-weight:800; padding:2px 9px; border-radius:3px; margin-right:6px; }}
.exc .meta {{ font-weight:700; color:#111; }}
.exc div.l {{ margin-top:3px; color:#222; }}

.vt-title {{ font-size:1.15rem; font-weight:800; margin:4px 0 6px; }}
.vt-sec {{ background:{AMBER}; font-weight:800; padding:5px 10px; margin-top:12px; color:#3b2a00; }}
.vt table {{ width:100%; border-collapse:collapse; font-size:.86rem; }}
.vt td {{ border-bottom:1px solid #eee; padding:3px 6px; vertical-align:top; color:#111; }}
.vt tr.h td {{ background:#EBEBEB; font-weight:700; }}
.vt td.crit {{ color:{RED}; font-weight:800; }}
.vt .note {{ color:#444; font-size:.85rem; margin:3px 0; }}
</style>
"""


# ─── UI helpers ───────────────────────────────────────────────────────────────
def render_sections(prefix, sections):
    """One row per check: label | entry | notes. Returns {key: (value, note)}."""
    out = {}
    for title, items in sections:
        st.markdown(f'<div class="sec">{html.escape(title)}</div>', unsafe_allow_html=True)
        h = st.columns([4, 2, 4])
        for c, t in zip(h, ["CHECK", "ENTRY", "NOTES (OPTIONAL)"]):
            c.markdown(f'<div class="colh">{t}</div>', unsafe_allow_html=True)
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
                if it["kind"] == "yn":
                    v = st.selectbox(it["label"], ["Yes", "No"], key=k, label_visibility="collapsed")
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
    if n:
        h = st.columns([4, 4, 2])
        for c, t in zip(h, ["ISSUE / EXCEPTION", "ACTION TAKEN", "ESCALATED TO"]):
            c.markdown(f'<div class="colh">{t}</div>', unsafe_allow_html=True)
    for i in range(n):
        c = st.columns([4, 4, 2])
        out.append({
            "issue":     c[0].text_input("Issue",        key=f"{prefix}_ei_{i}", label_visibility="collapsed"),
            "action":    c[1].text_input("Action taken", key=f"{prefix}_ea_{i}", label_visibility="collapsed"),
            "escalated": c[2].text_input("Escalated to", key=f"{prefix}_ee_{i}", label_visibility="collapsed"),
        })
    b = st.columns([1.3, 1.6, 6])
    if b[0].button("＋ Add exception", key=f"{prefix}_add_exc"):
        st.session_state[nkey] += 1
        st.rerun()
    if n and b[1].button("− Remove last", key=f"{prefix}_rm_exc"):
        for k in ("ei", "ea", "ee"):
            st.session_state.pop(f"{prefix}_{k}_{n - 1}", None)
        st.session_state[nkey] -= 1
        st.rerun()
    return out

def submit_gate(prefix, label, name):
    """Renders the submit button; True only when it's OK to write (handles overwrite confirm)."""
    if name in existing_tabs():
        st.warning(f"⚠️ **{name}** already exists. Submitting will **replace it completely**, "
                   "including the exceptions logged from it.")
    if not st.button(f"✅  Submit {label} Checklist", type="primary",
                     use_container_width=True, key=f"{prefix}_submit"):
        return False
    if tab_exists_fresh(name) and st.session_state.get(f"{prefix}_armed") != name:
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

    def flush():
        if table:
            out.append("<table>" + "".join(table) + "</table>")
            table.clear()

    for i, r in enumerate(vals):
        filled = [c for c in r if c.strip()]
        if not filled:
            flush()
            continue
        first = r[0].strip()
        if i == 0:
            out.append(f'<div class="vt-title">{html.escape(first)}</div>')
        elif len(filled) == 1 and re.match(r"^(\d+\.|EXCEPTIONS)", first):
            flush()
            out.append(f'<div class="vt-sec">{html.escape(first)}</div>')
        elif len(filled) == 1:
            flush()
            out.append(f'<div class="note">{html.escape(first)}</div>')
        else:
            head = first in ("Check / Field", "Issue / Exception")
            cells = "".join(
                f'<td class="{"crit" if c == "CRITICAL" else ""}">{html.escape(c)}</td>' for c in r)
            table.append(f'<tr class="{"h" if head else ""}">{cells}</tr>')
    flush()
    out.append("</div>")
    return "".join(out)


# ─── Page ─────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="SayaCare Checklist", layout="wide", page_icon="🏭")
st.markdown(CSS, unsafe_allow_html=True)
st.markdown(
    '<div class="banner"><div class="bt">Warehouse Checklist</div>'
    '<div class="bs">SayaCare Operations</div><span class="pill">Location: Noida</span></div>',
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

    ds_vals = render_sections("ds", DS_SECTIONS)
    ds_exc  = exceptions_input("ds")
    st.divider()

    ds_name = f"Day Start({ds_date:%d-%m})"
    if submit_gate("ds", "Day Start", ds_name):
        try:
            header = {"date": ds_date, "time": now_ist().strftime("%H:%M"),
                      "by": ds_by, "supervisor": ds_sup, "status": ds_status}
            name, n = save_day_start(header, ds_vals, ds_exc)
            st.session_state.pop("ds_armed", None)
            clear_caches()
            st.success(f"✅ Saved as **{name}**" + (f" · {n} exception(s) logged" if n else ""))
        except Exception as ex:
            st.error(f"Error saving: {ex}")

# ═══════════════════════════ DAY END ══════════════════════════════════════════
with tab_de:
    c = st.columns(4)
    de_date   = c[0].date_input("Date", value=now_ist().date(), key="de_date", format="DD/MM/YYYY")
    de_by     = c[1].selectbox("Completed by (Closing Supervisor)", SUPERVISORS, key="de_by")
    de_review = c[2].selectbox("Next-day review by", ["—"] + SUPERVISORS, key="de_review")
    de_status = c[3].selectbox("Overall status", ["Green", "Amber", "Red"], key="de_status")

    de_vals = render_sections("de", DE_SECTIONS)
    de_exc  = exceptions_input("de")
    st.divider()

    de_name = f"Day End({de_date:%d-%m})"
    if submit_gate("de", "Day End", de_name):
        try:
            header = {"date": de_date, "time": now_ist().strftime("%H:%M"), "by": de_by,
                      "review": "" if de_review == "—" else de_review, "status": de_status}
            name, n = save_day_end(header, de_vals, de_exc)
            st.session_state.pop("de_armed", None)
            clear_caches()
            st.success(f"✅ Saved as **{name}**" + (f" · {n} exception(s) logged" if n else ""))
        except Exception as ex:
            st.error(f"Error saving: {ex}")

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
        st.error(f"Could not load exceptions: {ex}")

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
                        st.error(f"Could not save: {ex}")

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
    try:
        ds_v = load_tab(f"Day Start({vd:%d-%m})")
        de_v = load_tab(f"Day End({vd:%d-%m})")
    except Exception as ex:
        ds_v = de_v = None
        st.error(f"Could not load: {ex}")
    if ds_v is None and de_v is None:
        st.info(f"No Day Start or Day End entry found for {vd:%d-%m}.")
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
