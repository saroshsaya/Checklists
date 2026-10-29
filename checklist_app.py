import streamlit as st
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials

# ─── Config ───────────────────────────────────────────────────────────────────
SHEET_ID   = "1jzQpd3MpCp0kT1djnJo_WfQMsbe-reK1oGtvgn6vE1Q"
SUPERVISORS = ["Ashutosh Singh", "Anas", "Sarosh"]
SCOPES      = ["https://www.googleapis.com/auth/spreadsheets"]

# ─── Google Sheets connection ─────────────────────────────────────────────────

@st.cache_resource
def get_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPES,
    )
    return gspread.authorize(creds)

def get_spreadsheet():
    return get_client().open_by_key(SHEET_ID)

def _ws(sh, name):
    """Delete tab with this name if it exists, then create a fresh one."""
    try:
        sh.del_worksheet(sh.worksheet(name))
    except gspread.WorksheetNotFound:
        pass
    return sh.add_worksheet(title=name, rows=120, cols=6)

# ─── Row-builder helpers ──────────────────────────────────────────────────────

def _section_row(title):
    return [title, "", "", "", ""]

def _header_row(col2="Entry"):
    return ["Check / Field", col2, "Notes / Details", "Status", "Critical?"]

def _dr(field, entry="", notes="", status="", critical=""):
    return [field, str(entry) if entry is not None else "", notes, status, critical]

def _exc_header():
    return ["Issue / Exception", "Action Taken", "Escalated To",
            "Follow-up by", "Follow-up Date", "Status"]

def _exc_row(exc):
    return [exc.get("issue", ""), exc.get("action", ""), exc.get("escalated_to", ""),
            exc.get("followup_by", ""), exc.get("followup_date", ""), exc.get("status", "")]

# ─── Sheet writers ────────────────────────────────────────────────────────────

def save_day_start(d):
    sh   = get_spreadsheet()
    name = f"Day Start({d['date'].strftime('%d-%m')})"
    ws   = _ws(sh, name)

    rows = [
        ["WAREHOUSE — DAY START CHECKLIST"],
        ["Complete after arrival (approx 10:00). Critical items are highlighted. "
         "Closing Supervisor owns Day End checklist."],
        ["Date", str(d["date"]), "Day", d["date"].strftime("%A"),
         "Time checked", d["time_checked"]],
        ["Completed by", d["completed_by"], "Supervisor", d["supervisor"],
         "Overall Status", d["overall_status"]],
        [],

        # ── 1. Environment ────────────────────────────────────────────────────
        _section_row("1. ENVIRONMENT & WAREHOUSE READINESS"),
        _header_row(),
    ]
    e = d["env"]
    rows += [
        _dr("AC / climate control working",             e["ac"],             e["ac_notes"]),
        _dr("Temperature (°C)",                         e["temp"]),
        _dr("Humidity (%)",                             e["humidity"]),
        _dr("Humidifier / dehumidifier status",         e["humidifier"],     e["humidifier_notes"]),
        _dr("Pest / insect / rodent activity observed", e["pest"],           e["pest_notes"]),
        _dr("Aisles / walking paths clear",             e["aisles"],         e["aisles_notes"]),
        _dr("Shelves / work areas tidy",                e["shelves"]),
        _dr("Packing tables clean",                     e["packing_tables"], e["packing_tables_notes"]),
        [],

        # ── 2. Stock ──────────────────────────────────────────────────────────
        _section_row("2. STOCK & STORAGE READINESS"),
        _header_row(),
    ]
    s = d["stock"]
    rows += [
        _dr("Stock left in temporary / unusual locations",        s["unusual"],     s["unusual_notes"]),
        _dr("Stock on floor that should be shelved",              s["floor"],       s["floor_notes"]),
        _dr("LT → ST movement still required?",                   s["lt_st"],       "", "", "CRITICAL"),
        _dr("Cubicles still needing movement (e.g. A3, B1, C7)", s["cubicles"],    "", "", "CRITICAL"),
        _dr("Damaged / open stock requiring attention",           s["damaged"],     "", "", "CRITICAL"),
        _dr("Stock discrepancy carried over from previous day",   s["discrepancy"], s["discrepancy_notes"]),
        [],

        # ── 3. Queue ──────────────────────────────────────────────────────────
        _section_row("3. SAMASYA / OPERATIONAL QUEUE (at time of check)"),
        _header_row(),
    ]
    q = d["queue"]
    rows += [
        _dr("Orders still in ASSEMBLING",  q["assembling"], "", "", "CRITICAL"),
        _dr("Orders still in PACKING",     q["packing"],    "", "", "CRITICAL"),
        _dr("RTO queue requiring action",  q["rto"]),
        _dr("DTO queue requiring action",  q["dto"]),
        _dr("Notes pending",               q["notes"]),
        [],

        # ── 4. Equipment ──────────────────────────────────────────────────────
        _section_row("4. EQUIPMENT & PACKING READINESS"),
        _header_row(),
    ]
    eq = d["equip"]
    rows += [
        _dr("Tablets working",                              eq["tablets"],         "", "", "CRITICAL"),
        _dr("Scanner working",                              eq["scanner"],         "", "", "CRITICAL"),
        _dr("Printers working",                             eq["printers"],        "", "", "CRITICAL"),
        _dr("Weighing scales working",                      eq["scales"],          "", "", "CRITICAL"),
        _dr("Thermal rolls available",                      eq["thermal"]),
        _dr("Boxes available",                              eq["boxes"],           "", "", "CRITICAL"),
        _dr("Pouches available",                            eq["pouches"],         "", "", "CRITICAL"),
        _dr("Bubble wrap / protective material available",  eq["bubble_wrap"],     "", "", "CRITICAL"),
        _dr("Tape available",                               eq["tape"]),
        _dr("Packing station ready",                        eq["packing_station"]),
        _dr("Any equipment issue / note",                   eq["eq_issue"],        "", "", "CRITICAL"),
        _dr("Cameras working",                              eq["cameras"],         eq["cameras_notes"]),
        [],

        # ── Exceptions ────────────────────────────────────────────────────────
        ["EXCEPTIONS / ISSUES LOG"],
        _exc_header(),
    ]
    for exc in d["exceptions"]:
        if any(exc.values()):
            rows.append(_exc_row(exc))

    rows.append(["Completed by (signature/name)", d["completed_by"], "", "Supervisor verification"])

    ws.append_rows(rows, value_input_option="USER_ENTERED")
    return name


def save_day_end(d):
    sh   = get_spreadsheet()
    name = f"Day End({d['date'].strftime('%d-%m')})"
    ws   = _ws(sh, name)

    rows = [
        ["WAREHOUSE — DAY END CHECKLIST"],
        ["TO BE COMPLETED BY CLOSING SUPERVISOR / SENIOR EXEC (target ~7:00 PM). "
         "Critical items highlighted. Review next morning."],
        ["Date", str(d["date"]), "Day", d["date"].strftime("%A"),
         "Time completed", d["time_completed"]],
        ["Completed by (Closing Supervisor)", d["completed_by"],
         "Overall Status", d["overall_status"]],
        [],

        # ── 1. Order Closure ──────────────────────────────────────────────────
        _section_row("1. ORDER CLOSURE / WORKLOAD CARRIED FORWARD  ★ CRITICAL"),
        _header_row("Entry / Count"),
    ]
    o = d["orders"]
    rows += [
        _dr("Orders remaining in ASSEMBLING",           o["assembling"],  "", "", "CRITICAL"),
        _dr("Orders remaining in PACKING",              o["packing"],     "", "", "CRITICAL"),
        _dr("Orders remaining in READY FOR SHIPPING",   o["rfs"],         "", "", "CRITICAL"),
        _dr("Issue / Exception",                        o["issue"],       o["issue_notes"], "", "CRITICAL"),
        _dr("Courier handovers completed / recorded",   o["courier"],     o["courier_notes"]),
        [],

        # ── 2. Move List ──────────────────────────────────────────────────────
        _section_row("2. MOVE LIST / CHECKING"),
        _header_row(),
    ]
    m = d["move"]
    rows += [
        _dr("Move List reviewed",                          m["move_list"]),
        _dr("Cubicles still pending LT → ST",             m["cubicles"],      "", "", "CRITICAL"),
        _dr("Checking reviewed",                           m["checking"]),
        _dr("Outstanding discrepancies",                   m["discrepancies"]),
        [],

        # ── 3. RTO / DTO ──────────────────────────────────────────────────────
        _section_row("3. RTO / DTO"),
        _header_row(),
    ]
    rt = d["rto"]
    rows += [
        _dr("RTOs processed today",               rt["rto_processed"]),
        _dr("Unresolved RTO discrepancies",       rt["rto_disc"]),
        _dr("DTOs processed today",               rt["dto_processed"]),
        _dr("Unresolved DTO discrepancies",       rt["dto_disc"]),
        _dr("Returned stock requiring restocking", rt["restocking"]),
        [],

        # ── 4. Warehouse Reset ────────────────────────────────────────────────
        _section_row("4. WAREHOUSE RESET  ★ CRITICAL (Close Clean)"),
        _header_row(),
    ]
    rs = d["reset"]
    rows += [
        _dr("Loose stock returned to appropriate location", rs["loose_stock"],       rs["loose_stock_notes"],  "", "CRITICAL"),
        _dr("LT → ST movement completed",                   rs["lt_st"],             rs["lt_st_notes"],        "", "CRITICAL"),
        _dr("Packing materials replenished",                rs["packing_materials"], "", "", "CRITICAL"),
        _dr("Picking / assembly area reset",                rs["assembly_reset"],    "", "", "CRITICAL"),
        _dr("Packing tables cleared",                       rs["packing_tables"],    "", "", "CRITICAL"),
        _dr("Aisles clear",                                 rs["aisles"],            "", "", "CRITICAL"),
        [],

        # ── 5. Equipment ──────────────────────────────────────────────────────
        _section_row("5. EQUIPMENT"),
        _header_row(),
    ]
    eq = d["equip"]
    rows += [
        _dr("Tablets kept on charge",            eq["tablets"]),
        _dr("Scanners kept on charge / stored",  eq["scanner"]),
        _dr("Printers left ready for next day",  eq["printers"]),
        _dr("Weighing scales left ready",        eq["scales"]),
        _dr("Camera working",                    eq["cameras"]),
        [],

        # ── 6. Environment ────────────────────────────────────────────────────
        _section_row("6. ENVIRONMENT"),
        _header_row(),
    ]
    en = d["env"]
    rows += [
        _dr("Final temperature (°C)",                      en["temp"]),
        _dr("Final humidity (%)",                          en["humidity"]),
        _dr("AC / climate control set for overnight",      en["ac"]),
        _dr("Humidifier / dehumidifier set appropriately", en["humidifier"]),
        _dr("Final pest / insect / rodent observation",    en["pest"]),
        [],

        # ── Exceptions ────────────────────────────────────────────────────────
        ["EXCEPTIONS / ISSUES LOG (Hand-over for next day)"],
        _exc_header(),
    ]
    for exc in d["exceptions"]:
        if any(exc.values()):
            rows.append(_exc_row(exc))

    rows.append(["Completed by (Closing Supervisor)", d["completed_by"],
                 "", "Next-day review by", "", d.get("next_day_review", "")])

    ws.append_rows(rows, value_input_option="USER_ENTERED")
    return name


# ─── Shared UI helpers ────────────────────────────────────────────────────────

def yn(label, key):
    return st.selectbox(label, ["Yes", "No"], key=key)

def exceptions_section(prefix):
    st.markdown("**Exceptions / Issues Log**")
    count_key = f"{prefix}_exc_n"
    if count_key not in st.session_state:
        st.session_state[count_key] = 0

    n = st.session_state[count_key]
    exceptions = []

    if n > 0:
        hcols = st.columns([3, 3, 2, 2, 2, 2])
        for h, col in zip(
            ["Issue / Exception", "Action Taken", "Escalated To",
             "Follow-up by", "Follow-up Date", "Status"],
            hcols,
        ):
            col.markdown(f"<small><b>{h}</b></small>", unsafe_allow_html=True)

    for i in range(n):
        cols = st.columns([3, 3, 2, 2, 2, 2])
        exc = {
            "issue":         cols[0].text_input("Issue",        key=f"{prefix}_ei_{i}", label_visibility="collapsed"),
            "action":        cols[1].text_input("Action",       key=f"{prefix}_ea_{i}", label_visibility="collapsed"),
            "escalated_to":  cols[2].text_input("Escalated To", key=f"{prefix}_ee_{i}", label_visibility="collapsed"),
            "followup_by":   cols[3].text_input("Follow-up by", key=f"{prefix}_ef_{i}", label_visibility="collapsed"),
            "followup_date": cols[4].text_input("Date",         key=f"{prefix}_ed_{i}", label_visibility="collapsed",
                                                placeholder="DD-MM-YYYY"),
            "status":        cols[5].text_input("Status",       key=f"{prefix}_es_{i}", label_visibility="collapsed"),
        }
        exceptions.append(exc)

    if st.button("＋ Add exception", key=f"{prefix}_add_exc"):
        st.session_state[count_key] += 1
        st.rerun()

    return exceptions


# ─── Page layout ──────────────────────────────────────────────────────────────

st.set_page_config(page_title="SayaCare Checklist", layout="wide", page_icon="🏭")
st.title("🏭 SayaCare Warehouse — Daily Checklist")

tab_ds, tab_de = st.tabs(["📋  Day Start", "📦  Day End"])

# ═══════════════════════════ DAY START TAB ════════════════════════════════════
with tab_ds:
    c1, c2, c3, c4 = st.columns(4)
    with c1: ds_date   = st.date_input("Date",          value=date.today(), key="ds_date")
    with c2: ds_by     = st.selectbox("Completed by",   SUPERVISORS,        key="ds_by")
    with c3: ds_sup    = st.selectbox("Supervisor",     SUPERVISORS,        key="ds_sup")
    with c4: ds_status = st.selectbox("Overall Status", ["Green","Amber","Red"], key="ds_status")

    st.divider()

    with st.expander("1. Environment & Warehouse Readiness", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            ds_ac               = yn("AC / climate control working",             "ds_ac")
            ds_ac_notes         = st.text_input("AC notes",          key="ds_ac_notes",      placeholder="Notes / details")
            ds_temp             = st.text_input("Temperature (°C)",  key="ds_temp",          placeholder="e.g. 24")
            ds_humidity         = st.text_input("Humidity (%)",      key="ds_humidity",      placeholder="e.g. 45")
        with c2:
            ds_humidifier       = st.text_input("Humidifier / dehumidifier status", key="ds_humidifier",
                                                placeholder="Working / Stopped / Not Working")
            ds_humidifier_notes = st.text_input("Humidifier notes",  key="ds_humidifier_notes",
                                                placeholder="Notes / details")
            ds_pest             = yn("Pest / insect / rodent activity observed",   "ds_pest")
            ds_pest_notes       = st.text_input("Pest notes",        key="ds_pest_notes",    placeholder="Notes / details")

        c1, c2, c3 = st.columns(3)
        with c1:
            ds_aisles            = yn("Aisles / walking paths clear", "ds_aisles")
            ds_aisles_notes      = st.text_input("Aisles notes",      key="ds_aisles_notes",  placeholder="Notes / details")
        with c2:
            ds_shelves           = yn("Shelves / work areas tidy",    "ds_shelves")
        with c3:
            ds_packing_tables       = yn("Packing tables clean",      "ds_packing_tables")
            ds_packing_tables_notes = st.text_input("Packing tables notes", key="ds_pt_notes",
                                                    placeholder="Notes / details")

    with st.expander("2. Stock & Storage Readiness", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            ds_unusual       = yn("Stock left in temporary / unusual locations", "ds_unusual")
            ds_unusual_notes = st.text_input("Unusual notes",   key="ds_unusual_notes", placeholder="Notes / details")
            ds_floor         = yn("Stock on floor that should be shelved",       "ds_floor")
            ds_floor_notes   = st.text_input("Floor notes",     key="ds_floor_notes",   placeholder="Notes / details")
        with c2:
            ds_lt_st         = yn("★ LT → ST movement still required?",         "ds_lt_st")
            ds_cubicles      = st.text_input("★ Cubicles still needing movement (count or list, e.g. A3 B1)",
                                             key="ds_cubicles")
            ds_damaged       = yn("★ Damaged / open stock requiring attention",  "ds_damaged")
            ds_discrepancy   = st.text_input("Stock discrepancy carried over from previous day",
                                             key="ds_discrepancy", placeholder="None / describe")
            ds_discrepancy_notes = st.text_input("Discrepancy notes", key="ds_disc_notes",
                                                 placeholder="Notes / details")

    with st.expander("3. Samasya / Operational Queue", expanded=True):
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: ds_assembling    = st.number_input("★ Assembling",  min_value=0, step=1, key="ds_assembling")
        with c2: ds_packing_q     = st.number_input("★ Packing",     min_value=0, step=1, key="ds_packing_q")
        with c3: ds_rto           = st.number_input("RTO queue",      min_value=0, step=1, key="ds_rto")
        with c4: ds_dto           = st.number_input("DTO queue",      min_value=0, step=1, key="ds_dto")
        with c5: ds_notes_pending = st.number_input("Notes pending",  min_value=0, step=1, key="ds_notes")

    with st.expander("4. Equipment & Packing Readiness", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            ds_tablets     = yn("★ Tablets working",    "ds_tablets")
            ds_thermal     = yn("Thermal rolls",         "ds_thermal")
            ds_bubble_wrap = yn("★ Bubble wrap",         "ds_bubble_wrap")
        with c2:
            ds_scanner     = yn("★ Scanner working",    "ds_scanner")
            ds_boxes       = yn("★ Boxes available",    "ds_boxes")
            ds_tape        = yn("Tape available",        "ds_tape")
        with c3:
            ds_printers    = yn("★ Printers working",   "ds_printers")
            ds_pouches     = yn("★ Pouches available",  "ds_pouches")
            ds_pst_ready   = yn("Packing station ready","ds_pst_ready")
        with c4:
            ds_scales      = yn("★ Weighing scales working", "ds_scales")

        ds_eq_issue   = st.text_input("★ Any equipment issue / note", key="ds_eq_issue",
                                      placeholder="None / describe")
        c1, c2 = st.columns([1, 2])
        with c1: ds_cameras       = st.text_input("Cameras working", key="ds_cameras",
                                                   placeholder="Yes / describe issue")
        with c2: ds_cameras_notes = st.text_input("Cameras notes",   key="ds_cam_notes",
                                                   placeholder="Notes / details",
                                                   label_visibility="collapsed")

    st.divider()
    ds_exceptions = exceptions_section("ds")

    st.divider()
    if st.button("✅  Submit Day Start Checklist", type="primary",
                 use_container_width=True, key="ds_submit"):
        data = {
            "date":           ds_date,
            "time_checked":   datetime.now().strftime("%H:%M"),
            "completed_by":   ds_by,
            "supervisor":     ds_sup,
            "overall_status": ds_status,
            "env": {
                "ac": ds_ac, "ac_notes": ds_ac_notes,
                "temp": ds_temp, "humidity": ds_humidity,
                "humidifier": ds_humidifier, "humidifier_notes": ds_humidifier_notes,
                "pest": ds_pest, "pest_notes": ds_pest_notes,
                "aisles": ds_aisles, "aisles_notes": ds_aisles_notes,
                "shelves": ds_shelves,
                "packing_tables": ds_packing_tables,
                "packing_tables_notes": ds_packing_tables_notes,
            },
            "stock": {
                "unusual": ds_unusual, "unusual_notes": ds_unusual_notes,
                "floor": ds_floor, "floor_notes": ds_floor_notes,
                "lt_st": ds_lt_st, "cubicles": ds_cubicles,
                "damaged": ds_damaged,
                "discrepancy": ds_discrepancy, "discrepancy_notes": ds_discrepancy_notes,
            },
            "queue": {
                "assembling": int(ds_assembling), "packing": int(ds_packing_q),
                "rto": int(ds_rto), "dto": int(ds_dto), "notes": int(ds_notes_pending),
            },
            "equip": {
                "tablets": ds_tablets, "scanner": ds_scanner,
                "printers": ds_printers, "scales": ds_scales,
                "thermal": ds_thermal, "boxes": ds_boxes,
                "pouches": ds_pouches, "bubble_wrap": ds_bubble_wrap,
                "tape": ds_tape, "packing_station": ds_pst_ready,
                "eq_issue": ds_eq_issue,
                "cameras": ds_cameras, "cameras_notes": ds_cameras_notes,
            },
            "exceptions": ds_exceptions,
        }
        try:
            sname = save_day_start(data)
            st.success(f"✅ Saved as tab **{sname}** in Google Sheets")
        except Exception as ex:
            st.error(f"Error saving: {ex}")


# ═══════════════════════════ DAY END TAB ═════════════════════════════════════
with tab_de:
    c1, c2, c3, c4 = st.columns(4)
    with c1: de_date   = st.date_input("Date",              value=date.today(),      key="de_date")
    with c2: de_by     = st.selectbox("Completed by",       SUPERVISORS,             key="de_by")
    with c3: de_review = st.selectbox("Next-day review by", ["—"] + SUPERVISORS,    key="de_review")
    with c4: de_status = st.selectbox("Overall Status",     ["Green","Amber","Red"], key="de_status")

    st.divider()

    with st.expander("1. Order Closure / Workload Carried Forward  ★ CRITICAL", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1: de_assembling = st.number_input("★ Orders remaining — ASSEMBLING", min_value=0, step=1, key="de_assembling")
        with c2: de_packing_q  = st.number_input("★ Orders remaining — PACKING",    min_value=0, step=1, key="de_packing_q")
        with c3: de_rfs        = st.text_input("Orders remaining — READY FOR SHIPPING",
                                               key="de_rfs", placeholder="None / count / describe")

        c1, c2 = st.columns([1, 2])
        with c1: de_issue       = st.text_input("★ Issue / Exception",  key="de_issue",
                                                placeholder="Order ID or None")
        with c2: de_issue_notes = st.text_input("Issue details",         key="de_issue_notes",
                                                placeholder="Details", label_visibility="collapsed")

        c1, c2 = st.columns([1, 2])
        with c1: de_courier       = st.text_input("Courier handovers completed / recorded",
                                                  key="de_courier", placeholder="Yes / No / details")
        with c2: de_courier_notes = st.text_input("Courier notes", key="de_courier_notes",
                                                  placeholder="Notes", label_visibility="collapsed")

    with st.expander("2. Move List / Checking", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        with c1: de_move_list     = yn("Move List reviewed",   "de_move_list")
        with c2: de_checking      = yn("Checking reviewed",    "de_checking")
        with c3: de_cubicles      = st.text_input("★ Cubicles still pending LT → ST (count or list)",
                                                  key="de_cubicles")
        with c4: de_discrepancies = st.text_input("Outstanding discrepancies", key="de_disc",
                                                  placeholder="None / describe")

    with st.expander("3. RTO / DTO", expanded=True):
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: de_rto_processed = yn("RTOs processed today",               "de_rto_proc")
        with c2: de_rto_disc      = st.text_input("Unresolved RTO disc.",     key="de_rto_disc",
                                                  placeholder="0 / count")
        with c3: de_dto_processed = yn("DTOs processed today",               "de_dto_proc")
        with c4: de_dto_disc      = st.text_input("Unresolved DTO disc.",     key="de_dto_disc",
                                                  placeholder="0 / count")
        with c5: de_restocking    = yn("Returned stock requiring restocking", "de_restocking")

    with st.expander("4. Warehouse Reset  ★ CRITICAL", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            de_loose_stock       = yn("★ Loose stock returned to appropriate location", "de_loose_stock")
            de_loose_stock_notes = st.text_input("Loose stock notes", key="de_ls_notes",
                                                 placeholder="Notes / details")
        with c2:
            de_lt_st             = yn("★ LT → ST movement completed", "de_lt_st")
            de_lt_st_notes       = st.text_input("LT→ST notes",        key="de_ltst_notes",
                                                 placeholder="Notes / details")
        with c3:
            de_packing_materials = yn("★ Packing materials replenished", "de_packing_materials")
            de_assembly_reset    = yn("★ Picking / assembly area reset",  "de_assembly_reset")

        c1, c2 = st.columns(2)
        with c1: de_packing_tables = yn("★ Packing tables cleared", "de_packing_tables")
        with c2: de_aisles         = yn("★ Aisles clear",            "de_aisles")

    with st.expander("5. Equipment", expanded=True):
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: de_tablets  = yn("Tablets on charge",  "de_tablets")
        with c2: de_scanner  = yn("Scanners stored",    "de_scanner")
        with c3: de_printers = yn("Printers ready",     "de_printers")
        with c4: de_scales   = yn("Scales left ready",  "de_scales")
        with c5: de_cameras  = yn("Camera working",     "de_cameras")

    with st.expander("6. Environment", expanded=True):
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: de_temp       = st.text_input("Final temperature (°C)", key="de_temp",
                                               placeholder="e.g. 24")
        with c2: de_humidity   = st.text_input("Final humidity (%)",     key="de_humidity",
                                               placeholder="e.g. 45")
        with c3: de_ac         = yn("AC set for overnight",              "de_ac")
        with c4: de_humidifier = yn("Humidifier set appropriately",      "de_humidifier")
        with c5: de_pest       = st.text_input("Pest observation",       key="de_pest",
                                               placeholder="None / describe")

    st.divider()
    de_exceptions = exceptions_section("de")

    st.divider()
    if st.button("✅  Submit Day End Checklist", type="primary",
                 use_container_width=True, key="de_submit"):
        data = {
            "date":            de_date,
            "time_completed":  datetime.now().strftime("%H:%M"),
            "completed_by":    de_by,
            "next_day_review": de_review if de_review != "—" else "",
            "overall_status":  de_status,
            "orders": {
                "assembling": int(de_assembling), "packing": int(de_packing_q),
                "rfs": de_rfs, "issue": de_issue, "issue_notes": de_issue_notes,
                "courier": de_courier, "courier_notes": de_courier_notes,
            },
            "move": {
                "move_list": de_move_list, "cubicles": de_cubicles,
                "checking": de_checking, "discrepancies": de_discrepancies,
            },
            "rto": {
                "rto_processed": de_rto_processed, "rto_disc": de_rto_disc,
                "dto_processed": de_dto_processed, "dto_disc": de_dto_disc,
                "restocking": de_restocking,
            },
            "reset": {
                "loose_stock": de_loose_stock, "loose_stock_notes": de_loose_stock_notes,
                "lt_st": de_lt_st, "lt_st_notes": de_lt_st_notes,
                "packing_materials": de_packing_materials,
                "assembly_reset": de_assembly_reset,
                "packing_tables": de_packing_tables, "aisles": de_aisles,
            },
            "equip": {
                "tablets": de_tablets, "scanner": de_scanner,
                "printers": de_printers, "scales": de_scales, "cameras": de_cameras,
            },
            "env": {
                "temp": de_temp, "humidity": de_humidity,
                "ac": de_ac, "humidifier": de_humidifier, "pest": de_pest,
            },
            "exceptions": de_exceptions,
        }
        try:
            sname = save_day_end(data)
            st.success(f"✅ Saved as tab **{sname}** in Google Sheets")
        except Exception as ex:
            st.error(f"Error saving: {ex}")
