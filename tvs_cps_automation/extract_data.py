"""
extract_data.py
Reads both .xlsb files and returns structured dicts for every slide in the deck.
"""
import pyxlsb
from datetime import datetime, timedelta

# ---------------------------------------------------------------------------
# Column layout of Area wise Summary / ISM Wise / Dealer Summary sheets
# col[0]  = Area / ISM / Dealer identifier cols (may span several)
# FY25-26 block: leads=1, retails=2, conv%=3
# Monthly pairs (Apr..Mar): leads at 4+2*(m-1), retails at 5+2*(m-1) for m=1..12
# FY26-27 block: leads=28, retails=29, conv%=30
# ---------------------------------------------------------------------------

MONTHS = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]

# Excel serial dates for FY25-26 months (Apr'25 .. Mar'26) in the FY25-26 file
FY2526_SERIALS = [45748, 45778, 45809, 45839, 45870, 45901, 45931, 45962, 45992, 46023, 46054, 46082]
# Excel serial dates for FY26-27 months (Apr'26 .. Mar'27) in the FY26-27 file
FY2627_SERIALS = [46113, 46143, 46174, 46204, 46235, 46266, 46296, 46327, 46357, 46388, 46419, 46447]

# Area → Zone mapping
AREA_ZONE = {
    "AP": "South 2", "BIH": "East", "CHT": "West", "CG": "West",
    "DEL": "Delhi/North", "GUJ": "West", "HAR": "North", "HP": "North",
    "JHA": "East", "JH": "East", "JK": "North", "KAR1": "South 2",
    "KAR2": "South 2", "KER": "South", "MAH1": "West", "MAH2": "West",
    "MAH3": "West", "MP1": "West", "MP2": "West", "NE": "East",
    "ODI": "East", "PUN": "North", "RAJ1": "North", "RAJ2": "North",
    "TN1": "South", "TN2": "South", "TN3": "South", "TS": "South 2",
    "UPC": "North", "UPE": "North", "UPH": "North", "WB": "East",
    "GOA": "South 2", "CHD": "North",
}

# Area → State name
AREA_STATE = {
    "AP": "Andhra Pradesh", "BIH": "Bihar", "CHT": "Chhattisgarh",
    "CG": "Chhattisgarh", "DEL": "Delhi", "GUJ": "Gujarat",
    "HAR": "Haryana", "HP": "Himachal Pradesh", "JHA": "Jharkhand",
    "JH": "Jharkhand", "JK": "J&K", "KAR1": "Karnataka",
    "KAR2": "Karnataka", "KER": "Kerala", "MAH1": "Maharashtra",
    "MAH2": "Maharashtra", "MAH3": "Maharashtra", "MP1": "Madhya Pradesh",
    "MP2": "Madhya Pradesh", "NE": "North East", "ODI": "Odisha",
    "PUN": "Punjab", "RAJ1": "Rajasthan", "RAJ2": "Rajasthan",
    "TN1": "Tamil Nadu", "TN2": "Tamil Nadu", "TN3": "Tamil Nadu",
    "TS": "Telangana", "UPC": "Uttar Pradesh", "UPE": "Uttar Pradesh",
    "UPH": "Uttar Pradesh", "WB": "West Bengal", "GOA": "Goa", "CHD": "Chandigarh",
}

# NSM zone groups (areas per zone, as labelled in deck)
NSM_ZONES = {
    "East":    ["WB", "BIH", "ODI", "JHA", "JH", "NE"],
    "North":   ["UPC", "UPE", "UPH", "DEL", "HAR", "PUN", "HP", "JK", "CHD", "RAJ1", "RAJ2"],
    "South":   ["TN1", "TN2", "TN3", "KER"],
    "South 2": ["KAR1", "KAR2", "TS", "AP", "GOA"],
    "West":    ["GUJ", "MAH1", "MAH2", "MAH3", "MP1", "MP2", "CHT", "CG"],
}

# State → list of area codes
STATE_AREAS = {}
for area, state in AREA_STATE.items():
    STATE_AREAS.setdefault(state, []).append(area)


def _month_col(m: int) -> tuple[int, int]:
    """Column indices for month m (1=Apr .. 12=Mar): returns (leads_col, retails_col)."""
    return 4 + (m - 1) * 2, 5 + (m - 1) * 2


def _safe(v):
    """Return float, 0.0 if None/missing."""
    if v is None:
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _read_summary_sheet(filepath: str, sheet_name: str) -> dict:
    """
    Read a summary sheet (Area wise Summary / ISM Wise / Dealer Summary) and
    return:
      {
        "header_key": str,          # e.g. "Area" / "ISM Name" / "Dealer code"
        "rows": {
            key_value: {
                "fy_prev_leads": float, "fy_prev_retails": float, "fy_prev_conv": float,
                "monthly_leads":  [L_apr, L_may, ..., L_mar],  # 12 values, 0 if not yet
                "monthly_retails":[R_apr, ...],
                "fy_curr_leads":  float,
                "fy_curr_retails": float,
                "fy_curr_conv":   float,
            }
        }
      }
    """
    result = {}
    with pyxlsb.open_workbook(filepath) as wb:
        with wb.get_sheet(sheet_name) as sheet:
            rows_raw = []
            for row in sheet.rows():
                rows_raw.append([c.v for c in row])

    # Row 0 = month headers, Row 1 = field names, Row 2+ = data
    if len(rows_raw) < 3:
        return result

    header_col_name = rows_raw[1][0]  # "Area" / "ISM Name" / "Dealer code"

    for row in rows_raw[2:]:
        key = row[0]
        if key is None:
            continue
        key = str(key).strip()
        if not key:
            continue

        monthly_leads = []
        monthly_retails = []
        for m in range(1, 13):
            lc, rc = _month_col(m)
            monthly_leads.append(_safe(row[lc] if lc < len(row) else None))
            monthly_retails.append(_safe(row[rc] if rc < len(row) else None))

        result[key] = {
            "fy_prev_leads":   _safe(row[1] if len(row) > 1 else None),
            "fy_prev_retails": _safe(row[2] if len(row) > 2 else None),
            "fy_prev_conv":    _safe(row[3] if len(row) > 3 else None),
            "monthly_leads":   monthly_leads,
            "monthly_retails": monthly_retails,
            "fy_curr_leads":   _safe(row[28] if len(row) > 28 else None),
            "fy_curr_retails": _safe(row[29] if len(row) > 29 else None),
            "fy_curr_conv":    _safe(row[30] if len(row) > 30 else None),
        }

    return {"header_key": header_col_name, "rows": result}


def _read_dealer_sheet(filepath: str) -> list[dict]:
    """
    Read Dealer Summary sheet and return list of dealer dicts with extra metadata.
    Each dict has all _read_summary_sheet fields plus: name, area, city, tty, state, zone, ism, nsm.
    """
    dealers = []
    with pyxlsb.open_workbook(filepath) as wb:
        with wb.get_sheet("Dealer Summary") as sheet:
            rows_raw = []
            for row in sheet.rows():
                rows_raw.append([c.v for c in row])

    # Row 0 = month headers, Row 1 = col names, Row 2+ = data
    for row in rows_raw[2:]:
        code = row[0]
        if code is None:
            continue
        code_str = str(code).strip()
        if not code_str or code_str.lower() == "grand total":
            continue

        # Metadata columns: 0=code, 1=name, 2=area, 3=city, 4=tty, 5=state, 6=zone, 7=ism, 8=nsm
        # Then: 9=FY prev leads, 10=FY prev retails, 11=FY prev conv%
        # Then monthly pairs from col 12 (Apr)...
        # Note: Dealer Summary has metadata cols 0-8 = 9 cols before data
        # Recalculate col offsets for Dealer Summary (9 metadata cols, not 1)
        META = 9  # number of metadata columns

        def dcol(meta_offset, m):
            # For monthly: leads at META+3+(m-1)*2, retails at META+3+(m-1)*2+1
            leads_c = META + 3 + (m - 1) * 2
            retails_c = leads_c + 1
            return leads_c, retails_c

        monthly_leads = []
        monthly_retails = []
        for m in range(1, 13):
            lc, rc = dcol(META, m)
            monthly_leads.append(_safe(row[lc] if lc < len(row) else None))
            monthly_retails.append(_safe(row[rc] if rc < len(row) else None))

        # FY totals at META+0, META+1, META+2 (prev) and META+27, META+28, META+29 (curr)
        dealers.append({
            "code":     code_str,
            "name":     str(row[1]).strip() if row[1] else "",
            "area":     str(row[2]).strip() if row[2] else "",
            "city":     str(row[3]).strip() if row[3] else "",
            "tty":      str(row[4]).strip() if row[4] else "",
            "state":    str(row[5]).strip() if row[5] else "",
            "zone":     str(row[6]).strip() if row[6] else "",
            "ism":      str(row[7]).strip() if row[7] else "",
            "nsm":      str(row[8]).strip() if row[8] else "",
            "fy_prev_leads":   _safe(row[META]     if len(row) > META     else None),
            "fy_prev_retails": _safe(row[META + 1] if len(row) > META + 1 else None),
            "fy_prev_conv":    _safe(row[META + 2] if len(row) > META + 2 else None),
            "monthly_leads":   monthly_leads,
            "monthly_retails": monthly_retails,
            "fy_curr_leads":   _safe(row[META + 27] if len(row) > META + 27 else None),
            "fy_curr_retails": _safe(row[META + 28] if len(row) > META + 28 else None),
            "fy_curr_conv":    _safe(row[META + 29] if len(row) > META + 29 else None),
        })

    return dealers


def compute_monthly_average(monthly: list[float]) -> float:
    """Average of non-zero monthly values (handles partial years)."""
    vals = [v for v in monthly if v > 0]
    return sum(vals) / len(vals) if vals else 0.0


def compute_ytd(monthly: list[float]) -> float:
    return sum(v for v in monthly if v > 0)


def compute_conv(leads: float, retails: float) -> float:
    return retails / leads if leads > 0 else 0.0


def aggregate_by_zone(area_rows: dict) -> dict[str, dict]:
    """Sum area data by NSM zone."""
    zones: dict[str, dict] = {}
    for zone_name, areas in NSM_ZONES.items():
        leads_prev = 0.0
        retails_prev = 0.0
        monthly_leads = [0.0] * 12
        monthly_retails = [0.0] * 12
        leads_curr = 0.0
        retails_curr = 0.0
        for area in areas:
            if area not in area_rows:
                continue
            d = area_rows[area]
            leads_prev += d["fy_prev_leads"]
            retails_prev += d["fy_prev_retails"]
            for i in range(12):
                monthly_leads[i] += d["monthly_leads"][i]
                monthly_retails[i] += d["monthly_retails"][i]
            leads_curr += d["fy_curr_leads"]
            retails_curr += d["fy_curr_retails"]

        zones[zone_name] = {
            "fy_prev_leads": leads_prev,
            "fy_prev_retails": retails_prev,
            "fy_prev_conv": compute_conv(leads_prev, retails_prev),
            "monthly_leads": monthly_leads,
            "monthly_retails": monthly_retails,
            "fy_curr_leads": leads_curr,
            "fy_curr_retails": retails_curr,
            "fy_curr_conv": compute_conv(leads_curr, retails_curr),
        }
    return zones


def aggregate_by_state(area_rows: dict) -> dict[str, dict]:
    """Sum area data by state name."""
    states: dict[str, dict] = {}
    for state_name, areas in STATE_AREAS.items():
        leads_prev = 0.0
        retails_prev = 0.0
        monthly_leads = [0.0] * 12
        monthly_retails = [0.0] * 12
        leads_curr = 0.0
        retails_curr = 0.0
        for area in areas:
            if area not in area_rows:
                continue
            d = area_rows[area]
            leads_prev += d["fy_prev_leads"]
            retails_prev += d["fy_prev_retails"]
            for i in range(12):
                monthly_leads[i] += d["monthly_leads"][i]
                monthly_retails[i] += d["monthly_retails"][i]
            leads_curr += d["fy_curr_leads"]
            retails_curr += d["fy_curr_retails"]

        states[state_name] = {
            "fy_prev_leads": leads_prev,
            "fy_prev_retails": retails_prev,
            "fy_prev_conv": compute_conv(leads_prev, retails_prev),
            "monthly_leads": monthly_leads,
            "monthly_retails": monthly_retails,
            "fy_curr_leads": leads_curr,
            "fy_curr_retails": retails_curr,
            "fy_curr_conv": compute_conv(leads_curr, retails_curr),
        }
    return states


def get_top_bottom_dealers(dealers: list[dict], state: str, n: int = 10) -> tuple[list, list]:
    """Return top-N and bottom-N dealers for a given state by FY curr conv%."""
    state_dealers = [
        d for d in dealers
        if d["state"] == state and d["fy_curr_leads"] > 50  # min leads threshold
    ]
    sorted_dealers = sorted(state_dealers, key=lambda d: d["fy_curr_conv"], reverse=True)
    top = sorted_dealers[:n]
    bottom = sorted(state_dealers, key=lambda d: d["fy_curr_conv"])[:n]
    return top, bottom


def detect_report_month(curr_xlsb: str) -> int:
    """
    Auto-detect the last month that has actual lead data in the current FY file.
    Returns 1=Apr ... 12=Mar.  Falls back to 1 if nothing found.
    """
    area_rows = _read_summary_sheet(curr_xlsb, "Area wise Summary")["rows"]
    # Aggregate monthly leads across all areas
    total = [0.0] * 12
    for d in area_rows.values():
        for i in range(12):
            total[i] += d["monthly_leads"][i]
    # Last non-zero month
    for m in range(11, -1, -1):
        if total[m] > 0:
            return m + 1  # 1-based
    return 1


def extract_plan(plan_xlsx: str, month_idx: int) -> dict:
    """
    Read a plan Excel file (model-wise monthly plan).
    Expected layout: Row 0 = section header ('Leads'), Row 1 = headers
    ["Model/Category", "Mar'PY", "Apr", "May", ..., "Mar", "Total"],
    Row N = 'Total' row with grand totals.
    Returns {leads_month, retails_month, leads_ytd, retails_ytd}.
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(plan_xlsx, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    except Exception as e:
        print(f"[plan] Could not read plan file: {e}")
        return {"leads_month": 0, "retails_month": 0, "leads_ytd": 0, "retails_ytd": 0}

    # Find header row (has "Apr" or starts with FY year)
    header_row_idx = None
    for i, row in enumerate(rows[:5]):
        if row and any(str(c).startswith("Apr") for c in row if c):
            header_row_idx = i
            break
    if header_row_idx is None:
        print("[plan] Could not find header row with 'Apr'")
        return {"leads_month": 0, "retails_month": 0, "leads_ytd": 0, "retails_ytd": 0}

    # Map month index to column: col 1 = Mar(PY), col 2 = Apr (m=1), ..., col 13 = Mar (m=12)
    # month_idx: 1=Apr, 7=Oct
    def _plan_col(m):
        return 1 + m  # col 2 = Apr, col 3 = May, ... (0-indexed: 1+1=2 for Apr)

    # Find Total row (leads section)
    leads_plan = [0] * 12
    retails_plan = [0] * 12
    section = "leads"
    for row in rows[header_row_idx + 1:]:
        if not row or row[0] is None:
            continue
        label = str(row[0]).strip()
        if label.lower() == "retails" or label.lower() == "retail":
            section = "retails"
            continue
        if label.lower() in ("total", "grand total"):
            for m in range(1, 13):
                col = _plan_col(m)
                val = row[col] if col < len(row) else None
                if isinstance(val, (int, float)):
                    if section == "leads":
                        leads_plan[m - 1] = int(val)
                    else:
                        retails_plan[m - 1] = int(val)

    mi = month_idx - 1
    leads_month = leads_plan[mi]
    retails_month = retails_plan[mi]
    leads_ytd = sum(leads_plan[:mi + 1])
    retails_ytd = sum(retails_plan[:mi + 1])

    print(f"[plan] month={MONTHS[mi]}: leads={leads_month:,}, retails={retails_month:,}, "
          f"ytd_leads={leads_ytd:,}, ytd_retails={retails_ytd:,}")
    return {
        "leads_month": leads_month,
        "retails_month": retails_month,
        "leads_ytd": leads_ytd,
        "retails_ytd": retails_ytd,
    }


def extract_all(
    curr_xlsb: str,
    prev_xlsb: str,
    report_month_idx: int = 0,  # 0 = auto-detect from data
    plan_leads_month: int = 0,
    plan_retails_month: int = 0,
    plan_leads_ytd: int = 0,
    plan_retails_ytd: int = 0,
    # Historical monthly averages (from prior FY files, hardcoded if not provided)
    ma_leads_2324: float = 136700,
    ma_retails_2324: float = 18346,
    ma_conv_2324: float = 0.1341644,
    ma_leads_2425: float = 185600,
    ma_retails_2425: float = 25870,
    ma_conv_2425: float = 0.139369,
) -> dict:
    """
    Master extraction function. Returns complete data dict for all deck slides.
    report_month_idx: 1=Apr, 2=May, 3=Jun, 4=Jul, 5=Aug, 6=Sep, 7=Oct...
    """
    print(f"[extract] Reading current year file: {curr_xlsb}")
    curr_area = _read_summary_sheet(curr_xlsb, "Area wise Summary")["rows"]

    # Auto-detect report month: last month with non-zero total leads
    if report_month_idx == 0:
        total_ml = [0.0] * 12
        for d in curr_area.values():
            for i in range(12):
                total_ml[i] += d["monthly_leads"][i]
        report_month_idx = next((m + 1 for m in range(11, -1, -1) if total_ml[m] > 0), 1)
        print(f"[extract] Auto-detected report month: {MONTHS[report_month_idx-1]} ({report_month_idx})")
    curr_ism  = _read_summary_sheet(curr_xlsb, "ISM Wise")["rows"]
    curr_dealers = _read_dealer_sheet(curr_xlsb)

    print(f"[extract] Reading prior year file:   {prev_xlsb}")
    prev_area = _read_summary_sheet(prev_xlsb, "Area wise Summary")["rows"]
    prev_ism  = _read_summary_sheet(prev_xlsb, "ISM Wise")["rows"]

    # Grand totals
    grand_curr = curr_area.get("Grand Total") or curr_area.get(None) or list(curr_area.values())[0]
    # Fallback: sum all areas
    if "Grand Total" not in curr_area:
        g_leads = [0.0]*12
        g_retails = [0.0]*12
        g_leads_prev = 0.0
        g_retails_prev = 0.0
        for d in curr_area.values():
            g_leads_prev += d["fy_prev_leads"]
            g_retails_prev += d["fy_prev_retails"]
            for i in range(12):
                g_leads[i] += d["monthly_leads"][i]
                g_retails[i] += d["monthly_retails"][i]
        grand_curr = {
            "fy_prev_leads": g_leads_prev, "fy_prev_retails": g_retails_prev,
            "monthly_leads": g_leads, "monthly_retails": g_retails,
            "fy_curr_leads": sum(g_leads), "fy_curr_retails": sum(g_retails),
        }

    # Prev year grand total (for LY same-month comparison)
    if "Grand Total" not in prev_area:
        g_leads_py = [0.0]*12
        g_retails_py = [0.0]*12
        for d in prev_area.values():
            for i in range(12):
                g_leads_py[i] += d["monthly_leads"][i]
                g_retails_py[i] += d["monthly_retails"][i]
        grand_prev = {
            "monthly_leads": g_leads_py, "monthly_retails": g_retails_py,
        }
    else:
        grand_prev = prev_area["Grand Total"]

    mi = report_month_idx - 1  # 0-based month index

    # Computed KPIs
    ty_act_leads   = grand_curr["monthly_leads"][mi]
    ty_act_retails = grand_curr["monthly_retails"][mi]
    ly_act_leads   = grand_prev["monthly_leads"][mi]
    ly_act_retails = grand_prev["monthly_retails"][mi]

    ytd_leads   = sum(grand_curr["monthly_leads"][:mi+1])
    ytd_retails = sum(grand_curr["monthly_retails"][:mi+1])
    ytd_ly_leads   = sum(grand_prev["monthly_leads"][:mi+1])
    ytd_ly_retails = sum(grand_prev["monthly_retails"][:mi+1])

    ma_leads_2526  = compute_monthly_average(grand_prev["monthly_leads"])
    ma_retails_2526= compute_monthly_average(grand_prev["monthly_retails"])
    ma_conv_2526   = compute_conv(ma_leads_2526, ma_retails_2526)
    ytd_act_ma_leads   = ytd_leads   / (mi + 1)
    ytd_act_ma_retails = ytd_retails / (mi + 1)
    ytd_plan_ma_leads  = plan_leads_ytd  / (mi + 1) if plan_leads_ytd else 0
    ytd_plan_ma_retails= plan_retails_ytd/ (mi + 1) if plan_retails_ytd else 0

    # Build prev year monthly arrays (12 values, in thousands for leads)
    prev_monthly_leads   = grand_prev["monthly_leads"]
    prev_monthly_retails = grand_prev["monthly_retails"]

    return {
        # ------- KPI scalars (slides 13-14) -------
        "kpi": {
            "leads": {
                "ly_act": ly_act_leads,
                "ty_plan": plan_leads_month,
                "ty_act": ty_act_leads,
                "pct_ach": ty_act_leads / plan_leads_month if plan_leads_month else 0,
                "growth":  (ty_act_leads - ly_act_leads) / ly_act_leads if ly_act_leads else 0,
                "ytd_ly":    ytd_ly_leads,
                "ytd_plan":  plan_leads_ytd,
                "ytd_act":   ytd_leads,
            },
            "retails": {
                "ly_act": ly_act_retails,
                "ty_plan": plan_retails_month,
                "ty_act": ty_act_retails,
                "pct_ach": ty_act_retails / plan_retails_month if plan_retails_month else 0,
                "growth":  (ty_act_retails - ly_act_retails) / ly_act_retails if ly_act_retails else 0,
                "ytd_ly":    ytd_ly_retails,
                "ytd_plan":  plan_retails_ytd,
                "ytd_act":   ytd_retails,
            },
        },
        # ------- Trend chart series (slides 13-15) -------
        "trend_charts": {
            "leads": {
                # 18 categories: MA23-24, MA24-25, MA25-26, Apr..Mar, LY MTD MA, YTD Plan MA, YTD Act MA
                "bar_trend": {
                    0: ma_leads_2324 / 1000,   # in K
                    1: ma_leads_2425 / 1000,
                    2: ma_leads_2526 / 1000,
                    15: ma_leads_2526 / 1000,  # LY MTD MA
                    16: ytd_plan_ma_leads / 1000,
                    17: ytd_act_ma_leads / 1000,
                },
                "line_curr": {i+3: v/1000 for i, v in enumerate(grand_curr["monthly_leads"]) if v > 0},
                "line_prev": {i+3: v/1000 for i, v in enumerate(prev_monthly_leads) if v > 0},
            },
            "retails": {
                "bar_trend": {
                    0: ma_retails_2324,
                    1: ma_retails_2425,
                    2: ma_retails_2526,
                    15: ma_retails_2526,
                    16: ytd_plan_ma_retails,
                    17: ytd_act_ma_retails,
                },
                "line_curr": {i+3: v for i, v in enumerate(grand_curr["monthly_retails"]) if v > 0},
                "line_prev": {i+3: v for i, v in enumerate(prev_monthly_retails) if v > 0},
            },
            "conv": {
                # Conversion as decimal (0.118 = 11.8%)
                "bar_trend": {
                    0: ma_conv_2324,
                    1: ma_conv_2425,
                    2: ma_conv_2526,
                    15: ma_conv_2526,
                    16: plan_retails_ytd / plan_leads_ytd if plan_leads_ytd else 0,
                    17: ytd_retails / ytd_leads if ytd_leads else 0,
                },
                "line_curr": {
                    i+3: compute_conv(grand_curr["monthly_leads"][i], grand_curr["monthly_retails"][i])
                    for i in range(12)
                    if grand_curr["monthly_leads"][i] > 0
                },
                "line_prev": {
                    i+3: compute_conv(prev_monthly_leads[i], prev_monthly_retails[i])
                    for i in range(12)
                    if prev_monthly_leads[i] > 0
                },
            },
        },
        # ------- Zones (slides 17-23) -------
        "zones": aggregate_by_zone(curr_area),
        "zones_prev": aggregate_by_zone(prev_area),
        # ------- ISMs (slide 25) -------
        "isms": curr_ism,
        "isms_prev": prev_ism,
        # ------- States (slides 27+) -------
        "states":      aggregate_by_state(curr_area),
        "states_prev": aggregate_by_state(prev_area),
        # ------- Dealers (state detail tables) -------
        "dealers": curr_dealers,
        # ------- Meta -------
        "report_month_idx": report_month_idx,
        "report_month_name": MONTHS[mi],
        "ma_leads_2526": ma_leads_2526,
        "ma_retails_2526": ma_retails_2526,
        "ma_conv_2526": ma_conv_2526,
    }
