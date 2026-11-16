"""
chart_updater.py
Updates chart XML numeric caches via regex.
Avoids re-serialising the XML (which rewrites namespace prefixes and corrupts the deck).
"""
import re
import os


# Categories axis (18 positions):
# 0=MA 23-24, 1=MA 24-25, 2=MA 25-26, 3=Apr, 4=May, ..., 14=Mar, 15=LY MTD MA,
# 16=YTD Plan MA, 17=YTD Act MA
MONTHS = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]


def _ser_blocks(xml_bytes: bytes) -> list[tuple[int, int]]:
    """
    Return list of (start, end) byte offsets for each <c:ser>...</c:ser> block.
    """
    blocks = []
    for m in re.finditer(rb"<c:ser>", xml_bytes):
        start = m.start()
        # Find matching </c:ser> — they don't nest
        end_m = xml_bytes.find(b"</c:ser>", start)
        if end_m >= 0:
            blocks.append((start, end_m + len(b"</c:ser>")))
    return blocks


def _numcache_in_ser(ser_bytes: bytes) -> tuple[int, int] | None:
    """
    Return (start, end) of the <c:numCache>...</c:numCache> inside a ser block.
    """
    start = ser_bytes.find(b"<c:numCache>")
    if start < 0:
        return None
    end = ser_bytes.find(b"</c:numCache>", start)
    if end < 0:
        return None
    return start, end + len(b"</c:numCache>")


def _set_pt(cache_bytes: bytes, idx: int, value) -> bytes:
    """
    Set the value at <c:pt idx="N"><c:v>...</c:v> in a numCache block.
    If the point doesn't exist, append it before </c:numCache>.
    """
    formatted = f"{value:.10g}" if isinstance(value, float) else str(value)
    formatted_b = formatted.encode()

    # Pattern: <c:pt idx="N" ...><c:v>...</c:v>
    # May have formatCode attr: <c:pt idx="N" formatCode="...">
    pat = re.compile(
        rb'(<c:pt\s+idx="' + str(idx).encode() + rb'"[^>]*>)<c:v>[^<]*</c:v>',
        re.DOTALL,
    )
    m = pat.search(cache_bytes)
    if m:
        return cache_bytes[:m.start()] + m.group(1) + b"<c:v>" + formatted_b + b"</c:v>" + cache_bytes[m.end():]

    # Point doesn't exist — insert before </c:numCache>
    end_tag = b"</c:numCache>"
    pos = cache_bytes.rfind(end_tag)
    if pos < 0:
        return cache_bytes
    new_pt = b'<c:pt idx="' + str(idx).encode() + b'"><c:v>' + formatted_b + b"</c:v></c:pt>"
    return cache_bytes[:pos] + new_pt + cache_bytes[pos:]


def _clear_pt(cache_bytes: bytes, idx: int) -> bytes:
    """Remove a <c:pt idx="N">...</c:pt> from a numCache block (makes it blank/gap)."""
    pat = re.compile(
        rb'<c:pt\s+idx="' + str(idx).encode() + rb'"[^>]*>[^<]*</c:pt>',
        re.DOTALL,
    )
    return pat.sub(b"", cache_bytes)


def update_series(xml_bytes: bytes, ser_idx: int, updates: dict) -> bytes:
    """
    Update points in series ser_idx.
    updates = {chart_cat_idx: value_or_None}  (None = remove/gap the point)
    ser_idx: 0-based series index (bar series = 0, line FY26-27 = 1, line FY25-26 = 2)
    """
    ser_blocks = _ser_blocks(xml_bytes)
    if ser_idx >= len(ser_blocks):
        return xml_bytes

    s_start, s_end = ser_blocks[ser_idx]
    ser_bytes = xml_bytes[s_start:s_end]

    nc = _numcache_in_ser(ser_bytes)
    if nc is None:
        return xml_bytes

    nc_start, nc_end = nc
    cache = ser_bytes[nc_start:nc_end]

    for idx, val in updates.items():
        if val is None:
            cache = _clear_pt(cache, idx)
        else:
            cache = _set_pt(cache, idx, val)

    new_ser = ser_bytes[:nc_start] + cache + ser_bytes[nc_end:]
    return xml_bytes[:s_start] + new_ser + xml_bytes[s_end:]


def update_chart_file(chart_path: str, series_updates: dict[int, dict]) -> None:
    """
    Update a chart XML file.
    series_updates = {ser_idx: {cat_idx: value}}
    """
    with open(chart_path, "rb") as f:
        xml_bytes = f.read()

    for ser_idx, updates in series_updates.items():
        xml_bytes = update_series(xml_bytes, ser_idx, updates)

    with open(chart_path, "wb") as f:
        f.write(xml_bytes)


# ---------------------------------------------------------------------------
# High-level builders: build series_updates dict for each chart type
# ---------------------------------------------------------------------------

def build_trend_updates(data: dict) -> dict:
    """
    Build series_updates for a standard trend chart (leads, retails, or conv).
    data: {"bar_trend": {idx: v}, "line_curr": {idx: v}, "line_prev": {idx: v}}
    Returns {0: bar_updates, 1: curr_updates, 2: prev_updates}
    """
    return {
        0: data["bar_trend"],
        1: data["line_curr"],
        2: data["line_prev"],
    }


def build_zone_bar_updates(zones: dict, zones_prev: dict, month_idx: int) -> dict:
    """
    Build series_updates for zone overview bar chart (slide 17).
    4 series: Leads, Retails, Conv%_curr, Conv%_prev — one bar per zone.
    Zone order: East, North, South, South 2, West (must match chart categories).
    """
    zone_order = ["East", "North", "South", "South 2", "West"]
    mi = month_idx - 1  # 0-based

    leads_curr    = {}
    retails_curr  = {}
    conv_curr     = {}
    conv_prev     = {}

    for i, z in enumerate(zone_order):
        zc = zones.get(z, {})
        zp = zones_prev.get(z, {})
        l = zc.get("monthly_leads", [0]*12)[mi]
        r = zc.get("monthly_retails", [0]*12)[mi]
        leads_curr[i]   = l / 1000  # in K
        retails_curr[i] = r
        conv_curr[i]    = r / l if l else 0
        lp = zp.get("monthly_leads",  [0]*12)[mi]
        rp = zp.get("monthly_retails",[0]*12)[mi]
        conv_prev[i]    = rp / lp if lp else 0

    return {
        0: leads_curr,
        1: retails_curr,
        2: conv_curr,
        3: conv_prev,
    }


def build_nsm_trend_updates(zone_data: dict, zone_prev: dict) -> dict:
    """
    Build series_updates for a single NSM zone trend chart (slides 19-23).
    Same 3-series structure as global trend charts.
    """
    curr_leads   = zone_data.get("monthly_leads", [0]*12)
    curr_retails = zone_data.get("monthly_retails", [0]*12)
    prev_leads   = zone_prev.get("monthly_leads", [0]*12)
    prev_retails = zone_prev.get("monthly_retails", [0]*12)

    # Category index mapping: 0=MA 23-24, 1=MA 24-25, 2=MA 25-26, 3=Apr..14=Mar
    # MA values come from the data averages stored in the extraction

    # Build bar series (reference lines only - no monthly bars needed beyond MA placeholders)
    # For NSM charts the bar shows leads, line shows retails + conv
    bar_updates = {}
    curr_line = {}
    prev_line = {}

    for i in range(12):
        if curr_leads[i] > 0:
            bar_updates[i + 3] = curr_leads[i] / 1000  # leads in K as bars
        if curr_retails[i] > 0:
            curr_line[i + 3] = curr_retails[i]
        if prev_retails[i] > 0:
            prev_line[i + 3] = prev_retails[i]

    # MA values (bar series 0-2) — use averages
    ma_curr = sum(v for v in curr_leads if v > 0) / max(1, sum(1 for v in curr_leads if v > 0)) / 1000
    ma_prev = sum(v for v in prev_leads if v > 0) / max(1, sum(1 for v in prev_leads if v > 0)) / 1000
    bar_updates[2] = ma_prev  # MA 25-26 reference

    # Conv lines (series 3 and 4 if present)
    conv_curr = {}
    conv_prev2 = {}
    for i in range(12):
        if curr_leads[i] > 0:
            conv_curr[i + 3]  = curr_retails[i] / curr_leads[i]
        if prev_leads[i] > 0:
            conv_prev2[i + 3] = prev_retails[i] / prev_leads[i]

    return {
        0: bar_updates,   # leads bars
        1: curr_line,     # retails line FY26-27
        2: prev_line,     # retails line FY25-26
        3: conv_curr,     # conv% FY26-27
        4: conv_prev2,    # conv% FY25-26
    }


def build_ism_bar_updates(isms: dict, isms_prev: dict, month_idx: int) -> dict:
    """
    Build series_updates for ISM comparison bar chart (slide 25).
    ISM order must match chart categories (from slide_text analysis).
    """
    ism_order = [
        "Abhishek Mishra", "Amit Bhakat", "Amit Verma", "Ankit Jaitly",
        "K Sivaramakrishna", "Prasanna Kumar", "Sumit Kale",
        "Swaminathan S", "Vaibhav Maheshwari", "Vivek Kumar Tiwari",
    ]
    mi = month_idx - 1

    leads_curr   = {}
    retails_curr = {}
    conv_curr    = {}
    conv_prev    = {}

    for i, name in enumerate(ism_order):
        d  = isms.get(name, {})
        dp = isms_prev.get(name, {})
        l  = d.get("monthly_leads",  [0]*12)[mi]
        r  = d.get("monthly_retails",[0]*12)[mi]
        lp = dp.get("monthly_leads", [0]*12)[mi]
        rp = dp.get("monthly_retails",[0]*12)[mi]
        leads_curr[i]   = l / 1000
        retails_curr[i] = r
        conv_curr[i]    = r / l if l else 0
        conv_prev[i]    = rp / lp if lp else 0

    return {0: leads_curr, 1: retails_curr, 2: conv_curr, 3: conv_prev}


def build_state_trend_updates(state_data: dict, state_prev: dict) -> dict:
    """Build trend chart updates for a single state (identical to NSM structure)."""
    return build_nsm_trend_updates(state_data, state_prev)
