"""
slide_updater.py
Updates text boxes and table cells in slide XML using regex.
Only replaces text node contents — never re-serialises XML.
"""
import re
import os


def _read(path: str) -> bytes:
    with open(path, "rb") as f:
        return f.read()


def _write(path: str, data: bytes) -> None:
    with open(path, "wb") as f:
        f.write(data)


# ---------------------------------------------------------------------------
# Text-box helpers
# ---------------------------------------------------------------------------

def _get_text_runs(xml_bytes: bytes) -> list[tuple[int, int, str]]:
    """
    Return list of (start, end, text) for every <a:t>...</a:t> in the XML.
    """
    runs = []
    for m in re.finditer(rb"<a:t>([^<]*)</a:t>", xml_bytes):
        runs.append((m.start(), m.end(), m.group(1).decode("utf-8", errors="replace")))
    return runs


def replace_text_exact(xml_bytes: bytes, old_text: str, new_text: str) -> bytes:
    """Replace first occurrence of old_text in an <a:t> run with new_text."""
    old_b = ("<a:t>" + old_text + "</a:t>").encode()
    new_b = ("<a:t>" + new_text + "</a:t>").encode()
    return xml_bytes.replace(old_b, new_b, 1)


def replace_all_text(xml_bytes: bytes, old_text: str, new_text: str) -> bytes:
    """Replace all occurrences of old_text in <a:t> runs."""
    old_b = ("<a:t>" + old_text + "</a:t>").encode()
    new_b = ("<a:t>" + new_text + "</a:t>").encode()
    return xml_bytes.replace(old_b, new_b)


# ---------------------------------------------------------------------------
# Table helpers
# ---------------------------------------------------------------------------

def _table_cells(xml_bytes: bytes) -> list[tuple[int, int, int, int]]:
    """
    Returns list of (row_idx, col_idx, start, end) for each <a:tc> in the first table.
    """
    cells = []
    tbl_start = xml_bytes.find(b"<a:tbl>")
    tbl_end   = xml_bytes.find(b"</a:tbl>", tbl_start)
    if tbl_start < 0:
        return cells
    tbl = xml_bytes[tbl_start:tbl_end]

    r = 0
    for tr_m in re.finditer(rb"<a:tr\b[^>]*>", tbl):
        tr_start = tr_m.start()
        tr_end   = tbl.find(b"</a:tr>", tr_start)
        if tr_end < 0:
            continue
        row_bytes = tbl[tr_start:tr_end]
        c = 0
        for tc_m in re.finditer(rb"<a:tc>", row_bytes):
            tc_start = tc_m.start()
            tc_end   = row_bytes.find(b"</a:tc>", tc_start)
            if tc_end < 0:
                continue
            abs_start = tbl_start + tr_start + tc_start
            abs_end   = tbl_start + tr_start + tc_end + len(b"</a:tc>")
            cells.append((r, c, abs_start, abs_end))
            c += 1
        r += 1

    return cells


def set_table_cell(xml_bytes: bytes, row: int, col: int, text: str, table_idx: int = 0) -> bytes:
    """
    Set the text content of table cell [row][col] (0-based).
    table_idx allows targeting a second table if the slide has two.
    """
    # Find the right table
    tbl_starts = [m.start() for m in re.finditer(rb"<a:tbl>", xml_bytes)]
    if table_idx >= len(tbl_starts):
        return xml_bytes

    tbl_start = tbl_starts[table_idx]
    tbl_end   = xml_bytes.find(b"</a:tbl>", tbl_start)
    tbl_bytes = xml_bytes[tbl_start:tbl_end]

    # Navigate to [row][col]
    rows = list(re.finditer(rb"<a:tr\b[^>]*>", tbl_bytes))
    if row >= len(rows):
        return xml_bytes

    tr_start = rows[row].start()
    tr_end   = tbl_bytes.find(b"</a:tr>", tr_start)
    row_bytes = tbl_bytes[tr_start:tr_end]

    cols = list(re.finditer(rb"<a:tc>", row_bytes))
    if col >= len(cols):
        return xml_bytes

    tc_start = cols[col].start()
    tc_end   = row_bytes.find(b"</a:tc>", tc_start) + len(b"</a:tc>")
    tc_bytes = row_bytes[tc_start:tc_end]

    # Replace first <a:t>...</a:t> inside the cell
    t_match = re.search(rb"<a:t>[^<]*</a:t>", tc_bytes)
    if t_match:
        new_tc = tc_bytes[:t_match.start()] + b"<a:t>" + text.encode() + b"</a:t>" + tc_bytes[t_match.end():]
    else:
        # No text node found, can't update
        return xml_bytes

    abs_tr_start = tbl_start + tr_start
    abs_tc_start = abs_tr_start + tc_start

    return (
        xml_bytes[:abs_tc_start]
        + new_tc
        + xml_bytes[abs_tc_start + tc_end:]
    )


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def fmt_num(n: float) -> str:
    """Format as integer with comma thousands separator."""
    return f"{int(round(n)):,}"


def fmt_pct(n: float) -> str:
    """Format as percentage string like '11.8%'."""
    return f"{n*100:.1f}%"


def fmt_K(n: float) -> str:
    """Format in thousands like '157.6K'."""
    return f"{n/1000:.1f}K"


# ---------------------------------------------------------------------------
# Slide 1: Title month
# ---------------------------------------------------------------------------

def update_title_slide(slide_path: str, old_month: str, new_month: str) -> None:
    """Replace month label on title slide (e.g. 'Aug-2026' → 'Sep-2026')."""
    xml = _read(slide_path)
    xml = replace_all_text(xml, old_month, new_month)
    _write(slide_path, xml)


# ---------------------------------------------------------------------------
# Slides 13-14: KPI callout text boxes
# ---------------------------------------------------------------------------

def update_kpi_slide(slide_path: str, kpi: dict, old_values: dict) -> None:
    """
    Replace KPI text values on slides 13 (Leads) or 14 (Retails).
    old_values = {field: "old string in slide"} — read from current deck
    kpi        = {field: float}
    """
    xml = _read(slide_path)

    mapping = {
        "ty_act":   fmt_num(kpi["ty_act"]),
        "ly_act":   fmt_num(kpi["ly_act"]),
        "ty_plan":  fmt_num(kpi["ty_plan"]) if kpi["ty_plan"] else "",
        "pct_ach":  fmt_pct(kpi["pct_ach"]),
        "growth":   (("+" if kpi["growth"] >= 0 else "") + fmt_pct(kpi["growth"])),
        "ytd_ly":   fmt_num(kpi["ytd_ly"]),
        "ytd_plan": fmt_num(kpi["ytd_plan"]) if kpi["ytd_plan"] else "",
        "ytd_act":  fmt_num(kpi["ytd_act"]),
    }

    for field, old_str in old_values.items():
        new_str = mapping.get(field, "")
        if new_str and old_str:
            xml = replace_all_text(xml, old_str, new_str)

    _write(slide_path, xml)


# ---------------------------------------------------------------------------
# State detail table slides (even slides: 28, 30, 32...)
# ---------------------------------------------------------------------------

def update_state_table_slide(
    slide_path: str,
    state_name: str,
    state_code: str,
    top_dealers: list[dict],
    bottom_dealers: list[dict],
    state_summary: dict,     # {month_name: (leads, retails)} + fy totals
    report_month_idx: int,
) -> None:
    """
    Update state detail tables on slides like 28, 30, 32...
    Each slide has:
      - 1 summary row table (state totals for last 4 months + FY totals)
      - 2 dealer tables (top 10 + bottom 10)
    """
    from extract_data import MONTHS
    xml = _read(slide_path)
    mi = report_month_idx - 1  # 0-based

    # --- Summary table (table_idx=0) ---
    # Row structure: header + data row
    # Columns: State | M-3 Leads | M-3 Retails | M-2 Leads | M-2 Retails | M-1 Leads | M-1 Retails | M Leads | M Retails | FY25-26 Leads | FY25-26 Retails | FY25-26 Conv% | FY26-27 Leads | FY26-27 Retails | FY26-27 Conv%
    if state_summary:
        curr_ml = state_summary.get("monthly_leads", [0]*12)
        curr_mr = state_summary.get("monthly_retails", [0]*12)
        # Data row (row=1, col=0 is state code)
        try:
            xml = set_table_cell(xml, 1, 0, state_code, table_idx=0)
            for offset, col_pair in enumerate(range(0, 8, 2)):
                m_idx = max(0, mi - 3 + offset)
                xml = set_table_cell(xml, 1, 1 + col_pair, fmt_num(curr_ml[m_idx]), table_idx=0)
                xml = set_table_cell(xml, 1, 2 + col_pair, fmt_num(curr_mr[m_idx]), table_idx=0)
            xml = set_table_cell(xml, 1, 9,  fmt_num(state_summary.get("fy_prev_leads", 0)), table_idx=0)
            xml = set_table_cell(xml, 1, 10, fmt_num(state_summary.get("fy_prev_retails", 0)), table_idx=0)
            xml = set_table_cell(xml, 1, 11, fmt_pct(state_summary.get("fy_prev_conv", 0)), table_idx=0)
            xml = set_table_cell(xml, 1, 12, fmt_num(state_summary.get("fy_curr_leads", 0)), table_idx=0)
            xml = set_table_cell(xml, 1, 13, fmt_num(state_summary.get("fy_curr_retails", 0)), table_idx=0)
            xml = set_table_cell(xml, 1, 14, fmt_pct(state_summary.get("fy_curr_conv", 0)), table_idx=0)
        except Exception:
            pass  # table structure may differ — skip gracefully

    # --- Top dealers table (table_idx=1) ---
    for i, dealer in enumerate(top_dealers[:10]):
        try:
            row = i + 1  # row 0 = header
            xml = set_table_cell(xml, row, 0, fmt_pct(dealer["fy_prev_conv"]), table_idx=1)
            xml = set_table_cell(xml, row, 1, str(dealer["code"]),             table_idx=1)
            xml = set_table_cell(xml, row, 2, dealer["name"][:30],             table_idx=1)
            xml = set_table_cell(xml, row, 3, dealer["city"],                  table_idx=1)
            xml = set_table_cell(xml, row, 4, fmt_num(dealer["fy_curr_leads"]),   table_idx=1)
            xml = set_table_cell(xml, row, 5, fmt_num(dealer["fy_curr_retails"]), table_idx=1)
            xml = set_table_cell(xml, row, 6, fmt_pct(dealer["fy_curr_conv"]),    table_idx=1)
        except Exception:
            pass

    # --- Bottom dealers table (table_idx=2) ---
    for i, dealer in enumerate(bottom_dealers[:10]):
        try:
            row = i + 1
            xml = set_table_cell(xml, row, 0, fmt_pct(dealer["fy_prev_conv"]), table_idx=2)
            xml = set_table_cell(xml, row, 1, str(dealer["code"]),             table_idx=2)
            xml = set_table_cell(xml, row, 2, dealer["name"][:30],             table_idx=2)
            xml = set_table_cell(xml, row, 3, dealer["city"],                  table_idx=2)
            xml = set_table_cell(xml, row, 4, fmt_num(dealer["fy_curr_leads"]),   table_idx=2)
            xml = set_table_cell(xml, row, 5, fmt_num(dealer["fy_curr_retails"]), table_idx=2)
            xml = set_table_cell(xml, row, 6, fmt_pct(dealer["fy_curr_conv"]),    table_idx=2)
        except Exception:
            pass

    _write(slide_path, xml)
