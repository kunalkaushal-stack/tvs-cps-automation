"""
validate.py
Cross-checks the generated deck's chart values against the source Excel data.
Runs in parallel with deck generation (checks after deck is written).
"""
import re
import os
import json
from chart_updater import _ser_blocks, _numcache_in_ser


def _read_pt(xml_bytes: bytes, ser_idx: int, cat_idx: int) -> float | None:
    """Read the numeric value at series ser_idx, category cat_idx from chart XML."""
    ser_blocks = _ser_blocks(xml_bytes)
    if ser_idx >= len(ser_blocks):
        return None
    ser_bytes = xml_bytes[ser_blocks[ser_idx][0]:ser_blocks[ser_idx][1]]
    nc = _numcache_in_ser(ser_bytes)
    if nc is None:
        return None
    cache = ser_bytes[nc[0]:nc[1]]
    pat = re.compile(rb'<c:pt\s+idx="' + str(cat_idx).encode() + rb'"[^>]*><c:v>([^<]*)</c:v>')
    m = pat.search(cache)
    if m:
        try:
            return float(m.group(1).decode())
        except ValueError:
            return None
    return None


def _rel_err(actual: float, expected: float, tol: float = 0.02) -> bool:
    """True if within tol relative error."""
    if expected == 0:
        return actual == 0
    return abs(actual - expected) / abs(expected) <= tol


CheckResult = dict  # {"check": str, "expected": float, "actual": float, "pass": bool}


def validate_deck(unpacked_dir: str, extracted: dict) -> dict:
    """
    Run validation checks on the unpacked deck directory.
    Returns {"passed": int, "failed": int, "checks": [CheckResult]}
    """
    charts_dir = os.path.join(unpacked_dir, "ppt", "charts")
    checks: list[CheckResult] = []

    mi = extracted["report_month_idx"] - 1  # 0-based month idx

    def check(name: str, chart_num: int, ser_idx: int, cat_idx: int, expected: float):
        chart_path = os.path.join(charts_dir, f"chart{chart_num}.xml")
        if not os.path.exists(chart_path):
            checks.append({"check": name, "expected": expected, "actual": None, "pass": False, "reason": "chart not found"})
            return
        with open(chart_path, "rb") as f:
            xml = f.read()
        actual = _read_pt(xml, ser_idx, cat_idx)
        passed = actual is not None and _rel_err(actual, expected)
        checks.append({
            "check":    name,
            "expected": round(expected, 4),
            "actual":   round(actual, 4) if actual is not None else None,
            "pass":     passed,
        })

    # --- Slide 13 (chart1): Leads ---
    kpi = extracted["kpi"]["leads"]
    tc  = extracted["trend_charts"]["leads"]
    # MA FY25-26 reference bar (ser 0, cat 2) — in K
    check("Leads MA25-26 bar", 1, 0, 2, extracted["ma_leads_2526"] / 1000)
    # TY ACT current month (ser 1, cat = mi+3)
    check(f"Leads TY ACT {extracted['report_month_name']} (ser1)", 1, 1, mi + 3, kpi["ty_act"] / 1000)
    # LY ACT current month (ser 2, cat = mi+3)
    check(f"Leads LY ACT {extracted['report_month_name']} (ser2)", 1, 2, mi + 3, kpi["ly_act"] / 1000)

    # --- Slide 14 (chart2): Retails ---
    kpi_r = extracted["kpi"]["retails"]
    check("Retails MA25-26 bar", 2, 0, 2, extracted["ma_retails_2526"])
    check(f"Retails TY ACT {extracted['report_month_name']} (ser1)", 2, 1, mi + 3, kpi_r["ty_act"])
    check(f"Retails LY ACT {extracted['report_month_name']} (ser2)", 2, 2, mi + 3, kpi_r["ly_act"])

    # --- Slide 15 (chart3): Conv% ---
    tc_c = extracted["trend_charts"]["conv"]
    if mi + 3 in tc_c["line_curr"]:
        check(f"Conv TY ACT {extracted['report_month_name']}", 3, 1, mi + 3, tc_c["line_curr"][mi + 3])

    # --- Slide 25 (chart10): ISM chart for first ISM ---
    isms = extracted["isms"]
    ism_order = [
        "Abhishek Mishra", "Amit Bhakat", "Amit Verma", "Ankit Jaitly",
        "K Sivaramakrishna", "Prasanna Kumar", "Sumit Kale",
        "Swaminathan S", "Vaibhav Maheshwari", "Vivek Kumar Tiwari",
    ]
    for i, ism_name in enumerate(ism_order[:3]):  # check first 3 ISMs
        d = isms.get(ism_name, {})
        expected_leads = d.get("monthly_leads", [0]*12)[mi] / 1000
        if expected_leads > 0:
            check(f"ISM {ism_name} leads (cat {i})", 10, 0, i, expected_leads)

    # --- Zone chart (chart4): East zone conv% ---
    zones = extracted["zones"]
    zones_p = extracted["zones_prev"]
    zone_order = ["East", "North", "South", "South 2", "West"]
    for i, z in enumerate(zone_order[:2]):
        zc = zones.get(z, {})
        l = zc.get("monthly_leads",  [0]*12)[mi]
        r = zc.get("monthly_retails",[0]*12)[mi]
        if l > 0:
            check(f"Zone {z} conv (cat {i})", 4, 2, i, r / l)

    # Summarise
    passed = sum(1 for c in checks if c["pass"])
    failed = len(checks) - passed

    return {
        "passed":  passed,
        "failed":  failed,
        "total":   len(checks),
        "checks":  checks,
    }


def print_report(result: dict) -> None:
    print("\n" + "=" * 60)
    print(f"VALIDATION REPORT  — {result['passed']}/{result['total']} checks passed")
    print("=" * 60)
    for c in result["checks"]:
        icon = "✓" if c["pass"] else "✗"
        reason = c.get("reason", "")
        actual = c["actual"] if c["actual"] is not None else "N/A"
        print(f"  {icon} {c['check']}")
        if not c["pass"]:
            print(f"      Expected: {c['expected']}   Got: {actual}   {reason}")
    print()
    if result["failed"] == 0:
        print("  ALL CHECKS PASSED ✓")
    else:
        print(f"  {result['failed']} FAILURES — review above")
    print("=" * 60)
