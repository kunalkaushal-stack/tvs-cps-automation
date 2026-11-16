"""
run_pipeline.py
Main orchestrator for TVS CPS Performance Deck automation.

Usage:
  py run_pipeline.py \\
    --curr   "path/to/FY26-27 Oct'26.xlsb" \\
    --prev   "path/to/FY25-26 Mar'26.xlsb" \\
    --template "path/to/TVS CPS Performance Deck_Aug'26.pptx" \\
    --month  5                           # 1=Apr, 5=Aug, 6=Sep, 7=Oct ...
    --month-label "Sep-2026"             # label for title slide
    --old-month-label "Aug-2026"         # label currently on template
    --plan-leads-month    180000
    --plan-retails-month  20500
    --plan-leads-ytd      890000
    --plan-retails-ytd    100000
    --out    "TVS_CPS_Sep26.pptx"

The script:
  1. Copies the template .pptx as the new deck (preserves all formatting)
  2. Unpacks the copy
  3. In parallel: extracts data from Excel + updates deck charts/tables/text
  4. Validates the deck numbers against source Excel
  5. Repacks the new deck
"""
import argparse
import os
import sys
import shutil
import zipfile
import tempfile
import threading
import time

# Add parent dir to path so imports work when run from anywhere
sys.path.insert(0, os.path.dirname(__file__))
import extract_data as ed
import chart_updater as cu
import slide_updater as su
import validate as val


# ---------------------------------------------------------------------------
# Slide → chart mapping (from map_charts.py analysis)
# ---------------------------------------------------------------------------
SLIDE_CHART_MAP = {
    13: 1,   # Leads trend
    14: 2,   # Retails trend
    15: 3,   # Conv% trend
    17: 4,   # Zone overview bar
    19: 5,   # NSM East
    20: 6,   # NSM South
    21: 7,   # NSM North
    22: 8,   # NSM West
    23: 9,   # NSM South 2
    25: 10,  # ISM comparison
    # State trend charts (odd slides starting 27)
    27: 11,  # Maharashtra
    29: 12,  # Rajasthan
    31: 13,  # Madhya Pradesh
    33: 14,  # Bihar
    35: 15,  # Karnataka & Goa
    37: 16,  # Gujarat
    39: 17,
    41: 18,
    43: 19,
    45: 20,
    47: 21,
    49: 22,
    51: 23,
    53: 24,
    55: 25,
    57: 26,
    59: 27,
    61: 28,
    63: 29,
}

# State detail table slides (even, follow each trend slide)
STATE_TABLE_SLIDES = list(range(28, 70, 2))  # 28, 30, 32, ... 68

# State order in deck (must match slide order above)
STATE_ORDER = [
    "Maharashtra",
    "Rajasthan",
    "Madhya Pradesh",
    "Bihar",
    "Karnataka",
    "Gujarat",
    "Uttar Pradesh",
    "Haryana",
    "Delhi",
    "Punjab",
    "Tamil Nadu",
    "Telangana",
    "Andhra Pradesh",
    "West Bengal",
    "Odisha",
    "North East",
    "Jharkhand",
    "Kerala",
    "Chhattisgarh",
]

NSM_ZONE_ORDER = ["East", "South", "North", "West", "South 2"]  # slides 19-23


# ---------------------------------------------------------------------------
# Step 1: Copy template → working copy
# ---------------------------------------------------------------------------
def copy_template(template_path: str, output_path: str) -> None:
    print(f"[deck] Copying template → {output_path}")
    shutil.copy2(template_path, output_path)


# ---------------------------------------------------------------------------
# Step 2: Unpack the deck to a temp directory
# ---------------------------------------------------------------------------
def unpack(pptx_path: str, unpack_dir: str) -> None:
    if os.path.exists(unpack_dir):
        shutil.rmtree(unpack_dir)
    os.makedirs(unpack_dir)
    with zipfile.ZipFile(pptx_path, "r") as z:
        z.extractall(unpack_dir)
    print(f"[deck] Unpacked to {unpack_dir}")


# ---------------------------------------------------------------------------
# Step 3a: Update all charts
# ---------------------------------------------------------------------------
def update_charts(unpacked: str, extracted: dict, args) -> None:
    charts_dir = os.path.join(unpacked, "ppt", "charts")
    mi = extracted["report_month_idx"] - 1

    # --- Slide 13: Leads trend ---
    chart_path = os.path.join(charts_dir, "chart1.xml")
    if os.path.exists(chart_path):
        upd = cu.build_trend_updates(extracted["trend_charts"]["leads"])
        cu.update_chart_file(chart_path, upd)
        print("[deck] chart1 (Leads) updated")

    # --- Slide 14: Retails trend ---
    chart_path = os.path.join(charts_dir, "chart2.xml")
    if os.path.exists(chart_path):
        upd = cu.build_trend_updates(extracted["trend_charts"]["retails"])
        cu.update_chart_file(chart_path, upd)
        print("[deck] chart2 (Retails) updated")

    # --- Slide 15: Conv% trend ---
    chart_path = os.path.join(charts_dir, "chart3.xml")
    if os.path.exists(chart_path):
        upd = cu.build_trend_updates(extracted["trend_charts"]["conv"])
        cu.update_chart_file(chart_path, upd)
        print("[deck] chart3 (Conv%) updated")

    # --- Slide 17: Zone overview bar ---
    chart_path = os.path.join(charts_dir, "chart4.xml")
    if os.path.exists(chart_path):
        upd = cu.build_zone_bar_updates(extracted["zones"], extracted["zones_prev"], extracted["report_month_idx"])
        cu.update_chart_file(chart_path, upd)
        print("[deck] chart4 (Zone bar) updated")

    # --- Slides 19-23: NSM zone trend charts ---
    for slide_num, zone_name in zip([19, 20, 21, 22, 23], NSM_ZONE_ORDER):
        chart_num = SLIDE_CHART_MAP.get(slide_num)
        if not chart_num:
            continue
        chart_path = os.path.join(charts_dir, f"chart{chart_num}.xml")
        if not os.path.exists(chart_path):
            continue
        zc = extracted["zones"].get(zone_name, {})
        zp = extracted["zones_prev"].get(zone_name, {})
        upd = cu.build_nsm_trend_updates(zc, zp)
        cu.update_chart_file(chart_path, upd)
        print(f"[deck] chart{chart_num} (NSM {zone_name}) updated")

    # --- Slide 25: ISM bar ---
    chart_path = os.path.join(charts_dir, "chart10.xml")
    if os.path.exists(chart_path):
        upd = cu.build_ism_bar_updates(extracted["isms"], extracted["isms_prev"], extracted["report_month_idx"])
        cu.update_chart_file(chart_path, upd)
        print("[deck] chart10 (ISM bar) updated")

    # --- State trend charts ---
    for i, state_name in enumerate(STATE_ORDER):
        slide_num = 27 + i * 2
        chart_num = SLIDE_CHART_MAP.get(slide_num)
        if not chart_num:
            continue
        chart_path = os.path.join(charts_dir, f"chart{chart_num}.xml")
        if not os.path.exists(chart_path):
            continue
        sc = extracted["states"].get(state_name, {})
        sp = extracted["states_prev"].get(state_name, {})
        if not sc:
            continue
        upd = cu.build_state_trend_updates(sc, sp)
        cu.update_chart_file(chart_path, upd)
        print(f"[deck] chart{chart_num} ({state_name}) updated")


# ---------------------------------------------------------------------------
# Step 3b: Update text boxes and tables in slides
# ---------------------------------------------------------------------------
def update_slides(unpacked: str, extracted: dict, args) -> None:
    slides_dir = os.path.join(unpacked, "ppt", "slides")

    # --- Slide 1: Title month ---
    slide1 = os.path.join(slides_dir, "slide1.xml")
    if os.path.exists(slide1) and args.old_month_label and args.month_label:
        su.update_title_slide(slide1, args.old_month_label, args.month_label)
        print(f"[deck] slide1 title updated: {args.old_month_label} → {args.month_label}")

    # --- Slides 13-14: KPI callouts ---
    # We use the known existing values from the Aug'26 template to find and replace
    kpi_leads   = extracted["kpi"]["leads"]
    kpi_retails = extracted["kpi"]["retails"]

    slide13 = os.path.join(slides_dir, "slide13.xml")
    if os.path.exists(slide13):
        old_vals = {
            "ty_act":  "157,618",
            "ly_act":  "166,636",
            "ty_plan": "180,000",
            "pct_ach": "88%",
            "ytd_act": "719,860",
            "ytd_ly":  "910,651",
            "ytd_plan": "890,000",
        }
        su.update_kpi_slide(slide13, kpi_leads, old_vals)
        print("[deck] slide13 KPIs updated")

    slide14 = os.path.join(slides_dir, "slide14.xml")
    if os.path.exists(slide14):
        old_vals_r = {
            "ty_act":  "18,664",
            "ly_act":  "20,386",
            "ty_plan": "20,500",
            "pct_ach": "91%",
            "ytd_act": "93,964",
            "ytd_ly":  "132,426",
            "ytd_plan": "100,000",
        }
        su.update_kpi_slide(slide14, kpi_retails, old_vals_r)
        print("[deck] slide14 KPIs updated")

    # --- State table slides ---
    from extract_data import get_top_bottom_dealers, MONTHS
    mi = extracted["report_month_idx"] - 1

    for i, state_name in enumerate(STATE_ORDER):
        table_slide_num = 28 + i * 2
        slide_path = os.path.join(slides_dir, f"slide{table_slide_num}.xml")
        if not os.path.exists(slide_path):
            continue

        # Guess state code from first area of this state
        from extract_data import STATE_AREAS, AREA_STATE
        state_code = next(
            (a for a, s in AREA_STATE.items() if s == state_name),
            state_name[:3].upper()
        )

        top, bottom = get_top_bottom_dealers(extracted["dealers"], state_name)
        sc = extracted["states"].get(state_name, {})

        su.update_state_table_slide(
            slide_path, state_name, state_code,
            top, bottom, sc, extracted["report_month_idx"]
        )
        print(f"[deck] slide{table_slide_num} ({state_name} table) updated")


# ---------------------------------------------------------------------------
# Step 4: Repack the deck
# ---------------------------------------------------------------------------
def repack(unpacked_dir: str, output_path: str) -> None:
    if os.path.exists(output_path):
        os.remove(output_path)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(unpacked_dir):
            for file in files:
                abs_path = os.path.join(root, file)
                arc_path = os.path.relpath(abs_path, unpacked_dir)
                z.write(abs_path, arc_path)
    print(f"[deck] Repacked → {output_path}")


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------
def run(args) -> None:
    t0 = time.time()
    print("\n" + "=" * 60)
    print("TVS CPS DECK AUTOMATION PIPELINE")
    print("=" * 60)

    # --- Stage 1: Extract data ---
    print("\n[1/4] Extracting data from Excel files...")
    extracted = ed.extract_all(
        curr_xlsb=args.curr,
        prev_xlsb=args.prev,
        report_month_idx=args.month,
        plan_leads_month=args.plan_leads_month,
        plan_retails_month=args.plan_retails_month,
        plan_leads_ytd=args.plan_leads_ytd,
        plan_retails_ytd=args.plan_retails_ytd,
    )
    print(f"  Report month: {extracted['report_month_name']}")
    print(f"  TY Leads this month: {extracted['kpi']['leads']['ty_act']:,.0f}")
    print(f"  TY Retails this month: {extracted['kpi']['retails']['ty_act']:,.0f}")

    # --- Stage 2: Copy template and unpack ---
    print("\n[2/4] Preparing deck copy...")
    output_pptx = args.out
    copy_template(args.template, output_pptx)
    unpack_dir = output_pptx + "_unpacked"
    unpack(output_pptx, unpack_dir)

    # --- Stage 3: Update deck (charts + slides) in parallel with validation prep ---
    print("\n[3/4] Updating deck content (charts + slides)...")

    validation_result = {}

    def run_deck_update():
        update_charts(unpack_dir, extracted, args)
        update_slides(unpack_dir, extracted, args)

    def run_validation_after_update():
        # Wait briefly for the update to start, then validate what's done
        time.sleep(0.5)
        result = val.validate_deck(unpack_dir, extracted)
        validation_result.update(result)

    # Deck update runs in main logic; validation starts in a background thread
    # (validation checks the chart XML files as they get written)
    val_thread = threading.Thread(target=run_validation_after_update, daemon=True)
    val_thread.start()

    run_deck_update()

    # Repack
    repack(unpack_dir, output_pptx)

    # Wait for validation to complete
    val_thread.join(timeout=30)

    # --- Stage 4: Report ---
    print("\n[4/4] Validation results:")
    if validation_result:
        val.print_report(validation_result)
    else:
        print("  (validation timed out)")

    elapsed = time.time() - t0
    print(f"\n✓ Done in {elapsed:.1f}s  →  {output_pptx}\n")


def main():
    p = argparse.ArgumentParser(description="TVS CPS Deck Automation")
    p.add_argument("--curr",     required=True, help="Current FY .xlsb file path")
    p.add_argument("--prev",     required=True, help="Prior FY .xlsb file path")
    p.add_argument("--template", required=True, help="Template .pptx (previous month's deck)")
    p.add_argument("--month",    required=True, type=int, help="Report month index: 1=Apr, 5=Aug, 6=Sep...")
    p.add_argument("--month-label",     default="",  help="New month label for title slide e.g. 'Sep-2026'")
    p.add_argument("--old-month-label", default="",  help="Existing month label to replace e.g. 'Aug-2026'")
    p.add_argument("--plan-leads-month",    type=int, default=0)
    p.add_argument("--plan-retails-month",  type=int, default=0)
    p.add_argument("--plan-leads-ytd",      type=int, default=0)
    p.add_argument("--plan-retails-ytd",    type=int, default=0)
    p.add_argument("--out", default="TVS_CPS_new.pptx", help="Output .pptx path")
    args = p.parse_args()
    run(args)


if __name__ == "__main__":
    main()
