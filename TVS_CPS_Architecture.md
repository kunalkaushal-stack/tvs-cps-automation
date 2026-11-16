# TVS CPS Reporting System — Complete Architecture & Logic

> Generated: 2026-11-16  
> Source files: Dealer and Modelwise Performance Report (FY25-26 Mar'26 & FY26-27 Oct'26) + TVS CPS Performance Deck Aug'26

---

## 1. OVERVIEW

This system produces the **TVS CPS (Customer Performance Summary) Monthly Performance Deck** — a PowerPoint presentation shared with the sales leadership summarising BikeDekho/Zigwheels lead-to-retail conversion performance across the TVS dealer network, broken out by Zone, NSM, ISM, State, Area, and Dealer.

Three artefacts exist:

| Artefact | File Type | Role |
|---|---|---|
| **Dealer Modelwise Performance Report FY25-26** | `.xlsb` | Historical baseline (full FY25-26, Apr'25–Mar'26) |
| **Dealer Modelwise Performance Report FY26-27** | `.xlsb` | Current year live data (FY26-27, Apr'26–running month) |
| **TVS CPS Performance Deck** | `.pptx` | Monthly output deck populated from both Excel files |

---

## 2. EXCEL FILE ARCHITECTURE

### 2.1 Sheet Structure (identical in both FY files)

Both `.xlsb` files share **exactly the same 14 sheets** in the same order:

| # | Sheet Name | Granularity | Purpose |
|---|---|---|---|
| 1 | `Dealer wise Model wise` | Dealer × Model × Month | Most granular — leads/retails/conv per dealer per model per month |
| 2 | `Dealer Summary` | Dealer × Month | Total 2W aggregated per dealer per month (no model split) |
| 3 | `Area wise Summary` | Area × Month | Area-level rollup of leads/retails/conv per month |
| 4 | `ISM Wise` | ISM × Month | ISM-level rollup per month |
| 5 | `Sumit Kale` | Dealer × Month (ISM-filtered) | Dealer data for ISM Sumit Kale's territory only |
| 6 | `Abhishek Mishra` | Dealer × Month (ISM-filtered) | Dealer data for ISM Abhishek Mishra's territory only |
| 7 | `Ankit Jaitly` | Dealer × Month (ISM-filtered) | ISM Ankit Jaitly |
| 8 | `Amit Bhakat` | Dealer × Month (ISM-filtered) | ISM Amit Bhakat |
| 9 | `K Sivaramakrishna` | Dealer × Month (ISM-filtered) | ISM K Sivaramakrishna |
| 10 | `Vaibhav Maheshwari` | Dealer × Month (ISM-filtered) | ISM Vaibhav Maheshwari |
| 11 | `Prasanna Kumar` | Dealer × Month (ISM-filtered) | ISM Prasanna Kumar |
| 12 | `Amit Verma` | Dealer × Month (ISM-filtered) | ISM Amit Verma |
| 13 | `Swaminathan S` | Dealer × Month (ISM-filtered) | ISM Swaminathan S |
| 14 | `Vivek Kumar Tiwari` | Dealer × Month (ISM-filtered) | ISM Vivek Kumar Tiwari |

---

### 2.2 Sheet 1: `Dealer wise Model wise` — Detailed Schema

**Header Row Structure (Row 1):**
- Legend text in col A, then model group columns: `Total 2W | APACHE | JUPITER | JUPITER 125 | NTORQ | RADEON | ZEST | SPORT | MOP | Star City Plus | Raider | Ronin | Ntorq 150 | Orbiter | iQube`
- Each model group has **3 columns: Leads | Retails | Conv%**

**Month Columns (Row 2):**
- Excel serial date values representing each month:  
  FY25-26: 46113 (Apr'25), 46143 (May'25), 46174 (Jun'25), 46204 (Jul'25), 46235 (Aug'25), 46266 (Sep'25), 46296 (Oct'25), 46327 (Nov'25), 46357 (Dec'25), 46388 (Jan'26), 46419 (Feb'26), 46447 (Mar'26)  
  FY26-27: 46113 (Apr'26), 46143 (May'26), 46174 (Jun'26), 46204 (Jul'26), 46235 (Aug'26), ...  
  Plus: `FY 25-26` and `FY 26-27` columns for FY totals

**Data Columns (Row 3 = column header):**
```
Dealer code | NAME | Area | City | TTY | State | Zone | ISM Name | NSM | [per model: Leads | Retails | Conv%]
```

**Data Rows:**  
One row per dealer. Grand Total row at the top.

**Key Fields:**
- `Dealer code` — unique 5-digit dealer ID (e.g. 10015, 14307)
- `NAME` — dealer name
- `Area` — geography shortcode (AP, BIH, CHT, DEL, GUJ, HAR, JHA, KAR1, KAR2, KER, MAH1, MAH2, MAH3, MP1, MP2, NE, ODI, PUN, RAJ1, RAJ2, TN1, TN2, TN3, TS, UPC, UPE, UPH, WB, HP, etc.)
- `City` — city name
- `TTY` — Territory name (e.g. `BH6-Muzaffarpur`, `DEL-Delhi 1`, `RAJ1-Jaipur`)
- `State` — full state name (Andhra Pradesh, Bihar, Delhi, Gujarat, Haryana, etc.)
- `Zone` — one of: East | North | South | South 2 | West
- `ISM Name` — one of 10 ISMs (see Sheet 4 names)
- `NSM` — one of 5 NSMs: NSM East | NSM North | NSM South | NSM South 2 | NSM West

**Conversion % Formula:**  
`Conv% = Billed Retails in month XYZ / Leads created in month XYZ`

---

### 2.3 Sheet 2: `Dealer Summary` — Schema

Same structure as Sheet 1 but only **Total 2W** (no model split). Columns per month:  
`Leads | Retails | Conv%`

FY total columns: `FY 25-26 Leads | FY 25-26 Retails | FY 25-26 Conv%`  
Then monthly: each month's `Leads | Retails`  
Then YTD: `FY 26-27 Leads | FY 26-27 Retails | FY 26-27 Conv%`

---

### 2.4 Sheet 3: `Area wise Summary` — Schema

**Row header:** `Area` (27 areas)
**Columns:** Same time structure as Dealer Summary — FY total + monthly columns

Areas covered (27 distinct):
AP, BIH, CHT, DEL, GUJ, HAR, JHA, KAR1, KAR2, KER, MAH1, MAH2, MAH3, MP1, MP2, NE, ODI, PUN, RAJ1, RAJ2, TN1, TN2, TN3, TS, UPC, UPE, UPH, WB, HP, CHD, HP, JK, GOA, etc.

---

### 2.5 Sheet 4: `ISM Wise` — Schema

**Row header:** ISM Name
**ISMs (10 total):**
1. Abhishek Mishra — covers North (UP, BIH, East)
2. Amit Bhakat — covers East (WB, NE, ODI, JH, BIH)
3. Amit Verma — covers North (HAR, DEL, HP, JK, UP)
4. Ankit Jaitly — covers West (MP, CG)
5. K Sivaramakrishna — covers South 2 (KAR, AP, TS)
6. Prasanna Kumar — covers South (TN, KER, PY, A&N)
7. Sumit Kale — covers West (MAH)
8. Swaminathan S — covers South 2 (AP, TS, KER)
9. Vaibhav Maheshwari — covers West (GUJ)
10. Vivek Kumar Tiwari — covers North (RAJ)

**Zone–NSM–ISM Hierarchy:**
```
East (NSM East)
  └── Amit Bhakat (WB, NE, ODI, JH, BIH)
North (NSM North)
  ├── Abhishek Mishra (UPC, UPE, BIH overlap)
  ├── Amit Verma (HAR, DEL, HP, JK, UP)
  └── Vivek Kumar Tiwari (RAJ)
South (NSM South)
  └── Prasanna Kumar (TN, KER, PY, A&N)
South 2 (NSM South 2)
  ├── K Sivaramakrishna (KAR, AP, TS)
  └── Swaminathan S (AP, TS, Goa)
West (NSM West)
  ├── Ankit Jaitly (MP, CG)
  ├── Sumit Kale (MAH)
  └── Vaibhav Maheshwari (GUJ)
```

---

### 2.6 ISM-Specific Sheets (Sheets 5–14)

Each ISM sheet (e.g. `Sumit Kale`) is a **filtered copy** of Dealer Summary restricted to that ISM's dealers. Schema is identical to Dealer Summary.

---

## 3. POWERPOINT DECK ARCHITECTURE

### 3.1 Deck Overview

**File:** `TVS CPS Performance Deck_Aug'26.pptx`  
**Month:** August 2026 edition  
**Total slides:** ~65+ slides  
**Sections:** 7 (matching Agenda on Slide 2)

---

### 3.2 Section Map

| Section # | Slide # | Title | Data Source |
|---|---|---|---|
| Intro | 1 | Title slide | Static |
| Agenda | 2 | Agenda | Static |
| **Section 1** | 3–10 | Lead Generation & Ops Process Flow | Static / Process |
| **Section 2** | 11–15 | TVS CPS Process Snapshot | Excel + History |
| **Section 3** | 16–23 | Zone (NSM)-wise Performance | Excel FY26-27 + FY25-26 |
| **Section 4** | 24–25 | ISM-wise Performance | Excel FY26-27 |
| **Section 5** | 26–65+ | State-wise Performance | Excel FY26-27 + FY25-26 |
| Final | Last | CPS Process Alignment and Asks | Static |

---

### 3.3 Per-Slide Breakdown

#### Slide 1 — Title
- "TVS CPS Performance Overview"
- Month: `Aug-2026` (dynamic — update per month)
- Logos: TVS + BikeDekho

---

#### Slide 2 — Agenda
Static list of 6 sections.

---

#### Slides 3–10 — Lead Gen & Ops Process Flow (STATIC)
All static/process slides. No data from Excel. Update only if process changes.

| Slide | Content |
|---|---|
| 3 | Section divider: "Current Lead Gen & Ops Process Flow" |
| 4 | BikeDekho website lead capture |
| 5 | OTP verification + TVS DMS trigger |
| 6 | SMS & Email to dealership |
| 7 | Consultant processing + lead flow diagram |
| 8 | Two-level audit process (L1 compliance + L2 quality) |
| 9 | BD Operations team process flow |
| 10 | MIS Reports table (report names, audience, frequency) |

---

#### Slide 11 — Section Divider: "Process Snapshot"
Static.

---

#### Slide 12 — CPS Timeline: Growth & Scale *(DYNAMIC)*
**Chart type:** Text callout / visual timeline  
**Data points:** Monthly Average (MA) Leads, MA Retails, Conversion Rate per FY  
```
FY20: MA Leads 76K, MA Retail 9.5K, CR 12.5%
FY21: MA Leads 82K, MA Retail 10.5K, CR 12.8%
FY22: MA Leads 136K, MA Retail 18.3K, CR 13.6%
FY23: MA Leads 185K, MA Retail 25.8K, CR 13.9%
FY24: MA Leads 175K, MA Retail 23.2K, CR 13.2%
FY25: MA Leads 179K, MA Retail 18.8K, CR 10.5%
FY26: Vs FY25 annotation
```
**Source:** Historical averages from both Excel files (grand total rows)

---

#### Slide 13 — Leads Delivered *(DYNAMIC)*
**KPI Table:**
| Field | Value (Aug'26 example) |
|---|---|
| LY ACT (Aug) | 166,636 |
| TY Plan (Aug) | 180,000 |
| TY ACT (Aug) | 157,618 |
| % Ach | 88% |
| Growth (YoY) | -5% |
| YTD LY | 910,651 |
| YTD Plan | 890,000 |
| YTD ACT | 719,860 |

**Line Chart (dual-year monthly trend):**
- Series 1: FY 26-27 monthly leads (Apr→current month)
- Series 2: FY 25-26 monthly leads (Apr→Mar, full year)
- Reference lines: MA FY23-24, MA FY24-25, MA FY25-26, YTD Plan MA, YTD Act MA

**Formula:** Grand Total row of `Dealer Summary` or `Area wise Summary` sheet

---

#### Slide 14 — Retails *(DYNAMIC)*
Identical structure to Slide 13 but for **Retails** values.  
**Source:** Grand Total row, `Dealer Summary` → Retails columns

---

#### Slide 15 — Retail Conversion on Leads *(DYNAMIC)*
**Line chart:** Conversion % (Retails/Leads by creation month)
- Series: FY26-27 monthly conv%, FY25-26 monthly conv%
- Reference lines: MA FY23-24 (13.4%), MA FY24-25 (13.9%), MA FY25-26 (14.5%), YTD Plan (12%), YTD Act

---

#### Slide 16 — Section Divider: "Zone(NSM) wise Performance"
Static.

---

#### Slide 17 — Zone-Level Bar Chart *(DYNAMIC)*
**Clustered bar chart:** 5 Zones × 4 metrics  
Columns: Leads | Retails | Conv% FY26-27 | Conv% FY25-26  
Data for month = current reporting month (Aug'26)  
**Source:** `Area wise Summary` rolled up by Zone

| Zone | Zones / Areas included |
|---|---|
| East | WB, NE, ODI, JH, BIH |
| North | UP (C/E/H), DEL, HAR, PUN, HP, JK, RAJ1, RAJ2 |
| South | TN1, TN2, TN3, KER |
| South 2 | KAR1, KAR2, TS, AP, GOA |
| West | GUJ, MAH1, MAH2, MAH3, MP1, MP2, CHT |

---

#### Slides 18–23 — NSM Zone Trend Charts *(DYNAMIC)*
One slide per NSM Zone (East, South, North, West, South 2).  

**Chart type:** Combo line+bar  
**X-axis:** Monthly (Apr-Mar), with MA FY23-24 / FY24-25 / FY25-26 reference lines  
**Left Y-axis:** Leads + Retails bars  
**Right Y-axis (line):** Conv% FY26-27 and Conv% FY25-26  
**Source:** Sum of areas belonging to each zone in `Area wise Summary`

| Slide | Zone | States / Areas |
|---|---|---|
| 19 | East | WB, BIH, ODI, JH, NE |
| 20 | South | TN, KER, PY, A&N |
| 21 | North | UP, PUN, HP, JK, CHD, RAJ, HAR, DEL |
| 22 | West | GUJ, MH, MP, CG |
| 23 | South 2 | KAR, TS, AP, KER, Goa |

---

#### Slide 24 — Section Divider: "ISM-wise Performance"
Static.

---

#### Slide 25 — ISM Comparison Bar Chart *(DYNAMIC)*
**Clustered bar chart:** All 10 ISMs  
Columns: Leads | Retails | Conv% FY26-27 | Conv% FY25-26  
Data: Month = current reporting month only  
**Source:** `ISM Wise` sheet → current month columns

---

#### Slide 26 — Section Divider: "State-wise Performance"
Static.

---

#### Slides 27–65+ — State Slides *(DYNAMIC)*

**Pattern: Every state gets TWO slides:**

**Slide A: State Trend Chart**
- Combo chart: Monthly Leads + Retails bars + Conv% lines (dual-year)
- Reference lines: MA FY23-24, FY24-25, FY25-26
- **Source:** `Area wise Summary` → sum of areas for that state

**Slide B: State Detail Table**
- Summary row: State code | Last 4 months | FY25-26 full year | FY26-27 YTD
- Top 10 dealers table (highest Conv% in FY26-27)
- Bottom 10 dealers table (lowest Conv% in FY26-27)
- **Source:** `Dealer Summary` filtered by state

**States covered (in order):**
1. Maharashtra (Slides 27–28) — Areas: MAH1, MAH2, MAH3
2. Rajasthan (Slides 29–30) — Areas: RAJ1, RAJ2
3. Madhya Pradesh (Slides 31–32) — Areas: MP1, MP2
4. Bihar (Slides 33–34) — Area: BIH
5. Karnataka & Goa (Slides 35–36) — Areas: KAR1, KAR2, GOA
6. Gujarat (Slides 37–38) — Area: GUJ
7. (Additional states follow same pattern)

---

## 4. DATA FLOW: Excel → PowerPoint

```
Excel FY26-27 file (current year)          Excel FY25-26 file (prior year)
         │                                              │
         ▼                                              ▼
  Dealer Summary sheet              →  LY comparisons (prior year monthly + FY total)
  Area wise Summary sheet           →  Zone & State charts (prior year lines)
  ISM Wise sheet                    →  ISM chart (prior year conv% column)
         │                                              │
         └─────────────────┬────────────────────────────┘
                            ▼
              Aggregation Layer (by Zone / State / ISM / Dealer)
                            │
                            ▼
              PowerPoint Deck (chart data + table values)
```

**Key aggregation rules:**
- **Monthly KPIs (Slide 13-15):** Grand Total row from `Dealer Summary` or `Area wise Summary`
- **Zone rollups:** Sum of areas by zone from `Area wise Summary`
- **State rollups:** Sum of areas by state from `Area wise Summary`
- **ISM data:** Direct from `ISM Wise` sheet
- **Dealer tables (Top/Bottom 10):** From `Dealer Summary`, filtered by state, sorted by FY26-27 Conv%

---

## 5. METRIC DEFINITIONS

| Metric | Definition | Formula |
|---|---|---|
| **Leads** | Online leads created in month XYZ via BikeDekho/Zigwheels | Count by creation month |
| **Retails** | Billed retails in month XYZ attributed to BD/Zigwheels leads | Count by retail/billing month |
| **Conversion %** | Efficiency of lead-to-retail | Billed Retails(month) / Leads created(month) |
| **MA (Monthly Average)** | Average per month for a full financial year | Sum for FY / 12 |
| **YTD** | Year-to-date cumulative from April | Sum Apr–current month |
| **% Ach** | Achievement against plan | TY ACT / TY Plan |
| **Growth** | YoY change | (TY ACT – LY ACT) / LY ACT |
| **LY ACT** | Last year same month actual | From FY25-26 file, same month column |
| **TY Plan** | This year plan/target | Provided separately (not in Excel) |

---

## 6. AUTOMATION PLAN

### 6.1 What Changes Every Month

When a new Excel file is uploaded (e.g. `FY26-27 Nov'26.xlsb`):

1. **Slide 1:** Update month label
2. **Slide 13 & 14:** Update KPI table (new month's LY ACT, TY ACT; recalculate % Ach, Growth, YTD)
3. **Slide 15:** Add new month's conversion point to chart
4. **Slides 17–23:** Update all zone trend charts (add new month bar+line)
5. **Slide 25:** Update ISM bar chart (replace current month values)
6. **All State slides (A):** Update trend chart (add new month)
7. **All State slides (B):** Update summary table row + refresh Top/Bottom dealer lists

### 6.2 What Stays Static

- Slides 3–10 (process flow) — only change if process changes
- Slide 12 (timeline) — update only at FY end
- Section divider slides — static
- Last slide (asks) — static

### 6.3 Automation Script Design

```python
# Pseudo-code for monthly automation

def update_deck(
    excel_new: str,          # Path to new FY26-27 xlsb
    excel_prev: str,         # Path to FY25-26 xlsb
    pptx_template: str,      # Last month's deck as template
    report_month: str,       # e.g. "Sep'26"
    plan_leads: int,         # Monthly lead plan
    plan_retails: int,       # Monthly retail plan
) -> str:                    # Returns path to new deck

    # Step 1: Extract all metrics
    df_dealer = read_dealer_summary(excel_new, excel_prev)
    df_area = read_area_summary(excel_new, excel_prev)
    df_ism = read_ism_summary(excel_new, excel_prev)

    # Step 2: Compute aggregations
    kpis = compute_kpis(df_dealer, report_month, plan_leads, plan_retails)
    zones = compute_zone_rollups(df_area, report_month)
    states = compute_state_rollups(df_area, df_dealer, report_month)
    isms = compute_ism_rollups(df_ism, report_month)

    # Step 3: Update PowerPoint
    prs = load_presentation(pptx_template)
    update_title_slide(prs, report_month)
    update_leads_slide(prs, kpis)
    update_retails_slide(prs, kpis)
    update_conversion_slide(prs, kpis)
    update_zone_slides(prs, zones)
    update_ism_slide(prs, isms)
    for state in states:
        update_state_chart_slide(prs, state)
        update_state_table_slide(prs, state)

    # Step 4: Save output
    output_path = f"TVS_CPS_{report_month}.pptx"
    prs.save(output_path)
    return output_path
```

### 6.4 Chart-to-Data Mapping (for automation)

| PowerPoint Slide | Chart Series | Excel Source | Column |
|---|---|---|---|
| 13 (Leads) | FY26-27 monthly | `Area wise Summary` Grand Total | Month columns (Apr–current) |
| 13 (Leads) | FY25-26 monthly | `Area wise Summary` FY25-26 Grand Total | All 12 month columns |
| 13 (Leads) | MA reference lines | Computed: sum/12 from prior FY files | — |
| 14 (Retails) | Same structure as Slide 13 | — | Retails columns |
| 15 (Conv%) | FY26-27 conv% | Computed: Grand Total Retails/Leads | Per month |
| 17 (Zone) | Leads+Retails+Conv% per zone | `Area wise Summary` → group by Zone | Current month |
| 19–23 (NSM) | Per-zone monthly trend | `Area wise Summary` → filter by Zone areas | All months |
| 25 (ISM) | Per-ISM current month | `ISM Wise` sheet | Current month columns |
| State A | Per-state monthly trend | `Area wise Summary` → filter by State areas | All months |
| State B | Dealer tables | `Dealer Summary` → filter by State | FY26-27 YTD |

---

## 7. NAMING CONVENTIONS & HIERARCHIES

### Geographic Hierarchy
```
Zone (5) → NSM → State → Area Code → TTY → City → Dealer
```

### Area Code to State Mapping

| Area Code | State | Zone |
|---|---|---|
| AP | Andhra Pradesh | South 2 |
| BIH | Bihar | East |
| CHT / CG | Chhattisgarh | West |
| DEL | Delhi | North |
| GUJ | Gujarat | West |
| HAR | Haryana | North |
| HP | Himachal Pradesh | North |
| JHA / JH | Jharkhand | East |
| JK | Jammu & Kashmir | North |
| KAR1 | Karnataka (North) | South 2 |
| KAR2 | Karnataka (South) | South 2 |
| KER | Kerala | South |
| MAH1 | Maharashtra (Pune/South) | West |
| MAH2 | Maharashtra (Mumbai) | West |
| MAH3 | Maharashtra (Aurangabad/Nagpur) | West |
| MP1 | Madhya Pradesh (Bhopal) | West |
| MP2 | Madhya Pradesh (Indore) | West |
| NE | North East (Assam, etc.) | East |
| ODI | Odisha | East |
| PUN | Punjab | North |
| RAJ1 | Rajasthan (Jaipur) | North |
| RAJ2 | Rajasthan (Jodhpur/Udaipur) | North |
| TN1 | Tamil Nadu (North) | South |
| TN2 | Tamil Nadu (Central) | South |
| TN3 | Tamil Nadu (South) | South |
| TS | Telangana | South 2 |
| UPC | Uttar Pradesh (Central) | North |
| UPE | Uttar Pradesh (East) | North |
| UPH | Uttar Pradesh (West/Hills) | North |
| WB | West Bengal | East |
| GOA | Goa | South 2 |
| CHD | Chandigarh | North |

---

## 8. KEY BUSINESS RULES

1. **Lead attribution:** Leads counted by creation month, not by retail month
2. **Retail attribution:** Retails counted by billing month
3. **Conversion % calculation:** Retails(month X) / Leads(month X) — these are from DIFFERENT cohorts
4. **Attribution window:** Lead-to-retail ≤ 60 days; Retail-to-retail ≤ 90 days
5. **Audit requirement:** 100% L1 compliance audit for all retail-identified cases; 10% L2 quality check
6. **Duplicate filter:** TVS EMS filters duplicate leads before dealer assignment
7. **Top/Bottom dealers:** Ranked by FY26-27 YTD Conversion % (not absolute retails)
8. **MA (Monthly Average):** Calculated over all 12 months of a full FY; partial-year FY uses YTD average

---

## 9. PLAN TARGETS

Plan targets (TY Plan column in slides) are provided externally — they are **not stored in the Excel files**. They must be input manually or maintained in a separate config file for the automation script.

Currently known:
- Aug'26: Leads Plan = 180,000; Retails Plan = 20,500
- YTD Aug'26: Leads Plan = 890,000; Retails Plan = 100,000

---

## 10. NEXT STEPS FOR AUTOMATION

Phase 1 (Data extraction):
- [ ] Script to read `.xlsb` files via `pyxlsb` / `pandas`
- [ ] Parse month columns (Excel serial date → month name mapping)
- [ ] Build area→zone mapping lookup table
- [ ] Compute all required aggregations

Phase 2 (PowerPoint update):
- [ ] Map each slide's chart series to data source
- [ ] Script to update chart data in `.pptx` XML (python-pptx or direct XML)
- [ ] Script to update data tables in state detail slides
- [ ] Script to update KPI callout text boxes

Phase 3 (End-to-end pipeline):
- [ ] Accept new `.xlsb` file + month + plan numbers as inputs
- [ ] Generate updated `.pptx` deck
- [ ] Validate output visually

