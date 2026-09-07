"""
Build script for hospital_operations_dashboard.xlsx

Synthetic data only. 10 sites x 5 departments x 12 months = 600 rows.
All derived metrics are Excel formulas; only the Data sheet and the blue
threshold cells contain typed-in numbers.
"""
import calendar
from datetime import date

import numpy as np
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.properties import PageSetupProperties

OUT = "hospital_operations_dashboard.xlsx"
SEED = 42

# ---------- styles ----------
ARIAL = "Arial"
F_TITLE = Font(name=ARIAL, size=14, bold=True)
F_HEAD = Font(name=ARIAL, size=10, bold=True)
F_BODY = Font(name=ARIAL, size=10)
F_ITAL = Font(name=ARIAL, size=10, italic=True, color="595959")
F_INPUT = Font(name=ARIAL, size=10, color="0000FF")
F_CALC = Font(name=ARIAL, size=10, color="000000")
F_CALC_B = Font(name=ARIAL, size=10, color="000000", bold=True)
F_LINK = Font(name=ARIAL, size=10, color="008000")
F_WARN = Font(name=ARIAL, size=10, bold=True, color="C00000")

FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_HEAD = PatternFill("solid", fgColor="D9E1F2")
FILL_KEY = PatternFill("solid", fgColor="E2EFDA")
FILL_NOTE = PatternFill("solid", fgColor="FFF2CC")
FILL_RED = PatternFill("solid", fgColor="F8696B")
FILL_AMBER = PatternFill("solid", fgColor="FFEB84")
FILL_GREEN = PatternFill("solid", fgColor="63BE7B")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")

FMT_INT = "#,##0"
FMT_PCT = "0.0%"
FMT_DEC1 = "0.0"
FMT_DEC2 = "0.00"
FMT_DATE = "mmm yyyy"
FMT_SIGNED = "+#,##0;-#,##0;0"
FMT_SIGNED_DEC = "+0.0;-0.0;0.0"
FMT_SIGNED_PCT = "+0.0%;-0.0%;0.0%"


def header(ws, row, labels, col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font = F_HEAD
        c.fill = FILL_HEAD
        c.border = BOX
        c.alignment = Alignment(vertical="center", wrap_text=True)


def add_name(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def as_text(cell):
    cell.data_type = "s"
    return cell


# =====================================================================
# Synthetic data generation
# =====================================================================
rng = np.random.default_rng(SEED)

DEPTS = ["Cardiology", "Orthopaedics", "General Medicine", "Nephrology", "Paediatrics"]
# base admissions per site-month, LOS, billing days, denial rate, beds per site
PROFILE = {
    "Cardiology":       dict(adm=105, los=5.2, bill=10.5, deny=0.08, beds=22),
    "Orthopaedics":     dict(adm=75,  los=6.1, bill=14.0, deny=0.13, beds=19),
    "General Medicine": dict(adm=150, los=4.0, bill=8.0,  deny=0.06, beds=24),
    "Nephrology":       dict(adm=55,  los=7.8, bill=12.5, deny=0.09, beds=18),
    "Paediatrics":      dict(adm=85,  los=3.1, bill=6.5,  deny=0.05, beds=11),
}
SITES = [f"Site {i:02d}" for i in range(1, 11)]
MONTHS = [date(2025, m, 1) for m in range(1, 13)]
# mild seasonality: winter and monsoon bumps for admissions (index by month-1)
SEASON = np.array([1.08, 1.05, 1.00, 0.97, 0.96, 0.98, 1.03, 1.06, 1.02, 0.98, 1.00, 1.04])
site_scale = rng.uniform(0.85, 1.15, size=len(SITES))

rows = []
rid = 1
for mi, month in enumerate(MONTHS):
    days = calendar.monthrange(month.year, month.month)[1]
    for si, site in enumerate(SITES):
        for dept in DEPTS:
            p = PROFILE[dept]
            adm = int(round(p["adm"] * SEASON[mi] * site_scale[si] * rng.normal(1.0, 0.07)))
            adm = int(np.clip(adm, 40, 180))
            dis = int(np.clip(adm + rng.integers(-6, 7), 35, 185))
            los_target = p["los"] * rng.normal(1.0, 0.06) * (1.03 if mi in (0, 1, 11) else 1.0)
            los_target = float(np.clip(los_target, 2.5, 8.5))
            bed_days = int(round(dis * los_target))
            avail = p["beds"] * days                      # fixed capacity per dept per site per month
            if bed_days > avail:                          # never exceed capacity
                bed_days = int(avail * rng.uniform(0.94, 0.99))
            los = round(bed_days / dis, 2)                # LOS reconciles exactly to bed days / discharges
            bill = round(float(np.clip(p["bill"] * rng.normal(1.0, 0.12), 4.0, 22.0)), 1)
            subm = int(round(dis * rng.uniform(0.86, 0.94)))
            deny_rate = float(np.clip(p["deny"] * rng.normal(1.0, 0.18), 0.03, 0.15))
            denied = int(round(subm * deny_rate))
            denied = min(denied, subm)
            rows.append([f"HOSP-{rid:04d}", month, site, dept, adm, dis, bed_days, avail, los, bill, subm, denied])
            rid += 1

N = len(rows)
LAST = N + 1  # last data row index (header is row 1)
assert N == 600

wb = Workbook()

# =====================================================================
# Sheet 1: Read Me
# =====================================================================
ws = wb.active
ws.title = "Read Me"
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 115
lines = [
    ("Hospital Operations Dashboard", F_TITLE),
    ("", F_BODY),
    ("PURPOSE", F_HEAD),
    ("A monthly operations pack for a ten-site hospital network across five departments. It takes a flat transactional table,", F_BODY),
    ("derives per-row metrics, summarises them by department and by month with SUMIFS / SUMPRODUCT, flags threshold breaches,", F_BODY),
    ("and shows the result on a single dashboard with KPI cards and four charts. The structure mirrors a hospital analytics", F_BODY),
    ("internship deliverable: define the KPIs, build the reporting layer, monitor against targets.", F_BODY),
    ("", F_BODY),
    ("THE FIVE METRICS (be able to define all of these cold)", F_HEAD),
    ("1. Average Length of Stay (days) = Total bed days / Discharges. Weighted by discharges when aggregated.", F_BODY),
    ("2. Bed Occupancy Rate (%) = Total bed days used / Available bed days (beds x days in month).", F_BODY),
    ("3. Admissions and Discharges (counts) = patients admitted and patients discharged in the month. Their gap is the change in census.", F_BODY),
    ("4. Billing Cycle Time (days) = days from discharge to bill finalisation. Weighted by discharges when aggregated.", F_BODY),
    ("5. Claim Denial Rate (%) = Claims denied / Claims submitted.", F_BODY),
    ("", F_BODY),
    ("AN ANALYTICAL POINT WORTH RAISING IN INTERVIEW", F_HEAD),
    ("Average LOS across departments is a weighted average (total bed days / total discharges), not a plain average of department averages.", F_BODY),
    ("A plain average would let a 40-discharge Nephrology row count as much as a 180-discharge General Medicine row. The Summary sheet", F_BODY),
    ("uses SUMPRODUCT for exactly this reason. The same applies to billing cycle time and denial rate.", F_BODY),
    ("", F_BODY),
    ("HOW TO READ THE SHEETS", F_HEAD),
    ("Data               -> 600 synthetic rows: 10 sites x 5 departments x 12 months. Typed-in values, no formulas.", F_BODY),
    ("Calculations       -> per-row derived metrics (occupancy, denial rate, discharge variance, LOS check). All formulas.", F_BODY),
    ("Summary            -> metrics by department, metrics by month, and a threshold table with red / amber / green flags.", F_BODY),
    ("Dashboard          -> KPI cards (current month vs prior month) and four native Excel charts.", F_BODY),
    ("Metric Definitions -> one row per metric. EDIT THIS IN YOUR OWN WORDS before you use the file.", F_BODY),
    ("", F_BODY),
    ("COLOUR LEGEND", F_HEAD),
    ("Blue text, yellow fill = input you may change (threshold targets, tolerance)", F_INPUT),
    ("Black text = calculated on this sheet", F_CALC),
    ("Green text = pulled from another sheet", F_LINK),
    ("", F_BODY),
    ("NOTE ON THE DATA", F_HEAD),
    ("The spec called for ~600 rows of department-by-month data. Five departments x twelve months is only 60 rows, so the table is", F_BODY),
    ("modelled as a ten-site network (Site column) to reach 600 rows while keeping every row a meaningful department-month record.", F_BODY),
    ("Departments have distinct profiles (Nephrology longer stays, Orthopaedics higher denials, Paediatrics short stays) and there is", F_BODY),
    ("mild winter / monsoon seasonality in admissions. Data is generated by build_hospital_dashboard.py with a fixed random seed.", F_BODY),
    ("", F_BODY),
    ("DISCLAIMER", F_WARN),
    ("All data in this workbook is synthetically generated for demonstration purposes.", F_WARN),
    ("It contains no real patient, hospital, or payer data.", F_WARN),
    ("", F_BODY),
    ("Built with openpyxl. No macros, no Power Query, no array formulas. Functions used: SUMIFS, SUMPRODUCT, INDEX, MATCH, IFERROR, IF, MAX, EDATE, ABS.", F_ITAL),
    ("A PivotTable can be added on the Data sheet in Excel (Insert > PivotTable) in two clicks; it is not generated here because openpyxl cannot write pivot caches.", F_ITAL),
]
for i, (text, font) in enumerate(lines, start=2):
    c = ws.cell(row=i, column=2, value=text)
    c.font = font
    c.alignment = Alignment(wrap_text=True, vertical="top")

# =====================================================================
# Sheet 2: Data
# =====================================================================
wd = wb.create_sheet("Data")
cols = ["Record ID", "Month", "Site", "Department", "Admissions", "Discharges", "Total Bed Days",
        "Available Bed Days", "Length of Stay", "Billing Cycle Days", "Claims Submitted", "Claims Denied"]
header(wd, 1, cols)
fmts = [None, FMT_DATE, None, None, FMT_INT, FMT_INT, FMT_INT, FMT_INT, FMT_DEC2, FMT_DEC1, FMT_INT, FMT_INT]
for r, row in enumerate(rows, start=2):
    for ci, v in enumerate(row, start=1):
        c = wd.cell(row=r, column=ci, value=v)
        c.font = F_BODY
        if fmts[ci - 1]:
            c.number_format = fmts[ci - 1]
widths = [11, 11, 9, 17, 11, 11, 14, 16, 13, 15, 15, 13]
for i, w in enumerate(widths):
    wd.column_dimensions[chr(ord("A") + i)].width = w
wd.freeze_panes = "A2"
wd.auto_filter.ref = f"A1:L{LAST}"

# named ranges for the data columns (readable SUMIFS)
names = {
    "Month_Col": "B", "Site_Col": "C", "Dept_Col": "D", "Admissions": "E", "Discharges": "F",
    "Bed_Days": "G", "Available_Bed_Days": "H", "LOS": "I", "Billing_Days": "J",
    "Claims_Submitted": "K", "Claims_Denied": "L",
}
for nm, col in names.items():
    add_name(wb, nm, f"Data!${col}$2:${col}${LAST}")

# =====================================================================
# Sheet 3: Calculations
# =====================================================================
wc = wb.create_sheet("Calculations")
header(wc, 1, ["Record ID", "Month", "Department", "Occupancy Rate", "Denial Rate",
               "Discharge Variance", "LOS check (bed days / discharges)", "LOS matches Data?"])
for r in range(2, LAST + 1):
    wc.cell(row=r, column=1, value=f"=Data!A{r}").font = F_LINK
    c = wc.cell(row=r, column=2, value=f"=Data!B{r}"); c.font = F_LINK; c.number_format = FMT_DATE
    wc.cell(row=r, column=3, value=f"=Data!D{r}").font = F_LINK
    c = wc.cell(row=r, column=4, value=f"=IFERROR(Data!G{r}/Data!H{r},0)"); c.font = F_CALC; c.number_format = FMT_PCT
    c = wc.cell(row=r, column=5, value=f"=IFERROR(Data!L{r}/Data!K{r},0)"); c.font = F_CALC; c.number_format = FMT_PCT
    c = wc.cell(row=r, column=6, value=f"=Data!F{r}-Data!E{r}"); c.font = F_CALC; c.number_format = FMT_SIGNED
    c = wc.cell(row=r, column=7, value=f"=IFERROR(Data!G{r}/Data!F{r},0)"); c.font = F_CALC; c.number_format = FMT_DEC2
    c = wc.cell(row=r, column=8, value=f'=IF(ABS(G{r}-Data!I{r})<Summary!$B$4,"OK","CHECK")'); c.font = F_CALC
for i, w in enumerate([11, 11, 17, 15, 12, 17, 26, 16]):
    wc.column_dimensions[chr(ord("A") + i)].width = w
wc.freeze_panes = "A2"
wc["J1"] = "Formulas (row 2 shown as text):"
wc["J1"].font = F_HEAD
notes = [
    "Occupancy = IFERROR(Data!G2 / Data!H2, 0)   bed days used / bed days available",
    "Denial rate = IFERROR(Data!L2 / Data!K2, 0)   claims denied / claims submitted",
    "Discharge variance = Data!F2 - Data!E2   discharges minus admissions",
    "LOS check = IFERROR(Data!G2 / Data!F2, 0)   recomputes LOS from bed days; column H flags any row where it drifts from the stored LOS by more than the tolerance on Summary!B4",
]
for i, n in enumerate(notes, start=2):
    wc.cell(row=i, column=10, value=n).font = F_ITAL
wc.column_dimensions["J"].width = 100
wc.conditional_formatting.add(f"H2:H{LAST}", CellIsRule(operator="equal", formula=['"CHECK"'], fill=FILL_RED))

# =====================================================================
# Sheet 4: Summary
# =====================================================================
wsu = wb.create_sheet("Summary")
wsu.column_dimensions["A"].width = 20
for col in "BCDEFGHIJ":
    wsu.column_dimensions[col].width = 15
wsu["A1"] = "Summary tables (SUMIFS / SUMPRODUCT over the Data sheet)"
wsu["A1"].font = F_TITLE
wsu["A2"] = "Every figure below is a formula over named ranges on the Data sheet (Admissions, Discharges, Bed_Days, ...). Nothing is typed in except the blue cells."
wsu["A2"].font = F_ITAL

wsu["A4"] = "LOS rounding tolerance"
wsu["A4"].font = F_BODY
c = wsu["B4"]; c.value = 0.01; c.font = F_INPUT; c.fill = FILL_INPUT; c.number_format = FMT_DEC2; c.border = BOX
wsu["C4"] = "Used by Calculations!H. Stored LOS is rounded to 2 dp, so a 0.01 tolerance is expected."
wsu["C4"].font = F_ITAL

# --- By department (rows 8-13)
wsu["A6"] = "METRICS BY DEPARTMENT (full year)"
wsu["A6"].font = F_HEAD
dep_heads = ["Department", "Admissions", "Discharges", "Total Bed Days", "Available Bed Days",
             "Occupancy Rate", "Avg LOS (weighted)", "Billing Cycle (weighted)", "Claims Submitted",
             "Claims Denied", "Denial Rate"]
header(wsu, 7, dep_heads)
fm = {2: FMT_INT, 3: FMT_INT, 4: FMT_INT, 5: FMT_INT, 6: FMT_PCT, 7: FMT_DEC2, 8: FMT_DEC1, 9: FMT_INT, 10: FMT_INT, 11: FMT_PCT}
for i, d in enumerate(DEPTS):
    r = 8 + i
    wsu.cell(row=r, column=1, value=d).font = F_BODY
    f = {
        2: f"=SUMIFS(Admissions,Dept_Col,$A{r})",
        3: f"=SUMIFS(Discharges,Dept_Col,$A{r})",
        4: f"=SUMIFS(Bed_Days,Dept_Col,$A{r})",
        5: f"=SUMIFS(Available_Bed_Days,Dept_Col,$A{r})",
        6: f"=IFERROR(D{r}/E{r},0)",
        7: f"=IFERROR(D{r}/C{r},0)",
        8: f"=IFERROR(SUMPRODUCT((Dept_Col=$A{r})*Billing_Days*Discharges)/C{r},0)",
        9: f"=SUMIFS(Claims_Submitted,Dept_Col,$A{r})",
        10: f"=SUMIFS(Claims_Denied,Dept_Col,$A{r})",
        11: f"=IFERROR(J{r}/I{r},0)",
    }
    for col, formula in f.items():
        c = wsu.cell(row=r, column=col, value=formula)
        c.font = F_CALC
        c.number_format = fm[col]
    for col in range(1, 12):
        wsu.cell(row=r, column=col).border = BOX
# network total row 13
r = 13
wsu.cell(row=r, column=1, value="Network total").font = F_HEAD
tot = {
    2: "=SUM(B8:B12)", 3: "=SUM(C8:C12)", 4: "=SUM(D8:D12)", 5: "=SUM(E8:E12)",
    6: "=IFERROR(D13/E13,0)", 7: "=IFERROR(D13/C13,0)",
    8: "=IFERROR(SUMPRODUCT(Billing_Days,Discharges)/C13,0)",
    9: "=SUM(I8:I12)", 10: "=SUM(J8:J12)", 11: "=IFERROR(J13/I13,0)",
}
for col, formula in tot.items():
    c = wsu.cell(row=r, column=col, value=formula)
    c.font = F_CALC_B
    c.number_format = fm[col]
for col in range(1, 12):
    wsu.cell(row=r, column=col).border = BOX
    wsu.cell(row=r, column=col).fill = FILL_KEY
wsu["A14"] = "Weighted LOS = total bed days / total discharges. Plain average of the five department LOS values would be wrong; see Read Me."
wsu["A14"].font = F_ITAL
wsu["A15"] = "Weighted billing cycle = SUMPRODUCT((Dept_Col = dept) * Billing_Days * Discharges) / Discharges. The boolean test inside SUMPRODUCT acts as the filter."
wsu["A15"].font = F_ITAL

# --- By month (rows 19-30)
wsu["A17"] = "METRICS BY MONTH (all sites, all departments)"
wsu["A17"].font = F_HEAD
mon_heads = ["Month", "Admissions", "Discharges", "Total Bed Days", "Available Bed Days",
             "Occupancy Rate", "Avg LOS (weighted)", "Billing Cycle (weighted)", "Claims Submitted",
             "Claims Denied", "Denial Rate"]
header(wsu, 18, mon_heads)
for i, m in enumerate(MONTHS):
    r = 19 + i
    c = wsu.cell(row=r, column=1, value=m)
    c.font = F_INPUT
    c.number_format = FMT_DATE
    f = {
        2: f"=SUMIFS(Admissions,Month_Col,$A{r})",
        3: f"=SUMIFS(Discharges,Month_Col,$A{r})",
        4: f"=SUMIFS(Bed_Days,Month_Col,$A{r})",
        5: f"=SUMIFS(Available_Bed_Days,Month_Col,$A{r})",
        6: f"=IFERROR(D{r}/E{r},0)",
        7: f"=IFERROR(D{r}/C{r},0)",
        8: f"=IFERROR(SUMPRODUCT((Month_Col=$A{r})*Billing_Days*Discharges)/C{r},0)",
        9: f"=SUMIFS(Claims_Submitted,Month_Col,$A{r})",
        10: f"=SUMIFS(Claims_Denied,Month_Col,$A{r})",
        11: f"=IFERROR(J{r}/I{r},0)",
    }
    for col, formula in f.items():
        c = wsu.cell(row=r, column=col, value=formula)
        c.font = F_CALC
        c.number_format = fm[col]
    for col in range(1, 12):
        wsu.cell(row=r, column=col).border = BOX
wsu["A31"] = "Month labels in column A are the twelve first-of-month dates that appear in Data. They are the only typed-in values in this table (blue)."
wsu["A31"].font = F_ITAL
add_name(wb, "Summary_Months", "Summary!$A$19:$A$30")

# --- Threshold table (rows 36-41)
wsu["A33"] = "THRESHOLD MONITOR (latest month vs target)"
wsu["A33"].font = F_HEAD
wsu["A34"] = "Latest month in Data"
wsu["A34"].font = F_BODY
c = wsu["B34"]; c.value = "=MAX(Month_Col)"; c.font = F_LINK; c.number_format = FMT_DATE; c.border = BOX
wsu["C34"] = "Amber tolerance"
wsu["C34"].font = F_BODY
c = wsu["D34"]; c.value = 0.10; c.font = F_INPUT; c.fill = FILL_INPUT; c.number_format = FMT_PCT; c.border = BOX
wsu["E34"] = "Amber = within this % of target on the wrong side. Illustrative - not sourced."
wsu["E34"].font = F_ITAL

header(wsu, 36, ["Metric", "Target", "Better when", "Latest month actual", "Status", "Note"])
thr = [
    (37, "Avg LOS (days)", 4.75, "Lower", "G", FMT_DEC2),
    (38, "Bed Occupancy Rate", 0.85, "Higher", "F", FMT_PCT),
    (39, "Discharges (count)", 5000, "Higher", "C", FMT_INT),
    (40, "Billing Cycle (days)", 8.5, "Lower", "H", FMT_DEC1),
    (41, "Claim Denial Rate", 0.07, "Lower", "K", FMT_PCT),
]
for r, lab, target, direction, src_col, fmt in thr:
    wsu.cell(row=r, column=1, value=lab).font = F_BODY
    c = wsu.cell(row=r, column=2, value=target); c.font = F_INPUT; c.fill = FILL_INPUT; c.number_format = fmt
    c = wsu.cell(row=r, column=3, value=direction); c.font = F_INPUT; c.fill = FILL_INPUT
    c = wsu.cell(row=r, column=4, value=f"=INDEX(${src_col}$19:${src_col}$30,MATCH($B$34,Summary_Months,0))")
    c.font = F_CALC; c.number_format = fmt
    status = (f'=IF($C{r}="Lower",'
              f'IF($D{r}<=$B{r},"Green",IF($D{r}<=$B{r}*(1+$D$34),"Amber","Red")),'
              f'IF($D{r}>=$B{r},"Green",IF($D{r}>=$B{r}*(1-$D$34),"Amber","Red")))')
    c = wsu.cell(row=r, column=5, value=status); c.font = F_CALC_B; c.alignment = Alignment(horizontal="center")
    wsu.cell(row=r, column=6, value="Target is illustrative - not sourced. Set it from your own hospital's policy.").font = F_ITAL
    for col in range(1, 7):
        wsu.cell(row=r, column=col).border = BOX
wsu.conditional_formatting.add("E37:E41", CellIsRule(operator="equal", formula=['"Red"'], fill=FILL_RED))
wsu.conditional_formatting.add("E37:E41", CellIsRule(operator="equal", formula=['"Amber"'], fill=FILL_AMBER))
wsu.conditional_formatting.add("E37:E41", CellIsRule(operator="equal", formula=['"Green"'], fill=FILL_GREEN))
wsu["A43"] = "Status logic (row 37 as text):"
wsu["A43"].font = F_HEAD
wsu["A44"] = '=IF(C37="Lower", IF(D37<=B37,"Green", IF(D37<=B37*(1+D34),"Amber","Red")), IF(D37>=B37,"Green", IF(D37>=B37*(1-D34),"Amber","Red")))'
as_text(wsu["A44"])
wsu["A44"].font = F_ITAL
wsu["A45"] = "Latest month actual = INDEX(month table column, MATCH(latest month, Summary_Months, 0)). Change the target or tolerance (blue) and the flags update."
wsu["A45"].font = F_ITAL
wsu.freeze_panes = "A3"

# =====================================================================
# Sheet 5: Dashboard
# =====================================================================
wdb = wb.create_sheet("Dashboard")
wdb.column_dimensions["A"].width = 24
for col in "BCDE":
    wdb.column_dimensions[col].width = 15
wdb["A1"] = "Operations Dashboard"
wdb["A1"].font = F_TITLE
wdb["A2"] = '=CONCATENATE("Latest month: ",TEXT(Summary!$B$34,"mmm yyyy"))'
wdb["A2"].font = F_LINK
wdb["A3"] = "Current month"
wdb["A3"].font = F_BODY
c = wdb["B3"]; c.value = "=Summary!$B$34"; c.font = F_LINK; c.number_format = FMT_DATE; c.border = BOX
wdb["A4"] = "Prior month"
wdb["A4"].font = F_BODY
c = wdb["B4"]; c.value = "=EDATE(B3,-1)"; c.font = F_CALC; c.number_format = FMT_DATE; c.border = BOX
wdb["C3"] = "Row in month table"; wdb["C3"].font = F_ITAL
c = wdb["D3"]; c.value = "=MATCH(B3,Summary_Months,0)"; c.font = F_CALC; c.border = BOX
c = wdb["D4"]; c.value = "=MATCH(B4,Summary_Months,0)"; c.font = F_CALC; c.border = BOX
wdb["C4"] = "(INDEX/MATCH keys)"; wdb["C4"].font = F_ITAL

header(wdb, 6, ["KPI", "Current month", "Prior month", "Variance", "Variance %"])
kpis = [
    (7, "Avg Length of Stay (days)", "G", FMT_DEC2, FMT_SIGNED_DEC),
    (8, "Bed Occupancy Rate", "F", FMT_PCT, FMT_SIGNED_PCT),
    (9, "Admissions", "B", FMT_INT, FMT_SIGNED),
    (10, "Discharges", "C", FMT_INT, FMT_SIGNED),
    (11, "Billing Cycle Time (days)", "H", FMT_DEC1, FMT_SIGNED_DEC),
    (12, "Claim Denial Rate", "K", FMT_PCT, FMT_SIGNED_PCT),
]
for r, lab, col, fmt, vfmt in kpis:
    wdb.cell(row=r, column=1, value=lab).font = F_HEAD
    c = wdb.cell(row=r, column=2, value=f"=INDEX(Summary!${col}$19:${col}$30,$D$3)"); c.number_format = fmt
    c.font = Font(name=ARIAL, size=12, bold=True, color="008000")
    c = wdb.cell(row=r, column=3, value=f"=INDEX(Summary!${col}$19:${col}$30,$D$4)"); c.font = F_LINK; c.number_format = fmt
    c = wdb.cell(row=r, column=4, value=f"=B{r}-C{r}"); c.font = F_CALC; c.number_format = vfmt
    c = wdb.cell(row=r, column=5, value=f"=IFERROR(D{r}/C{r},0)"); c.font = F_CALC; c.number_format = FMT_SIGNED_PCT
    for cc in range(1, 6):
        wdb.cell(row=r, column=cc).border = BOX
        wdb.cell(row=r, column=cc).fill = FILL_KEY if cc == 2 else PatternFill()
    wdb.row_dimensions[r].height = 22
wdb["A13"] = "Current = INDEX(Summary month-table column, row of latest month). Prior = same with EDATE(latest, -1)."
wdb["A13"].font = F_ITAL

# charts
months_ref = Reference(wsu, min_col=1, min_row=19, max_row=30)

ch1 = LineChart()
ch1.title = "Average LOS by month (days, weighted)"
ch1.y_axis.title = "Days"
ch1.x_axis.number_format = "mmm"
ch1.add_data(Reference(wsu, min_col=7, min_row=18, max_row=30), titles_from_data=True)
ch1.set_categories(months_ref)
ch1.height, ch1.width = 7.5, 15
ch1.legend = None
wdb.add_chart(ch1, "G2")

ch2 = LineChart()
ch2.title = "Bed occupancy rate by month"
ch2.y_axis.title = "Occupancy"
ch2.y_axis.number_format = "0%"
ch2.x_axis.number_format = "mmm"
ch2.add_data(Reference(wsu, min_col=6, min_row=18, max_row=30), titles_from_data=True)
ch2.set_categories(months_ref)
ch2.height, ch2.width = 7.5, 15
ch2.legend = None
wdb.add_chart(ch2, "P2")

ch3 = BarChart()
ch3.type = "col"
ch3.title = "Claim denial rate by department"
ch3.y_axis.title = "Denial rate"
ch3.y_axis.number_format = "0%"
ch3.add_data(Reference(wsu, min_col=11, min_row=7, max_row=12), titles_from_data=True)
ch3.set_categories(Reference(wsu, min_col=1, min_row=8, max_row=12))
ch3.height, ch3.width = 7.5, 15
ch3.legend = None
wdb.add_chart(ch3, "G18")

ch4 = BarChart()
ch4.type = "col"
ch4.grouping = "clustered"
ch4.title = "Admissions vs discharges by month"
ch4.y_axis.title = "Patients"
ch4.x_axis.number_format = "mmm"
ch4.add_data(Reference(wsu, min_col=2, max_col=3, min_row=18, max_row=30), titles_from_data=True)
ch4.set_categories(months_ref)
ch4.height, ch4.width = 7.5, 15
wdb.add_chart(ch4, "P18")

# =====================================================================
# Sheet 6: Metric Definitions
# =====================================================================
wm = wb.create_sheet("Metric Definitions")
wm.column_dimensions["A"].width = 26
wm.column_dimensions["B"].width = 44
wm.column_dimensions["C"].width = 52
wm.column_dimensions["D"].width = 52
wm["A1"] = "Metric Definitions"
wm["A1"].font = F_TITLE
wm["A2"] = "Starter text. Rewrite every cell in your own words before this file goes anywhere near an application. Yellow cells are yours to edit."
wm["A2"].font = F_WARN
header(wm, 4, ["Metric", "Formula in words", "Why it matters operationally", "Common pitfall"])
defs = [
    ("Average Length of Stay (days)",
     "Total bed days used in the period divided by discharges in the period.",
     "Drives bed capacity. A half-day reduction across a busy ward frees beds without building any.",
     "Averaging department averages instead of weighting by discharges. Also: including patients still in the bed (no discharge yet) distorts the numerator."),
    ("Bed Occupancy Rate (%)",
     "Bed days used divided by bed days available (staffed beds x days in month).",
     "Too low wastes fixed cost; too high (above roughly 85-90%) means no slack for emergencies and longer ED waits.",
     "Using licensed beds instead of staffed beds in the denominator. Mid-month bed closures also break a fixed-capacity assumption."),
    ("Admissions and Discharges (counts)",
     "Number of patients admitted, and number discharged, in the period.",
     "Throughput. When discharges lag admissions the census rises and occupancy climbs.",
     "Counting transfers between departments as new admissions, which double counts the patient at network level."),
    ("Billing Cycle Time (days)",
     "Days from discharge date to the date the final bill is raised, averaged over discharges.",
     "Every day of delay is cash the hospital has not collected. Also a proxy for documentation quality.",
     "Measuring from admission instead of discharge, which mixes clinical LOS into a finance metric."),
    ("Claim Denial Rate (%)",
     "Claims denied by the payer divided by claims submitted, in the period.",
     "Denied claims mean rework, delayed cash, and sometimes write-offs. Department-level rates point to coding or documentation gaps.",
     "Counting first-pass denials and final denials together. Many first-pass denials are overturned on resubmission."),
]
for i, (m, f, w, p) in enumerate(defs, start=5):
    for col, val in enumerate((m, f, w, p), start=1):
        c = wm.cell(row=i, column=col, value=val)
        c.font = F_HEAD if col == 1 else F_INPUT
        c.fill = PatternFill() if col == 1 else FILL_NOTE
        c.alignment = WRAP
        c.border = BOX
    wm.row_dimensions[i].height = 62

# ---------- Arial sweep ----------
for sheet in wb.worksheets:
    for row in sheet.iter_rows():
        for cell in row:
            if cell.font is None or cell.font.name != ARIAL:
                cell.font = Font(name=ARIAL, size=cell.font.size or 10, bold=cell.font.bold,
                                 italic=cell.font.italic, color=cell.font.color)

# ---------- print setup: one page wide per sheet (for README screenshots) ----------
for sheet in wb.worksheets:
    sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0 if sheet.title in ("Data", "Calculations") else 1

wb.save(OUT)
print("saved", OUT, "rows:", N)
