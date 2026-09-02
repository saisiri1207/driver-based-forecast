#!/usr/bin/env python3
"""Build Northline_Driver_Based_Forecast.xlsx — driver-based P&L, plan vs reforecast."""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart.marker import Marker

yellow = PatternFill("solid", fgColor="FFF2CC")
header_fill = PatternFill("solid", fgColor="1F4E79")
section_fill = PatternFill("solid", fgColor="D6E3F0")
green_fill = PatternFill("solid", fgColor="C6EFCE")
amber_fill = PatternFill("solid", fgColor="FFE699")
red_fill = PatternFill("solid", fgColor="F8CBAD")
tile_fill = PatternFill("solid", fgColor="E9EDF4")
light_gray = PatternFill("solid", fgColor="F5F5F5")
input_font = Font(name="Calibri", size=11, color="0000FF")
black = Font(name="Calibri", size=11, color="000000")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
title_font = Font(name="Calibri", size=16, bold=True, color="1F4E79")
section_font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
bold = Font(name="Calibri", size=11, bold=True)
bold_black = Font(name="Calibri", size=11, bold=True, color="000000")
italic_grey = Font(name="Calibri", size=10, italic=True, color="666666")
small_grey = Font(name="Calibri", size=9, italic=True, color="666666")
link_font = Font(name="Calibri", size=11, color="0563C1", underline="single")
thin = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
money = '_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)'
pct = "0.0%"
num = "#,##0.0"
ppu = "0.00"
int_fmt = "#,##0"

MONTHS = [
    "Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26", "Jun-26",
    "Jul-26", "Aug-26", "Sep-26", "Oct-26", "Nov-26", "Dec-26",
]
PRODUCTS = ["RTD Beverages", "Salty Snacks", "Household Care", "Personal Care"]

# Plan unit economics
PLAN_PRICE = [11.90, 14.20, 16.50, 18.80]
FCST_PRICE = [11.75, 14.50, 16.55, 19.00]
MAT = [5.40, 6.20, 7.40, 8.10]
LAB = [1.05, 1.20, 1.35, 1.50]
OH = [0.85, 0.95, 1.10, 1.25]

# Seasonal index (Q4 holiday lift), mean ≈ 1
SEAS = [0.90, 0.86, 0.93, 0.95, 0.98, 1.00, 1.02, 1.01, 0.98, 1.04, 1.11, 1.19]
smean = sum(SEAS) / 12
SEAS = [x / smean for x in SEAS]

ANN_VOL = [31800, 17550, 10850, 7580]  # 000 cases
PLAN_VOL = [[round(ann * m / 12) for m in SEAS] for ann in ANN_VOL]
# Rebalance last month so annual totals land
for i, ann in enumerate(ANN_VOL):
    gap = ann - sum(PLAN_VOL[i])
    PLAN_VOL[i][-1] += gap

# Latest-view / reforecast volume (Bev softness, snacks + PC strength)
FCST_LIFT = [0.982, 1.048, 1.008, 1.062]
FCST_VOL = [[round(v * FCST_LIFT[i]) for v in PLAN_VOL[i]] for i in range(4)]

# Monthly fixed OpEx $000s — plan
PLAN_SELL_FIX = [1100] * 12
PLAN_MKT_FIX = [900, 800, 1100, 1400, 1600, 1800, 1500, 1400, 1300, 2000, 2800, 3200]
PLAN_DIST_FIX = [650] * 12
PLAN_GA_FIX = [5100] * 12
PLAN_RD_FIX = [1550] * 12
# Forecast fixed: marketing up in H2, G&A +50, R&D +80, dist -20, selling flat
FCST_SELL_FIX = [1100] * 12
FCST_MKT_FIX = [920, 820, 1120, 1450, 1680, 1900, 1620, 1520, 1420, 2200, 3100, 3500]
FCST_DIST_FIX = [630] * 12
FCST_GA_FIX = [5150] * 12
FCST_RD_FIX = [1630] * 12

ASSUMP = "01_Assumptions"
REV = "02_Revenue_Build"
COGS = "03_COGS_Build"
OPEX = "04_OpEx"
PL = "05_P&L_Forecast"
BR = "06_Bridge"


def style_input(cell):
    cell.fill = yellow
    cell.font = input_font
    cell.border = thin
    cell.alignment = Alignment(horizontal="center")


def style_formula(cell, key=False):
    cell.font = bold_black if key else black
    cell.border = thin
    cell.alignment = Alignment(horizontal="center")
    if key:
        cell.fill = green_fill


def style_header_cell(cell, value):
    cell.value = value
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
    cell.border = thin


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def month_headers(ws, row, start_col=3, extra=None):
    for i, m in enumerate(MONTHS):
        style_header_cell(ws.cell(row, start_col + i), m)
    extras = extra if extra is not None else ["FY26"]
    for j, lab in enumerate(extras):
        style_header_cell(ws.cell(row, start_col + 12 + j), lab)


def landscape(ws, fit_height=0):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_height
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
    ws.sheet_view.showGridLines = False
    ws.page_setup.horizontalCentered = True
    ws.oddFooter.left.text = "Northline Consumer Products  |  fictional sample  |  $000s"
    ws.oddFooter.right.text = "Page &P of &N"


def shade_section(ws, row, until=16):
    ws.cell(row, 2).fill = section_fill
    ws.cell(row, 2).font = section_font
    for c in range(3, until):
        ws.cell(row, c).fill = section_fill


def col(i):
    return get_column_letter(3 + i)


def mcol(c):
    return get_column_letter(c)


wb = Workbook()
wb.properties.creator = "Sai Siri Bandaru"
wb.properties.title = "Northline Driver-Based Forecast"

# ========== 00_Cover ==========
ws = wb.active
ws.title = "00_Cover"
ws.sheet_properties.tabColor = "1F4E79"
set_col_widths(ws, [4, 96])
landscape(ws, fit_height=1)

ws["B2"] = "Northline Consumer Products"
ws["B2"].font = title_font
ws["B3"] = "Driver-based P&L forecast  ·  volume × price  ·  COGS drivers  ·  OpEx by cost center"
ws["B3"].font = section_font
ws["B4"] = "FY2026 monthly  ·  plan vs reforecast  ·  fictional CPG company  ·  $000s"
ws["B4"].font = italic_grey

ws["B6"] = "What this file is"
ws["B6"].font = bold
ws["B7"] = (
    "A twelve-month driver-based P&L for a fictional mid-size CPG company. "
    "Revenue is volume × net price by product line. COGS is volume × material / labor / overhead per case. "
    "OpEx is built by cost center (rate + monthly fixed). Plan and latest-view forecast live on Assumptions; "
    "the P&L and the plan-vs-forecast bridge are formulas. A scenario toggle (Base / Upside / Downside) "
    "scales forecast volume, price, unit cost, and discretionary OpEx."
)
ws["B7"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[7].height = 62

ws["B9"] = "How to use"
ws["B9"].font = bold
ws["B10"] = "1. Edit yellow cells on 01_Assumptions (scenario, prices, unit costs, volumes, OpEx rates and monthly fixed)."
ws["B11"] = "2. Read volume × price and mix on 02_Revenue_Build (plan, scenario-applied forecast, variance)."
ws["B12"] = "3. Read unit-cost COGS and gross margin on 03_COGS_Build."
ws["B13"] = "4. Read cost-center OpEx (selling, marketing, distribution, G&A, R&D) on 04_OpEx."
ws["B14"] = "5. Read the 12-month P&L on 05_P&L_Forecast. Walk plan → forecast on 06_Bridge (volume / mix / price / cost / OpEx)."

ws["B16"] = "File conventions"
ws["B16"].font = bold
ws["B17"] = "Yellow cells with blue font = inputs. Black font = formulas. Green cells = key outputs."
ws["B17"].fill = yellow
ws["B17"].font = Font(name="Calibri", size=11, color="0000FF", bold=True)
ws["B18"] = "Forecast column is the latest view. Scenario factors scale that view (not the locked plan)."
ws["B19"] = "All figures are fictional. There is no employer data in this file. Dollar figures in $000s. Volume in 000 cases."

ws["B21"] = "Portfolio"
ws["B21"].font = bold
ws["B22"] = "Sai Siri Bandaru — Financial Analyst | FP&A | forecasting, variance analysis, Excel"
ws["B23"] = "https://github.com/saisiri-bandaru"
ws["B23"].font = link_font

# ========== 01_Assumptions ==========
ws = wb.create_sheet(ASSUMP)
ws.sheet_properties.tabColor = "F7C948"
set_col_widths(ws, [4, 36, 14, 14, 14, 14, 14, 14, 16] + [11] * 12)
landscape(ws)
ws.freeze_panes = "C6"

ws["B2"] = "Assumptions"
ws["B2"].font = title_font
ws["B3"] = "Yellow + blue = inputs. Scenario, prices, unit costs, volumes, and OpEx rates drive every other tab."
ws["B3"].font = italic_grey

ws["B5"] = "Control panel"
shade_section(ws, 5, until=6)
ws["B6"] = "Scenario (1 = Base, 2 = Upside, 3 = Downside)"
ws["C6"] = 1
style_input(ws["C6"])
ws["C6"].number_format = "0"
ws["D6"] = '=INDEX({"Base","Upside","Downside"},1,C6)'
style_formula(ws["D6"], key=True)
ws["E6"] = "Scales the latest-view forecast only. Plan stays put."
ws["E6"].font = small_grey

ws["B7"] = "Volume factor (on forecast volume)"
ws["C7"] = "=INDEX($D$15:$D$17,$C$6)"
style_formula(ws["C7"])
ws["C7"].number_format = "0.000"
ws["B8"] = "Price factor (on forecast net $/case)"
ws["C8"] = "=INDEX($E$15:$E$17,$C$6)"
style_formula(ws["C8"])
ws["C8"].number_format = "0.000"
ws["B9"] = "Unit-cost factor (on forecast CPU)"
ws["C9"] = "=INDEX($F$15:$F$17,$C$6)"
style_formula(ws["C9"])
ws["C9"].number_format = "0.000"
ws["B10"] = "Discretionary OpEx factor (marketing + R&D fixed)"
ws["C10"] = "=INDEX($G$15:$G$17,$C$6)"
style_formula(ws["C10"])
ws["C10"].number_format = "0.000"

ws["B12"] = "Scenario table (inputs)"
shade_section(ws, 12, until=8)
for i, h in enumerate(["#", "Name", "Volume", "Price", "Unit cost", "Disc. OpEx"]):
    style_header_cell(ws.cell(13, 2 + i), h)
scenarios = [
    (1, "Base", 1.00, 1.00, 1.00, 1.00),
    (2, "Upside", 1.045, 1.012, 0.985, 1.030),
    (3, "Downside", 0.955, 0.988, 1.025, 0.970),
]
for i, rowv in enumerate(scenarios):
    r = 15 + i
    ws.cell(r, 2, rowv[0]).border = thin
    ws.cell(r, 3, rowv[1]).border = thin
    for j, v in enumerate(rowv[2:]):
        cell = ws.cell(r, 4 + j, v)
        style_input(cell)
        cell.number_format = "0.000"

dv = DataValidation(type="list", formula1="1,2,3", allow_blank=False)
dv.error = "Enter 1, 2, or 3"
dv.errorTitle = "Scenario"
dv.prompt = "1 Base / 2 Upside / 3 Downside"
dv.promptTitle = "Scenario"
ws.add_data_validation(dv)
dv.add(ws["C6"])

ws["B19"] = "Other drivers"
shade_section(ws, 19, until=5)
ws["B20"] = "Tax rate"
ws["C20"] = 0.25
style_input(ws["C20"])
ws["C20"].number_format = pct
ws["B21"] = "D&A % of sales"
ws["C21"] = 0.034
style_input(ws["C21"])
ws["C21"].number_format = pct
ws["B22"] = "Cost inflation (forecast vs plan CPU)"
ws["C22"] = 0.018
style_input(ws["C22"])
ws["C22"].number_format = pct
ws["D22"] = "Applied before the scenario unit-cost factor. 1.8% commodity / conversion inflation in the sample."
ws["D22"].font = small_grey

ws["B24"] = "Product economics"
shade_section(ws, 24, until=10)
econ_h = [
    "Product", "Plan net $/case", "Fcst net $/case", "Material $/case",
    "Labor $/case", "Overhead $/case", "Plan CPU", "Fcst CPU",
]
for i, h in enumerate(econ_h):
    style_header_cell(ws.cell(25, 2 + i), h)
for i, name in enumerate(PRODUCTS):
    r = 26 + i
    ws.cell(r, 2, name).font = bold
    ws.cell(r, 2).border = thin
    for col_i, val, fmt in [
        (3, PLAN_PRICE[i], ppu),
        (4, FCST_PRICE[i], ppu),
        (5, MAT[i], ppu),
        (6, LAB[i], ppu),
        (7, OH[i], ppu),
    ]:
        cell = ws.cell(r, col_i, val)
        style_input(cell)
        cell.number_format = fmt
    cell = ws.cell(r, 8, f"=E{r}+F{r}+G{r}")
    style_formula(cell)
    cell.number_format = ppu
    cell = ws.cell(r, 9, f"=H{r}*(1+$C$22)*$C$9")
    style_formula(cell, key=True)
    cell.number_format = ppu

ws["B31"] = "OpEx rate drivers"
shade_section(ws, 31, until=6)
style_header_cell(ws.cell(32, 2), "Driver")
style_header_cell(ws.cell(32, 3), "Plan")
style_header_cell(ws.cell(32, 4), "Forecast")
ws["B33"] = "Selling commission % of sales"
ws["C33"] = 0.028
ws["D33"] = 0.028
ws["B34"] = "Marketing % of sales"
ws["C34"] = 0.046
ws["D34"] = 0.049
ws["B35"] = "Distribution $ per case"
ws["C35"] = 0.68
ws["D35"] = 0.66
for r in (33, 34, 35):
    style_input(ws.cell(r, 3))
    style_input(ws.cell(r, 4))
    fmt = ppu if r == 35 else pct
    ws.cell(r, 3).number_format = fmt
    ws.cell(r, 4).number_format = fmt

ws["B37"] = "Plan monthly fixed OpEx ($000s)"
shade_section(ws, 37)
month_headers(ws, 38)
plan_fix = [
    ("Selling", PLAN_SELL_FIX),
    ("Marketing (campaigns)", PLAN_MKT_FIX),
    ("Distribution", PLAN_DIST_FIX),
    ("G&A", PLAN_GA_FIX),
    ("R&D", PLAN_RD_FIX),
]
for i, (name, vals) in enumerate(plan_fix):
    r = 39 + i
    ws.cell(r, 2, name).border = thin
    ws.cell(r, 2).font = bold
    for m, v in enumerate(vals):
        cell = ws.cell(r, 3 + m, v)
        style_input(cell)
        cell.number_format = money
    cell = ws.cell(r, 15, f"=SUM(C{r}:N{r})")
    style_formula(cell, key=True)
    cell.number_format = money
ws["B44"] = "Total plan fixed OpEx"
ws["B44"].font = bold
ws["B44"].border = thin
for m in range(13):
    cell = ws.cell(44, 3 + m, f"=SUM({col(m)}39:{col(m)}43)" if m < 12 else "=SUM(C44:N44)")
    if m == 12:
        cell.value = "=SUM(C44:N44)"
        # wait C44:N44 would be circular. Use SUM of FY column of rows 39-43
        cell.value = "=SUM(O39:O43)"
    else:
        cell.value = f"=SUM({col(m)}39:{col(m)}43)"
    style_formula(cell, key=True)
    cell.number_format = money

ws["B46"] = "Forecast monthly fixed OpEx ($000s) — before discretionary factor on marketing + R&D"
shade_section(ws, 46)
month_headers(ws, 47)
fcst_fix = [
    ("Selling", FCST_SELL_FIX),
    ("Marketing (campaigns)", FCST_MKT_FIX),
    ("Distribution", FCST_DIST_FIX),
    ("G&A", FCST_GA_FIX),
    ("R&D", FCST_RD_FIX),
]
for i, (name, vals) in enumerate(fcst_fix):
    r = 48 + i
    ws.cell(r, 2, name).border = thin
    ws.cell(r, 2).font = bold
    for m, v in enumerate(vals):
        cell = ws.cell(r, 3 + m, v)
        style_input(cell)
        cell.number_format = money
    cell = ws.cell(r, 15, f"=SUM(C{r}:N{r})")
    style_formula(cell, key=True)
    cell.number_format = money
ws["B53"] = "Total forecast fixed OpEx (raw)"
ws["B53"].font = bold
ws["B53"].border = thin
for m in range(12):
    cell = ws.cell(53, 3 + m, f"=SUM({col(m)}48:{col(m)}52)")
    style_formula(cell, key=True)
    cell.number_format = money
ws["O53"] = "=SUM(O48:O52)"
style_formula(ws["O53"], key=True)
ws["O53"].number_format = money

ws["B55"] = "Plan volume (000 cases)"
shade_section(ws, 55)
month_headers(ws, 56)
for i, name in enumerate(PRODUCTS):
    r = 57 + i
    ws.cell(r, 2, name).border = thin
    ws.cell(r, 2).font = bold
    for m, v in enumerate(PLAN_VOL[i]):
        cell = ws.cell(r, 3 + m, v)
        style_input(cell)
        cell.number_format = int_fmt
    cell = ws.cell(r, 15, f"=SUM(C{r}:N{r})")
    style_formula(cell, key=True)
    cell.number_format = int_fmt
ws["B61"] = "Total plan volume"
ws["B61"].font = bold
ws["B61"].border = thin
for m in range(12):
    cell = ws.cell(61, 3 + m, f"=SUM({col(m)}57:{col(m)}60)")
    style_formula(cell, key=True)
    cell.number_format = int_fmt
ws["O61"] = "=SUM(O57:O60)"
style_formula(ws["O61"], key=True)
ws["O61"].number_format = int_fmt

ws["B63"] = "Forecast volume — latest view (000 cases), before scenario factor"
shade_section(ws, 63)
month_headers(ws, 64)
for i, name in enumerate(PRODUCTS):
    r = 65 + i
    ws.cell(r, 2, name).border = thin
    ws.cell(r, 2).font = bold
    for m, v in enumerate(FCST_VOL[i]):
        cell = ws.cell(r, 3 + m, v)
        style_input(cell)
        cell.number_format = int_fmt
    cell = ws.cell(r, 15, f"=SUM(C{r}:N{r})")
    style_formula(cell, key=True)
    cell.number_format = int_fmt
ws["B69"] = "Total forecast volume (raw)"
ws["B69"].font = bold
ws["B69"].border = thin
for m in range(12):
    cell = ws.cell(69, 3 + m, f"=SUM({col(m)}65:{col(m)}68)")
    style_formula(cell, key=True)
    cell.number_format = int_fmt
ws["O69"] = "=SUM(O65:O68)"
style_formula(ws["O69"], key=True)
ws["O69"].number_format = int_fmt

ws["B71"] = (
    "Plan is the locked annual plan. Forecast is the latest view (reforecast). "
    "Scenario factors on C7:C10 scale forecast volume, price, unit cost, and discretionary OpEx. "
    "Volume in 000 cases; prices and CPU in $ per case; OpEx and P&L in $000s."
)
ws["B71"].font = small_grey
ws["B71"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[71].height = 36

# ========== 02_Revenue_Build ==========
ws = wb.create_sheet(REV)
ws.sheet_properties.tabColor = "5B9BD5"
set_col_widths(ws, [4, 32] + [11] * 13 + [14])
landscape(ws)
ws.freeze_panes = "C8"

ws["B2"] = "Revenue build — volume × net price"
ws["B2"].font = title_font
ws["B3"] = "Plan uses plan volume × plan price. Forecast uses latest-view volume × volume factor × forecast price × price factor."
ws["B3"].font = italic_grey
ws["B4"] = "Scenario"
ws["C4"] = f"='{ASSUMP}'!D6"
style_formula(ws["C4"], key=True)

# PLAN
ws["B6"] = "PLAN"
shade_section(ws, 6)
month_headers(ws, 7)

# Plan volume 8-11, price 13-16, revenue 18-21
for i, name in enumerate(PRODUCTS):
    r = 8 + i
    ar = 57 + i
    ws.cell(r, 2, name + " volume").border = thin
    for m in range(12):
        cell = ws.cell(r, 3 + m, f"='{ASSUMP}'!{col(m)}{ar}")
        style_formula(cell)
        cell.number_format = int_fmt
    cell = ws.cell(r, 15, f"=SUM(C{r}:N{r})")
    style_formula(cell, key=True)
    cell.number_format = int_fmt

ws["B12"] = "Plan total volume"
ws["B12"].font = bold
ws["B12"].border = thin
for m in range(12):
    cell = ws.cell(12, 3 + m, f"=SUM({col(m)}8:{col(m)}11)")
    style_formula(cell, key=True)
    cell.number_format = int_fmt
ws["O12"] = "=SUM(O8:O11)"
style_formula(ws["O12"], key=True)
ws["O12"].number_format = int_fmt

for i, name in enumerate(PRODUCTS):
    r = 14 + i
    pr = 26 + i
    ws.cell(r, 2, name + " net $/case").border = thin
    for m in range(12):
        cell = ws.cell(r, 3 + m, f"='{ASSUMP}'!$C${pr}")
        style_formula(cell)
        cell.number_format = ppu
    cell = ws.cell(r, 15, f"=AVERAGE(C{r}:N{r})")
    style_formula(cell)
    cell.number_format = ppu

for i, name in enumerate(PRODUCTS):
    r = 19 + i
    ws.cell(r, 2, name + " revenue").border = thin
    ws.cell(r, 2).font = bold
    vr, pr = 8 + i, 14 + i
    for m in range(12):
        cell = ws.cell(r, 3 + m, f"={col(m)}{vr}*{col(m)}{pr}")
        style_formula(cell)
        cell.number_format = money
    cell = ws.cell(r, 15, f"=SUM(C{r}:N{r})")
    style_formula(cell, key=True)
    cell.number_format = money

ws["B23"] = "Plan net revenue"
ws["B23"].font = bold
ws["B23"].border = thin
for m in range(12):
    cell = ws.cell(23, 3 + m, f"=SUM({col(m)}19:{col(m)}22)")
    style_formula(cell, key=True)
    cell.number_format = money
ws["O23"] = "=SUM(O19:O22)"
style_formula(ws["O23"], key=True)
ws["O23"].number_format = money

ws["B24"] = "Plan ASP $/case"
for m in range(12):
    cell = ws.cell(24, 3 + m, f"=IF({col(m)}12=0,0,{col(m)}23/{col(m)}12)")
    style_formula(cell)
    cell.number_format = ppu
ws["O24"] = "=IF(O12=0,0,O23/O12)"
style_formula(ws["O24"])
ws["O24"].number_format = ppu

ws["B25"] = "Plan mix % of volume"
for i, name in enumerate(PRODUCTS):
    r = 26 + i
    ws.cell(r, 2, name).border = thin
    for m in range(13):
        cl = col(m) if m < 12 else "O"
        cell = ws.cell(r, 3 + m if m < 12 else 15, f"=IF({cl}12=0,0,{cl}{8+i}/{cl}12)")
        style_formula(cell)
        cell.number_format = pct

# FORECAST
ws["B31"] = "FORECAST (scenario applied)"
shade_section(ws, 31)
month_headers(ws, 32)

for i, name in enumerate(PRODUCTS):
    r = 33 + i
    ar = 65 + i
    ws.cell(r, 2, name + " volume").border = thin
    for m in range(12):
        cell = ws.cell(r, 3 + m, f"='{ASSUMP}'!{col(m)}{ar}*'{ASSUMP}'!$C$7")
        style_formula(cell)
        cell.number_format = num
    cell = ws.cell(r, 15, f"=SUM(C{r}:N{r})")
    style_formula(cell, key=True)
    cell.number_format = num

ws["B37"] = "Forecast total volume"
ws["B37"].font = bold
ws["B37"].border = thin
for m in range(12):
    cell = ws.cell(37, 3 + m, f"=SUM({col(m)}33:{col(m)}36)")
    style_formula(cell, key=True)
    cell.number_format = num
ws["O37"] = "=SUM(O33:O36)"
style_formula(ws["O37"], key=True)
ws["O37"].number_format = num

for i, name in enumerate(PRODUCTS):
    r = 39 + i
    pr = 26 + i
    ws.cell(r, 2, name + " net $/case").border = thin
    for m in range(12):
        cell = ws.cell(r, 3 + m, f"='{ASSUMP}'!$D${pr}*'{ASSUMP}'!$C$8")
        style_formula(cell)
        cell.number_format = ppu
    cell = ws.cell(r, 15, f"=AVERAGE(C{r}:N{r})")
    style_formula(cell)
    cell.number_format = ppu

for i, name in enumerate(PRODUCTS):
    r = 44 + i
    ws.cell(r, 2, name + " revenue").border = thin
    ws.cell(r, 2).font = bold
    vr, pr = 33 + i, 39 + i
    for m in range(12):
        cell = ws.cell(r, 3 + m, f"={col(m)}{vr}*{col(m)}{pr}")
        style_formula(cell)
        cell.number_format = money
    cell = ws.cell(r, 15, f"=SUM(C{r}:N{r})")
    style_formula(cell, key=True)
    cell.number_format = money

ws["B48"] = "Forecast net revenue"
ws["B48"].font = bold
ws["B48"].border = thin
for m in range(12):
    cell = ws.cell(48, 3 + m, f"=SUM({col(m)}44:{col(m)}47)")
    style_formula(cell, key=True)
    cell.number_format = money
ws["O48"] = "=SUM(O44:O47)"
style_formula(ws["O48"], key=True)
ws["O48"].number_format = money

ws["B49"] = "Forecast ASP $/case"
for m in range(12):
    cell = ws.cell(49, 3 + m, f"=IF({col(m)}37=0,0,{col(m)}48/{col(m)}37)")
    style_formula(cell)
    cell.number_format = ppu
ws["O49"] = "=IF(O37=0,0,O48/O37)"
style_formula(ws["O49"])
ws["O49"].number_format = ppu

ws["B50"] = "Forecast mix % of volume"
for i, name in enumerate(PRODUCTS):
    r = 51 + i
    ws.cell(r, 2, name).border = thin
    for m in range(13):
        cl = col(m) if m < 12 else "O"
        cell = ws.cell(r, 3 + m if m < 12 else 15, f"=IF({cl}37=0,0,{cl}{33+i}/{cl}37)")
        style_formula(cell)
        cell.number_format = pct

# VARIANCE
ws["B56"] = "VARIANCE (forecast − plan)"
shade_section(ws, 56)
month_headers(ws, 57)
ws["B58"] = "Volume var (000 cases)"
for m in range(13):
    cl = col(m) if m < 12 else "O"
    cidx = 3 + m if m < 12 else 15
    cell = ws.cell(58, cidx, f"={cl}37-{cl}12")
    style_formula(cell)
    cell.number_format = num
ws["B59"] = "Revenue var ($000s)"
ws["B59"].font = bold
for m in range(13):
    cl = col(m) if m < 12 else "O"
    cidx = 3 + m if m < 12 else 15
    cell = ws.cell(59, cidx, f"={cl}48-{cl}23")
    style_formula(cell, key=True)
    cell.number_format = money
ws["B60"] = "Revenue var %"
for m in range(13):
    cl = col(m) if m < 12 else "O"
    cidx = 3 + m if m < 12 else 15
    cell = ws.cell(60, cidx, f"=IF({cl}23=0,0,{cl}59/{cl}23)")
    style_formula(cell)
    cell.number_format = pct

ws["B62"] = "Volume is 000 cases. Revenue is $000s (= 000 cases × $ per case). Mix % is share of total cases."
ws["B62"].font = small_grey

# ========== 03_COGS_Build ==========
ws = wb.create_sheet(COGS)
ws.sheet_properties.tabColor = "ED7D31"
set_col_widths(ws, [4, 32] + [11] * 13 + [14])
landscape(ws)
ws.freeze_panes = "C8"

ws["B2"] = "COGS build — volume × cost per unit"
ws["B2"].font = title_font
ws["B3"] = "Plan CPU = material + labor + overhead. Forecast CPU = plan CPU × (1 + inflation) × unit-cost factor."
ws["B3"].font = italic_grey
ws["B4"] = "Scenario"
ws["C4"] = f"='{ASSUMP}'!D6"
style_formula(ws["C4"], key=True)

ws["B6"] = "PLAN"
shade_section(ws, 6)
month_headers(ws, 7)

for i, name in enumerate(PRODUCTS):
    r = 8 + i
    pr = 26 + i
    ws.cell(r, 2, name + " CPU").border = thin
    for m in range(12):
        cell = ws.cell(r, 3 + m, f"='{ASSUMP}'!$H${pr}")
        style_formula(cell)
        cell.number_format = ppu
    cell = ws.cell(r, 15, f"=AVERAGE(C{r}:N{r})")
    style_formula(cell)
    cell.number_format = ppu

for i, name in enumerate(PRODUCTS):
    r = 13 + i
    ws.cell(r, 2, name + " COGS").border = thin
    ws.cell(r, 2).font = bold
    for m in range(12):
        cell = ws.cell(r, 3 + m, f"='{REV}'!{col(m)}{8+i}*{col(m)}{8+i}")
        style_formula(cell)
        cell.number_format = money
    cell = ws.cell(r, 15, f"=SUM(C{r}:N{r})")
    style_formula(cell, key=True)
    cell.number_format = money

ws["B17"] = "Plan COGS"
ws["B17"].font = bold
ws["B17"].border = thin
for m in range(12):
    cell = ws.cell(17, 3 + m, f"=SUM({col(m)}13:{col(m)}16)")
    style_formula(cell, key=True)
    cell.number_format = money
ws["O17"] = "=SUM(O13:O16)"
style_formula(ws["O17"], key=True)
ws["O17"].number_format = money

ws["B18"] = "Plan gross profit"
for m in range(13):
    cl = col(m) if m < 12 else "O"
    cidx = 3 + m if m < 12 else 15
    cell = ws.cell(18, cidx, f"='{REV}'!{cl}23-{cl}17")
    style_formula(cell, key=True)
    cell.number_format = money
ws["B19"] = "Plan gross margin %"
for m in range(13):
    cl = col(m) if m < 12 else "O"
    cidx = 3 + m if m < 12 else 15
    cell = ws.cell(19, cidx, f"=IF('{REV}'!{cl}23=0,0,{cl}18/'{REV}'!{cl}23)")
    style_formula(cell)
    cell.number_format = pct

ws["B21"] = "FORECAST (scenario applied)"
shade_section(ws, 21)
month_headers(ws, 22)

for i, name in enumerate(PRODUCTS):
    r = 23 + i
    pr = 26 + i
    ws.cell(r, 2, name + " CPU").border = thin
    for m in range(12):
        cell = ws.cell(r, 3 + m, f"='{ASSUMP}'!$I${pr}")
        style_formula(cell)
        cell.number_format = ppu
    cell = ws.cell(r, 15, f"=AVERAGE(C{r}:N{r})")
    style_formula(cell)
    cell.number_format = ppu

for i, name in enumerate(PRODUCTS):
    r = 28 + i
    ws.cell(r, 2, name + " COGS").border = thin
    ws.cell(r, 2).font = bold
    for m in range(12):
        cell = ws.cell(r, 3 + m, f"='{REV}'!{col(m)}{33+i}*{col(m)}{23+i}")
        style_formula(cell)
        cell.number_format = money
    cell = ws.cell(r, 15, f"=SUM(C{r}:N{r})")
    style_formula(cell, key=True)
    cell.number_format = money

ws["B32"] = "Forecast COGS"
ws["B32"].font = bold
ws["B32"].border = thin
for m in range(12):
    cell = ws.cell(32, 3 + m, f"=SUM({col(m)}28:{col(m)}31)")
    style_formula(cell, key=True)
    cell.number_format = money
ws["O32"] = "=SUM(O28:O31)"
style_formula(ws["O32"], key=True)
ws["O32"].number_format = money

ws["B33"] = "Forecast gross profit"
for m in range(13):
    cl = col(m) if m < 12 else "O"
    cidx = 3 + m if m < 12 else 15
    cell = ws.cell(33, cidx, f"='{REV}'!{cl}48-{cl}32")
    style_formula(cell, key=True)
    cell.number_format = money
ws["B34"] = "Forecast gross margin %"
for m in range(13):
    cl = col(m) if m < 12 else "O"
    cidx = 3 + m if m < 12 else 15
    cell = ws.cell(34, cidx, f"=IF('{REV}'!{cl}48=0,0,{cl}33/'{REV}'!{cl}48)")
    style_formula(cell)
    cell.number_format = pct

ws["B36"] = "VARIANCE (forecast − plan)"
shade_section(ws, 36)
month_headers(ws, 37)
ws["B38"] = "COGS var"
for m in range(13):
    cl = col(m) if m < 12 else "O"
    cidx = 3 + m if m < 12 else 15
    cell = ws.cell(38, cidx, f"={cl}32-{cl}17")
    style_formula(cell)
    cell.number_format = money
ws["B39"] = "Gross profit var"
for m in range(13):
    cl = col(m) if m < 12 else "O"
    cidx = 3 + m if m < 12 else 15
    cell = ws.cell(39, cidx, f"={cl}33-{cl}18")
    style_formula(cell, key=True)
    cell.number_format = money
ws["B40"] = "Gross margin pts"
for m in range(13):
    cl = col(m) if m < 12 else "O"
    cidx = 3 + m if m < 12 else 15
    cell = ws.cell(40, cidx, f"={cl}34-{cl}19")
    style_formula(cell)
    cell.number_format = "0.0%"

ws["B42"] = "CPU is $ per case. COGS $000s = 000 cases × $ per case. Inflation and the scenario unit-cost factor both sit in forecast CPU on Assumptions (column I)."
ws["B42"].font = small_grey

# ========== 04_OpEx ==========
ws = wb.create_sheet(OPEX)
ws.sheet_properties.tabColor = "7030A0"
set_col_widths(ws, [4, 36] + [11] * 13 + [14])
landscape(ws)
ws.freeze_panes = "C8"

ws["B2"] = "OpEx by cost center"
ws["B2"].font = title_font
ws["B3"] = "Selling and marketing = % of sales + monthly fixed. Distribution = $ per case + monthly fixed. G&A and R&D are monthly fixed. Discretionary factor scales forecast marketing + R&D fixed."
ws["B3"].font = italic_grey
ws["B3"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[3].height = 32
ws["B4"] = "Scenario"
ws["C4"] = f"='{ASSUMP}'!D6"
style_formula(ws["C4"], key=True)

ws["B6"] = "PLAN"
shade_section(ws, 6)
month_headers(ws, 7)

# Selling = commission * plan rev + fixed
ws["B8"] = "Selling"
for m in range(12):
    cell = ws.cell(8, 3 + m, f"='{ASSUMP}'!$C$33*'{REV}'!{col(m)}23+'{ASSUMP}'!{col(m)}39")
    style_formula(cell)
    cell.number_format = money
ws["O8"] = "=SUM(C8:N8)"
style_formula(ws["O8"], key=True)
ws["O8"].number_format = money

ws["B9"] = "Marketing"
for m in range(12):
    cell = ws.cell(9, 3 + m, f"='{ASSUMP}'!$C$34*'{REV}'!{col(m)}23+'{ASSUMP}'!{col(m)}40")
    style_formula(cell)
    cell.number_format = money
ws["O9"] = "=SUM(C9:N9)"
style_formula(ws["O9"], key=True)
ws["O9"].number_format = money

ws["B10"] = "Distribution"
for m in range(12):
    cell = ws.cell(10, 3 + m, f"='{ASSUMP}'!$C$35*'{REV}'!{col(m)}12+'{ASSUMP}'!{col(m)}41")
    style_formula(cell)
    cell.number_format = money
ws["O10"] = "=SUM(C10:N10)"
style_formula(ws["O10"], key=True)
ws["O10"].number_format = money

ws["B11"] = "G&A"
for m in range(12):
    cell = ws.cell(11, 3 + m, f"='{ASSUMP}'!{col(m)}42")
    style_formula(cell)
    cell.number_format = money
ws["O11"] = "=SUM(C11:N11)"
style_formula(ws["O11"], key=True)
ws["O11"].number_format = money

ws["B12"] = "R&D"
for m in range(12):
    cell = ws.cell(12, 3 + m, f"='{ASSUMP}'!{col(m)}43")
    style_formula(cell)
    cell.number_format = money
ws["O12"] = "=SUM(C12:N12)"
style_formula(ws["O12"], key=True)
ws["O12"].number_format = money

for r in range(8, 13):
    ws.cell(r, 2).border = thin
    ws.cell(r, 2).font = bold

ws["B13"] = "Plan total OpEx"
ws["B13"].font = bold
ws["B13"].border = thin
for m in range(12):
    cell = ws.cell(13, 3 + m, f"=SUM({col(m)}8:{col(m)}12)")
    style_formula(cell, key=True)
    cell.number_format = money
ws["O13"] = "=SUM(O8:O12)"
style_formula(ws["O13"], key=True)
ws["O13"].number_format = money

ws["B14"] = "Plan OpEx % of sales"
for m in range(13):
    cl = col(m) if m < 12 else "O"
    cidx = 3 + m if m < 12 else 15
    cell = ws.cell(14, cidx, f"=IF('{REV}'!{cl}23=0,0,{cl}13/'{REV}'!{cl}23)")
    style_formula(cell)
    cell.number_format = pct

ws["B16"] = "FORECAST (scenario applied)"
shade_section(ws, 16)
month_headers(ws, 17)

# Selling: fcst commission * fcst rev + fcst fixed
ws["B18"] = "Selling"
for m in range(12):
    cell = ws.cell(18, 3 + m, f"='{ASSUMP}'!$D$33*'{REV}'!{col(m)}48+'{ASSUMP}'!{col(m)}48")
    style_formula(cell)
    cell.number_format = money
ws["O18"] = "=SUM(C18:N18)"
style_formula(ws["O18"], key=True)
ws["O18"].number_format = money

ws["B19"] = "Marketing"
for m in range(12):
    cell = ws.cell(19, 3 + m, f"='{ASSUMP}'!$D$34*'{REV}'!{col(m)}48+'{ASSUMP}'!{col(m)}49*'{ASSUMP}'!$C$10")
    style_formula(cell)
    cell.number_format = money
ws["O19"] = "=SUM(C19:N19)"
style_formula(ws["O19"], key=True)
ws["O19"].number_format = money

ws["B20"] = "Distribution"
for m in range(12):
    cell = ws.cell(20, 3 + m, f"='{ASSUMP}'!$D$35*'{REV}'!{col(m)}37+'{ASSUMP}'!{col(m)}50")
    style_formula(cell)
    cell.number_format = money
ws["O20"] = "=SUM(C20:N20)"
style_formula(ws["O20"], key=True)
ws["O20"].number_format = money

ws["B21"] = "G&A"
for m in range(12):
    cell = ws.cell(21, 3 + m, f"='{ASSUMP}'!{col(m)}51")
    style_formula(cell)
    cell.number_format = money
ws["O21"] = "=SUM(C21:N21)"
style_formula(ws["O21"], key=True)
ws["O21"].number_format = money

ws["B22"] = "R&D"
for m in range(12):
    cell = ws.cell(22, 3 + m, f"='{ASSUMP}'!{col(m)}52*'{ASSUMP}'!$C$10")
    style_formula(cell)
    cell.number_format = money
ws["O22"] = "=SUM(C22:N22)"
style_formula(ws["O22"], key=True)
ws["O22"].number_format = money

for r in range(18, 23):
    ws.cell(r, 2).border = thin
    ws.cell(r, 2).font = bold

ws["B23"] = "Forecast total OpEx"
ws["B23"].font = bold
ws["B23"].border = thin
for m in range(12):
    cell = ws.cell(23, 3 + m, f"=SUM({col(m)}18:{col(m)}22)")
    style_formula(cell, key=True)
    cell.number_format = money
ws["O23"] = "=SUM(O18:O22)"
style_formula(ws["O23"], key=True)
ws["O23"].number_format = money

ws["B24"] = "Forecast OpEx % of sales"
for m in range(13):
    cl = col(m) if m < 12 else "O"
    cidx = 3 + m if m < 12 else 15
    cell = ws.cell(24, cidx, f"=IF('{REV}'!{cl}48=0,0,{cl}23/'{REV}'!{cl}48)")
    style_formula(cell)
    cell.number_format = pct

ws["B26"] = "VARIANCE (forecast − plan)"
shade_section(ws, 26)
month_headers(ws, 27)
centers = ["Selling", "Marketing", "Distribution", "G&A", "R&D", "Total OpEx"]
for i, name in enumerate(centers):
    r = 28 + i
    pr, fr = (8 + i, 18 + i) if i < 5 else (13, 23)
    ws.cell(r, 2, name).border = thin
    if i == 5:
        ws.cell(r, 2).font = bold
    for m in range(13):
        cl = col(m) if m < 12 else "O"
        cidx = 3 + m if m < 12 else 15
        cell = ws.cell(r, cidx, f"={cl}{fr}-{cl}{pr}")
        style_formula(cell, key=(i == 5))
        cell.number_format = money

ws["B35"] = "Distribution $ per case × 000 cases = $000s. Discretionary factor (Assumptions C10) applies to forecast marketing campaigns and R&D only."
ws["B35"].font = small_grey

# ========== 05_P&L_Forecast ==========
ws = wb.create_sheet(PL)
ws.sheet_properties.tabColor = "70AD47"
set_col_widths(ws, [4, 28] + [11] * 12 + [13, 13, 12, 11])
landscape(ws)
ws.freeze_panes = "C8"

ws["B2"] = "P&L forecast — 12 months"
ws["B2"].font = title_font
ws["B3"] = "Main grid is the latest-view forecast (scenario applied). Right-hand columns compare FY forecast vs FY plan."
ws["B3"].font = italic_grey
ws["B4"] = "Scenario"
ws["C4"] = f"='{ASSUMP}'!D6"
style_formula(ws["C4"], key=True)

ws["B6"] = "Forecast P&L ($000s)"
shade_section(ws, 6, until=19)
month_headers(ws, 7, extra=["FY Fcst", "FY Plan", "Var", "Var %"])

# Row map
# 8 Net revenue
# 9 COGS
# 10 Gross profit
# 11 Gross margin %
# 12 Selling
# 13 Marketing
# 14 Distribution
# 15 G&A
# 16 R&D
# 17 Total OpEx
# 18 EBITDA
# 19 EBITDA %
# 20 D&A
# 21 EBIT
# 22 Tax
# 23 Net income
# 24 NI %

pl_rows = [
    (8, "Net revenue", f"'{REV}'", 48, f"'{REV}'", 23, money, True),
    (9, "COGS", f"'{COGS}'", 32, f"'{COGS}'", 17, money, False),
    (10, "Gross profit", None, None, None, None, money, True),
    (11, "Gross margin %", None, None, None, None, pct, False),
    (12, "Selling", f"'{OPEX}'", 18, f"'{OPEX}'", 8, money, False),
    (13, "Marketing", f"'{OPEX}'", 19, f"'{OPEX}'", 9, money, False),
    (14, "Distribution", f"'{OPEX}'", 20, f"'{OPEX}'", 10, money, False),
    (15, "G&A", f"'{OPEX}'", 21, f"'{OPEX}'", 11, money, False),
    (16, "R&D", f"'{OPEX}'", 22, f"'{OPEX}'", 12, money, False),
    (17, "Total OpEx", f"'{OPEX}'", 23, f"'{OPEX}'", 13, money, True),
    (18, "EBITDA", None, None, None, None, money, True),
    (19, "EBITDA %", None, None, None, None, pct, False),
    (20, "D&A", None, None, None, None, money, False),
    (21, "EBIT", None, None, None, None, money, True),
    (22, "Tax", None, None, None, None, money, False),
    (23, "Net income", None, None, None, None, money, True),
    (24, "Net income %", None, None, None, None, pct, False),
]

for r, name, fsheet, frow, psheet, prow, fmt, key in pl_rows:
    ws.cell(r, 2, name).border = thin
    ws.cell(r, 2).font = bold if key else black
    if r in (11, 19, 24):
        ws.cell(r, 2).fill = light_gray

    for m in range(12):
        cl = col(m)
        if r == 8:
            f = f"={fsheet}!{cl}{frow}"
        elif r == 9:
            f = f"={fsheet}!{cl}{frow}"
        elif r == 10:
            f = f"={cl}8-{cl}9"
        elif r == 11:
            f = f"=IF({cl}8=0,0,{cl}10/{cl}8)"
        elif r in (12, 13, 14, 15, 16, 17):
            f = f"={fsheet}!{cl}{frow}"
        elif r == 18:
            f = f"={cl}10-{cl}17"
        elif r == 19:
            f = f"=IF({cl}8=0,0,{cl}18/{cl}8)"
        elif r == 20:
            f = f"='{ASSUMP}'!$C$21*{cl}8"
        elif r == 21:
            f = f"={cl}18-{cl}20"
        elif r == 22:
            f = f"='{ASSUMP}'!$C$20*{cl}21"
        elif r == 23:
            f = f"={cl}21-{cl}22"
        else:
            f = f"=IF({cl}8=0,0,{cl}23/{cl}8)"
        cell = ws.cell(r, 3 + m, f)
        style_formula(cell, key=key)
        cell.number_format = fmt
        if r in (11, 19, 24):
            cell.fill = light_gray
            if key:
                cell.fill = green_fill

    # FY Fcst = SUM of months (or average for %)
    if fmt == pct:
        if r == 11:
            fyf = "=IF(O8=0,0,O10/O8)"
            fyp = "=IF(P8=0,0,P10/P8)"
        elif r == 19:
            fyf = "=IF(O8=0,0,O18/O8)"
            fyp = "=IF(P8=0,0,P18/P8)"
        else:
            fyf = "=IF(O8=0,0,O23/O8)"
            fyp = "=IF(P8=0,0,P23/P8)"
        cell = ws.cell(r, 15, fyf)
        style_formula(cell, key=key)
        cell.number_format = fmt
        cell = ws.cell(r, 16, fyp)
        style_formula(cell)
        cell.number_format = fmt
        cell = ws.cell(r, 17, f"=O{r}-P{r}")
        style_formula(cell)
        cell.number_format = "0.0%"
        cell = ws.cell(r, 18, '="n/a"')
        style_formula(cell)
    else:
        cell = ws.cell(r, 15, f"=SUM(C{r}:N{r})")
        style_formula(cell, key=True)
        cell.number_format = money
        # FY Plan
        if r == 8:
            fyp = f"={psheet}!O{prow}"
        elif r == 9:
            fyp = f"={psheet}!O{prow}"
        elif r == 10:
            fyp = "=P8-P9"
        elif r in (12, 13, 14, 15, 16, 17):
            fyp = f"={psheet}!O{prow}"
        elif r == 18:
            fyp = "=P10-P17"
        elif r == 20:
            fyp = f"='{ASSUMP}'!$C$21*P8"
        elif r == 21:
            fyp = "=P18-P20"
        elif r == 22:
            fyp = f"='{ASSUMP}'!$C$20*P21"
        elif r == 23:
            fyp = "=P21-P22"
        else:
            fyp = "=0"
        cell = ws.cell(r, 16, fyp)
        style_formula(cell)
        cell.number_format = money
        cell = ws.cell(r, 17, f"=O{r}-P{r}")
        style_formula(cell, key=key)
        cell.number_format = money
        cell = ws.cell(r, 18, f"=IF(P{r}=0,0,Q{r}/P{r})")
        style_formula(cell)
        cell.number_format = pct

# Plan monthly block for the chart
ws["B26"] = "Plan P&L ($000s) — for the chart"
shade_section(ws, 26)
month_headers(ws, 27)
plan_chart_rows = [
    (28, "Plan net revenue", f"='{REV}'!{{c}}23"),
    (29, "Plan EBITDA", f"='{REV}'!{{c}}23-'{COGS}'!{{c}}17-'{OPEX}'!{{c}}13"),
]
ws["B28"] = "Plan net revenue"
ws["B29"] = "Plan EBITDA"
ws["B28"].border = thin
ws["B29"].border = thin
for m in range(12):
    cl = col(m)
    cell = ws.cell(28, 3 + m, f"='{REV}'!{cl}23")
    style_formula(cell)
    cell.number_format = money
    cell = ws.cell(29, 3 + m, f"='{REV}'!{cl}23-'{COGS}'!{cl}17-'{OPEX}'!{cl}13")
    style_formula(cell)
    cell.number_format = money
ws["O28"] = "=SUM(C28:N28)"
style_formula(ws["O28"])
ws["O28"].number_format = money
ws["O29"] = "=SUM(C29:N29)"
style_formula(ws["O29"])
ws["O29"].number_format = money

ws["B31"] = "Chart data (hidden-style, used by the line chart)"
ws["B32"] = "Forecast revenue"
ws["B33"] = "Plan revenue"
ws["B34"] = "Forecast EBITDA"
ws["B35"] = "Plan EBITDA"
for m in range(12):
    cl = col(m)
    style_header_cell(ws.cell(31, 3 + m), MONTHS[m])
    cell = ws.cell(32, 3 + m, f"={cl}8")
    style_formula(cell)
    cell.number_format = money
    cell = ws.cell(33, 3 + m, f"={cl}28")
    style_formula(cell)
    cell.number_format = money
    cell = ws.cell(34, 3 + m, f"={cl}18")
    style_formula(cell)
    cell.number_format = money
    cell = ws.cell(35, 3 + m, f"={cl}29")
    style_formula(cell)
    cell.number_format = money

chart1 = LineChart()
chart1.title = "Revenue: forecast vs plan ($000s)"
chart1.style = 10
chart1.y_axis.title = "$000s"
chart1.height = 7
chart1.width = 15
chart1.legend.position = "b"
data = Reference(ws, min_col=2, min_row=32, max_col=14, max_row=33)
cats = Reference(ws, min_col=3, min_row=31, max_col=14)
chart1.add_data(data, from_rows=True, titles_from_data=True)
chart1.set_categories(cats)
ws.add_chart(chart1, "B37")

chart2 = LineChart()
chart2.title = "EBITDA: forecast vs plan ($000s)"
chart2.style = 12
chart2.y_axis.title = "$000s"
chart2.height = 7
chart2.width = 15
chart2.legend.position = "b"
data2 = Reference(ws, min_col=2, min_row=34, max_col=14, max_row=35)
chart2.add_data(data2, from_rows=True, titles_from_data=True)
chart2.set_categories(cats)
ws.add_chart(chart2, "I37")

ws["B52"] = "Switch the scenario on Assumptions to watch FY var and the charts move. Green cells on this tab are the conversation numbers."
ws["B52"].font = small_grey

# Conditional formatting on FY var % for revenue / EBITDA / NI
ws.conditional_formatting.add(
    "R8:R8",
    CellIsRule(operator="greaterThan", formula=["0"], fill=green_fill),
)
ws.conditional_formatting.add(
    "R8:R8",
    CellIsRule(operator="lessThan", formula=["0"], fill=red_fill),
)
ws.conditional_formatting.add(
    "R18:R18",
    CellIsRule(operator="greaterThan", formula=["0"], fill=green_fill),
)
ws.conditional_formatting.add(
    "R18:R18",
    CellIsRule(operator="lessThan", formula=["0"], fill=red_fill),
)

# ========== 06_Bridge ==========
ws = wb.create_sheet(BR)
ws.sheet_properties.tabColor = "C00000"
set_col_widths(ws, [4, 28, 16, 16, 14, 14, 14, 16, 16, 16, 16, 16, 16])
landscape(ws)
ws.freeze_panes = "C8"

ws["B2"] = "Plan vs forecast bridge"
ws["B2"].font = title_font
ws["B3"] = "FY walk: volume (constant mix) + mix + price for revenue; volume + mix + cost/unit for COGS; rate vs volume vs fixed for OpEx; then EBITDA."
ws["B3"].font = italic_grey
ws["B4"] = "Scenario"
ws["C4"] = f"='{ASSUMP}'!D6"
style_formula(ws["C4"], key=True)

ws["B6"] = "Product helpers (FY)"
shade_section(ws, 6, until=14)
helpers = [
    "Product", "Plan vol", "Fcst vol", "Plan $/case", "Fcst $/case",
    "Plan CPU", "Fcst CPU", "Plan rev", "Fcst rev", "Plan COGS", "Fcst COGS",
    "Δvol × plan price", "Fcst vol × Δprice",
]
for i, h in enumerate(helpers):
    style_header_cell(ws.cell(7, 2 + i), h)

for i, name in enumerate(PRODUCTS):
    r = 8 + i
    pr = 26 + i
    ws.cell(r, 2, name).font = bold
    ws.cell(r, 2).border = thin
    cell = ws.cell(r, 3, f"='{REV}'!O{8+i}")
    style_formula(cell)
    cell.number_format = int_fmt
    cell = ws.cell(r, 4, f"='{REV}'!O{33+i}")
    style_formula(cell)
    cell.number_format = num
    cell = ws.cell(r, 5, f"='{ASSUMP}'!C{pr}")
    style_formula(cell)
    cell.number_format = ppu
    cell = ws.cell(r, 6, f"='{REV}'!O{39+i}")
    style_formula(cell)
    cell.number_format = ppu
    cell = ws.cell(r, 7, f"='{ASSUMP}'!H{pr}")
    style_formula(cell)
    cell.number_format = ppu
    cell = ws.cell(r, 8, f"='{ASSUMP}'!I{pr}")
    style_formula(cell)
    cell.number_format = ppu
    cell = ws.cell(r, 9, f"='{REV}'!O{19+i}")
    style_formula(cell)
    cell.number_format = money
    cell = ws.cell(r, 10, f"='{REV}'!O{44+i}")
    style_formula(cell)
    cell.number_format = money
    cell = ws.cell(r, 11, f"='{COGS}'!O{13+i}")
    style_formula(cell)
    cell.number_format = money
    cell = ws.cell(r, 12, f"='{COGS}'!O{28+i}")
    style_formula(cell)
    cell.number_format = money
    cell = ws.cell(r, 13, f"=(D{r}-C{r})*E{r}")
    style_formula(cell)
    cell.number_format = money
    cell = ws.cell(r, 14, f"=D{r}*(F{r}-E{r})")
    style_formula(cell)
    cell.number_format = money

ws["B12"] = "Total"
ws["B12"].font = bold
ws["B12"].border = thin
for c in range(3, 15):
    letter = get_column_letter(c)
    if c in (5, 6, 7, 8):
        # ASP / CPU weighted later; leave blank or average
        cell = ws.cell(12, c, f"=IF(C12=0,0,I12/C12)" if c == 5 else
                       (f"=IF(D12=0,0,J12/D12)" if c == 6 else
                        (f"=IF(C12=0,0,K12/C12)" if c == 7 else f"=IF(D12=0,0,L12/D12)")))
        # That's messy. Simpler:
    cell = ws.cell(12, c, f"=SUM({letter}8:{letter}11)")
    style_formula(cell, key=True)
    cell.number_format = money if c >= 9 else (num if c in (3, 4) else ppu)

# Fix C12 D12 as volume sums (int)
ws["C12"] = "=SUM(C8:C11)"
style_formula(ws["C12"], key=True)
ws["C12"].number_format = num
ws["D12"] = "=SUM(D8:D11)"
style_formula(ws["D12"], key=True)
ws["D12"].number_format = num
ws["E12"] = "=IF(C12=0,0,I12/C12)"
style_formula(ws["E12"], key=True)
ws["E12"].number_format = ppu
ws["F12"] = "=IF(D12=0,0,J12/D12)"
style_formula(ws["F12"], key=True)
ws["F12"].number_format = ppu
ws["G12"] = "=IF(C12=0,0,K12/C12)"
style_formula(ws["G12"], key=True)
ws["G12"].number_format = ppu
ws["H12"] = "=IF(D12=0,0,L12/D12)"
style_formula(ws["H12"], key=True)
ws["H12"].number_format = ppu

# Revenue walk
ws["B14"] = "Revenue walk ($000s)"
shade_section(ws, 14, until=5)
style_header_cell(ws.cell(15, 2), "Step")
style_header_cell(ws.cell(15, 3), "$000s")
style_header_cell(ws.cell(15, 4), "Note")

rev_steps = [
    (16, "Plan revenue", "=I12", "Locked annual plan"),
    (17, "Volume (constant mix)", "=(D12-C12)*E12", "Δ total cases × plan ASP"),
    (18, "Mix", "=M12-C17", "Product mix vs plan ASP: Σ(Δvol × plan price) − volume-at-ASP"),
    (19, "Price", "=N12", "Σ forecast vol × Δ net $/case (includes scenario price factor)"),
    (20, "Forecast revenue", "=C16+C17+C18+C19", "Must equal 02_Revenue_Build FY forecast"),
]
for r, lab, f, note in rev_steps:
    ws.cell(r, 2, lab).border = thin
    ws.cell(r, 2).font = bold if r in (16, 20) else black
    cell = ws.cell(r, 3, f)
    style_formula(cell, key=(r in (16, 20)))
    cell.number_format = money
    ws.cell(r, 4, note).font = small_grey
    ws.cell(r, 4).border = thin

ws["B21"] = "Check vs revenue build"
ws["C21"] = f"=C20-'{REV}'!O48"
style_formula(ws["C21"])
ws["C21"].number_format = money
ws["D21"] = "Should be ~0 (rounding)."
ws["D21"].font = small_grey

# COGS walk
ws["B23"] = "COGS walk ($000s)"
shade_section(ws, 23, until=5)
style_header_cell(ws.cell(24, 2), "Step")
style_header_cell(ws.cell(24, 3), "$000s")
style_header_cell(ws.cell(24, 4), "Note")

# Volume const mix on COGS = Δtotal_vol * plan CPU avg
# Mix = Σ(Δvol_p * plan_cpu_p) - volume_const
# Cost/unit = Σ(fcst_vol_p * Δcpu_p)
ws["C25"] = "=K12"
ws["C26"] = "=(D12-C12)*G12"
ws["C27"] = "=SUMPRODUCT((D8:D11-C8:C11)*G8:G11)-C26"
ws["C28"] = "=SUMPRODUCT(D8:D11,(H8:H11-G8:G11))"
ws["C29"] = "=C25+C26+C27+C28"

cogs_labs = [
    (25, "Plan COGS", "Locked annual plan"),
    (26, "Volume (constant mix)", "Δ total cases × plan CPU"),
    (27, "Mix", "Product mix vs plan CPU"),
    (28, "Cost / unit", "Inflation + scenario unit-cost factor at forecast volume"),
    (29, "Forecast COGS", "Must equal 03_COGS_Build FY forecast"),
]
for r, lab, note in cogs_labs:
    ws.cell(r, 2, lab).border = thin
    ws.cell(r, 2).font = bold if r in (25, 29) else black
    style_formula(ws.cell(r, 3), key=(r in (25, 29)))
    ws.cell(r, 3).number_format = money
    ws.cell(r, 4, note).font = small_grey
    ws.cell(r, 4).border = thin

ws["B30"] = "Check vs COGS build"
ws["C30"] = f"=C29-'{COGS}'!O32"
style_formula(ws["C30"])
ws["C30"].number_format = money

# OpEx walk
ws["B32"] = "OpEx walk ($000s)"
shade_section(ws, 32, until=5)
style_header_cell(ws.cell(33, 2), "Step")
style_header_cell(ws.cell(33, 3), "$000s")
style_header_cell(ws.cell(33, 4), "Note")

# Plan opex O13, fcst O23
# Sales-driven at plan rates: (fcst rev - plan rev) * (plan selling% + plan mkt%)
# Rate change: fcst rev * Δselling% + fcst rev * Δmkt%
# Dist volume: (fcst vol - plan vol) * plan $/case
# Dist rate: fcst vol * Δ$/case
# Fixed + disc: remaining
ws["C34"] = f"='{OPEX}'!O13"
ws["C35"] = f"=('{REV}'!O48-'{REV}'!O23)*('{ASSUMP}'!C33+'{ASSUMP}'!C34)"
ws["C36"] = f"='{REV}'!O48*(('{ASSUMP}'!D33-'{ASSUMP}'!C33)+('{ASSUMP}'!D34-'{ASSUMP}'!C34))"
ws["C37"] = f"=('{REV}'!O37-'{REV}'!O12)*'{ASSUMP}'!C35"
ws["C38"] = f"='{REV}'!O37*('{ASSUMP}'!D35-'{ASSUMP}'!C35)"
ws["C39"] = f"='{OPEX}'!O23-C34-C35-C36-C37-C38"
ws["C40"] = "=C34+C35+C36+C37+C38+C39"

opex_labs = [
    (34, "Plan OpEx", "Locked annual plan"),
    (35, "Sales-driven (plan rates)", "Δ revenue × (plan commission + plan marketing %)"),
    (36, "Selling / marketing rate", "Forecast revenue × Δ rates"),
    (37, "Distribution volume", "Δ cases × plan $ per case"),
    (38, "Distribution rate", "Forecast cases × Δ $ per case"),
    (39, "Fixed & discretionary", "Monthly fixed, G&A/R&D, marketing campaigns, scenario OpEx factor — plug to forecast"),
    (40, "Forecast OpEx", "Must equal 04_OpEx FY forecast"),
]
for r, lab, note in opex_labs:
    ws.cell(r, 2, lab).border = thin
    ws.cell(r, 2).font = bold if r in (34, 40) else black
    style_formula(ws.cell(r, 3), key=(r in (34, 40)))
    ws.cell(r, 3).number_format = money
    ws.cell(r, 4, note).font = small_grey
    ws.cell(r, 4).alignment = Alignment(wrap_text=True)
    ws.cell(r, 4).border = thin
    if r == 39:
        ws.row_dimensions[r].height = 32

ws["B41"] = "Check vs OpEx"
ws["C41"] = f"=C40-'{OPEX}'!O23"
style_formula(ws["C41"])
ws["C41"].number_format = money

# EBITDA walk
ws["B43"] = "EBITDA walk ($000s)"
shade_section(ws, 43, until=5)
style_header_cell(ws.cell(44, 2), "Step")
style_header_cell(ws.cell(44, 3), "$000s")
style_header_cell(ws.cell(44, 4), "Note")

ws["C45"] = f"='{PL}'!P18"
ws["C46"] = "=C17+C18+C19"
ws["C47"] = "=-(C26+C27+C28)"
ws["C48"] = "=-(C35+C36+C37+C38+C39)"
ws["C49"] = "=C45+C46+C47+C48"

ebitda_labs = [
    (45, "Plan EBITDA", "From 05_P&L_Forecast FY plan"),
    (46, "Gross profit — revenue", "Volume + mix + price (helps EBITDA)"),
    (47, "Gross profit — COGS", "Volume + mix + cost/unit (sign flipped: higher COGS hurts)"),
    (48, "OpEx", "Sign flipped: higher OpEx hurts EBITDA"),
    (49, "Forecast EBITDA", "Must equal 05_P&L_Forecast FY forecast"),
]
for r, lab, note in ebitda_labs:
    ws.cell(r, 2, lab).border = thin
    ws.cell(r, 2).font = bold if r in (45, 49) else black
    style_formula(ws.cell(r, 3), key=(r in (45, 49)))
    ws.cell(r, 3).number_format = money
    ws.cell(r, 4, note).font = small_grey
    ws.cell(r, 4).border = thin

ws["B50"] = "Check vs P&L"
ws["C50"] = f"=C49-'{PL}'!O18"
style_formula(ws["C50"])
ws["C50"].number_format = money

# Waterfall chart data (positive/negative split for a bar chart)
ws["B52"] = "Waterfall chart data — EBITDA"
shade_section(ws, 52, until=5)
style_header_cell(ws.cell(53, 2), "Step")
style_header_cell(ws.cell(53, 3), "Base")
style_header_cell(ws.cell(53, 4), "Up")
style_header_cell(ws.cell(53, 5), "Down")
style_header_cell(ws.cell(53, 6), "Connector")

# Invisible connector + up + down stacked bar is the usual Excel waterfall trick
# Rows: Plan, Vol, Mix, Price, COGS vol/mix/cpu combined as GP-COGS, OpEx, Forecast
# Simpler waterfall:
# Plan EBITDA | Rev volume | Rev mix | Rev price | COGS | OpEx | Forecast EBITDA
wf = [
    (54, "Plan EBITDA", "=C45", 0, 0),
    (55, "Volume + mix + price", "=MAX(C46,0)", "=MAX(-C46,0)", "=C45"),
    (56, "COGS", "=MAX(C47,0)", "=MAX(-C47,0)", "=C45+C46"),
    (57, "OpEx", "=MAX(C48,0)", "=MAX(-C48,0)", "=C45+C46+C47"),
    (58, "Forecast EBITDA", "=C49", 0, 0),
]
# For plan and forecast, show as Base (full value). For bridges, stacked.
# Connector for plan/forecast = 0. For others, connector = running total if the step is an increment
# Standard:
# Plan: base=plan, up=0, down=0
# Positive step: connector=prior total, up=step, down=0
# Negative step: connector=prior total + step, up=0, down=-step
# Forecast: base=fcst

ws["B54"] = "Plan EBITDA"
ws["C54"] = "=C45"
ws["D54"] = 0
ws["E54"] = 0
ws["F54"] = 0

ws["B55"] = "Revenue (vol/mix/price)"
ws["C55"] = 0
ws["D55"] = "=MAX(C46,0)"
ws["E55"] = "=MAX(-C46,0)"
ws["F55"] = "=IF(C46>=0,C45,C45+C46)"

ws["B56"] = "COGS"
ws["C56"] = 0
ws["D56"] = "=MAX(C47,0)"
ws["E56"] = "=MAX(-C47,0)"
ws["F56"] = "=IF(C47>=0,C45+C46,C45+C46+C47)"

ws["B57"] = "OpEx"
ws["C57"] = 0
ws["D57"] = "=MAX(C48,0)"
ws["E57"] = "=MAX(-C48,0)"
ws["F57"] = "=IF(C48>=0,C45+C46+C47,C45+C46+C47+C48)"

ws["B58"] = "Forecast EBITDA"
ws["C58"] = "=C49"
ws["D58"] = 0
ws["E58"] = 0
ws["F58"] = 0

for r in range(54, 59):
    ws.cell(r, 2).border = thin
    for c in range(3, 7):
        style_formula(ws.cell(r, c))
        ws.cell(r, c).number_format = money

chart = BarChart()
chart.type = "col"
chart.grouping = "stacked"
chart.title = "EBITDA bridge — plan to forecast ($000s)"
chart.y_axis.title = "$000s"
chart.height = 8
chart.width = 16
chart.legend.position = "b"
data = Reference(ws, min_col=3, min_row=53, max_col=6, max_row=58)
cats = Reference(ws, min_col=2, min_row=54, max_row=58)
chart.add_data(data, from_rows=False, titles_from_data=True)
chart.set_categories(cats)
# Connector series should be no-fill; openpyxl support is limited — leave as stacked, recruiter still sees the walk in the table
ws.add_chart(chart, "B60")

ws["B78"] = (
    "Volume (constant mix) uses plan ASP so mix is the residual of product-level volume. "
    "Price is forecast volume × Δ net $/case. Checks next to each walk should be ~0."
)
ws["B78"].font = small_grey
ws["B78"].alignment = Alignment(wrap_text=True)

# ========== 07_Data_Dictionary ==========
ws = wb.create_sheet("07_Data_Dictionary")
ws.sheet_properties.tabColor = "7F7F7F"
set_col_widths(ws, [4, 34, 22, 82])
landscape(ws)
ws["B2"] = "Data dictionary"
ws["B2"].font = title_font
ws["B3"] = "Field definitions so another analyst can inherit the file."
ws["B3"].font = italic_grey
headers = ["Field", "Tab", "Definition"]
for i, h in enumerate(headers):
    cell = ws.cell(5, 2 + i, h)
    cell.font = header_font
    cell.fill = header_fill
    cell.border = thin

defs = [
    ("Scenario", ASSUMP, "1 Base / 2 Upside / 3 Downside. Indexes the factor table and scales the latest-view forecast only."),
    ("Volume / price / unit-cost / disc. OpEx factors", ASSUMP, "Applied to forecast volume, forecast net $/case, forecast CPU, and forecast marketing + R&D fixed."),
    ("Tax rate", ASSUMP, "Applied to EBIT on the P&L. Sample 25%."),
    ("D&A % of sales", ASSUMP, "Simple D&A driver for the P&L (not an asset rollforward)."),
    ("Cost inflation", ASSUMP, "Forecast CPU = plan CPU × (1 + inflation) × unit-cost factor."),
    ("Plan net $/case", ASSUMP, "Yellow input. Net selling price used in the locked plan."),
    ("Fcst net $/case", ASSUMP, "Yellow input. Latest-view price, then × price factor."),
    ("Material / labor / overhead $/case", ASSUMP, "Plan unit-cost stack. Plan CPU = sum of the three."),
    ("Selling commission %", ASSUMP, "Variable selling cost. Plan and forecast rates can differ."),
    ("Marketing % of sales", ASSUMP, "Brand / trade rate on top of monthly campaign dollars."),
    ("Distribution $ per case", ASSUMP, "Freight and warehousing per case, plus a monthly fixed line."),
    ("Monthly fixed OpEx", ASSUMP, "Yellow inputs by cost center × month, plan and forecast."),
    ("Plan volume", ASSUMP, "Locked plan cases (000s) by product × month."),
    ("Forecast volume (latest view)", ASSUMP, "Reforecast cases (000s) by product × month, before the scenario volume factor."),
    ("Plan / forecast revenue", REV, "Volume × net $/case. $000s = 000 cases × $ per case."),
    ("ASP", REV, "Net revenue ÷ cases. Mix % is share of total cases."),
    ("COGS", COGS, "Volume × CPU. Gross profit = revenue − COGS."),
    ("OpEx by cost center", OPEX, "Selling, marketing, distribution, G&A, R&D. Forecast marketing + R&D fixed × discretionary factor."),
    ("Forecast P&L", PL, "12-month latest view. FY Plan / Var / Var % on the right. D&A = rate × revenue; tax = rate × EBIT."),
    ("Volume (constant mix)", BR, "(Forecast cases − plan cases) × plan ASP."),
    ("Mix", BR, "Σ(Δvol × plan price) − volume-at-ASP. Mix vs the locked plan assortment."),
    ("Price", BR, "Σ forecast vol × (forecast $/case − plan $/case)."),
    ("Cost / unit (COGS walk)", BR, "Σ forecast vol × (forecast CPU − plan CPU): inflation + scenario."),
    ("EBITDA walk", BR, "Plan EBITDA + revenue walk − COGS walk − OpEx walk = forecast EBITDA."),
    ("Units", "All", "Dollars in $000s. Volume in 000 cases. Prices and CPU in $ per case. Rates in %."),
]
for i, (field, tab, definition) in enumerate(defs):
    r = 6 + i
    ws.cell(r, 2, field).font = bold
    ws.cell(r, 2).border = thin
    ws.cell(r, 3, tab).border = thin
    cell = ws.cell(r, 4, definition)
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    cell.border = thin
    ws.row_dimensions[r].height = 32

ws.cell(6 + len(defs), 2, "All sample numbers are fictional. Built for a public GitHub portfolio — no employer data.").font = small_grey

# Sanity print of plan revenue from Python (not Excel)
plan_rev = 0
for i in range(4):
    plan_rev += sum(PLAN_VOL[i]) * PLAN_PRICE[i]
fcst_rev = 0
for i in range(4):
    fcst_rev += sum(FCST_VOL[i]) * FCST_PRICE[i]
print("Python check plan revenue $000s", round(plan_rev))
print("Python check raw fcst revenue $000s (no scenario)", round(fcst_rev))
print("Plan cases", [sum(v) for v in PLAN_VOL], "total", sum(sum(v) for v in PLAN_VOL))
print("Fcst cases", [sum(v) for v in FCST_VOL], "total", sum(sum(v) for v in FCST_VOL))

out = Path(__file__).resolve().parent / "Northline_Driver_Based_Forecast.xlsx"
wb.save(out)
print("Wrote", out)
print("Sheets:", wb.sheetnames)
