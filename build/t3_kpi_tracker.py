"""QuotaKit — Sales KPI Dashboard & Monthly Tracker (Excel + Google Sheets)."""
import sys, random
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import PatternFill, Font
from qk_style import *

SAMPLE = "--sample" in sys.argv
OUT = sys.argv[-1] if sys.argv[-1].endswith(".xlsx") else "out.xlsx"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

wb = Workbook(); wb.remove(wb.active)
instructions_sheet(
    wb, "Sales KPI Dashboard & Monthly Tracker",
    steps=[
        "Settings tab: currency, year, monthly revenue targets and your KPI benchmarks (win rate, deal size, activity).",
        "Monthly Input tab: at month end, type the seven numbers for that month — leads, meetings, opportunities, deals won, deals lost, revenue, new customers.",
        "KPI Calc fills in conversion rates, attainment, growth and red/amber/green status automatically.",
        "Dashboard shows year-to-date tiles, revenue vs target, the funnel and a month-by-month scorecard.",
        "Print or screenshot the Dashboard for your monthly review.",
    ],
    tabs=[
        ("Settings", "Currency, year, monthly revenue targets, benchmarks and the RAG thresholds."),
        ("Monthly Input", "The only data-entry tab. Seven inputs per month."),
        ("KPI Calc", "Conversion rates, attainment, YTD, run-rate, MoM growth and status per month."),
        ("Dashboard", "Tiles, charts and the scorecard."),
    ],
    tips=[
        "Only fill months that have ended. Empty months are ignored in YTD and averages so your run-rate stays honest.",
        "RAG thresholds: green at or above the benchmark, amber within the tolerance you set, red below that.",
        "Add extra KPIs by copying a row in KPI Calc and pointing it at the new input row.",
        "Works for a team or a single rep — the maths is the same.",
    ],
)

# ---------------------------------------------------------------- Settings
ws = wb.create_sheet("Settings"); hide_gridlines(ws)
set_widths(ws, {"A": 3, "B": 34, "C": 16, "D": 3, "E": 12, "F": 16, "G": 3, "H": 60})
title_band(ws, "Settings", "Edit the yellow cells.", 8)
section(ws, 4, 2, "General")
label_cell(ws, "B5", "Currency symbol / code", bold=True, fill=FILL_GREY); input_cell(ws, "C5", "$")
label_cell(ws, "B6", "Year", bold=True, fill=FILL_GREY); input_cell(ws, "C6", 2026, fmt="0")
section(ws, 8, 2, "Benchmarks (what 'good' looks like)")
bench = [("Win rate (won ÷ (won + lost))", 0.30, FMT_PCT0), ("Average deal size", 8000, FMT_MONEY), ("Leads per month", 120, FMT_INT),
         ("Meetings per month", 40, FMT_INT), ("Lead → meeting rate", 0.35, FMT_PCT0), ("Meeting → opportunity rate", 0.50, FMT_PCT0),
         ("Revenue attainment", 1.0, FMT_PCT0)]
for i, (n, v, f) in enumerate(bench):
    r = 9 + i
    label_cell(ws, f"B{r}", n, bold=True, fill=FILL_GREY); input_cell(ws, f"C{r}", v, fmt=f)
label_cell(ws, "B17", "Amber tolerance (below benchmark)", bold=True, fill=FILL_GREY); input_cell(ws, "C17", 0.15, fmt=FMT_PCT0)
note(ws, "H17", "Green = at or above benchmark. Amber = within this % below it. Red = worse than that.")
section(ws, 4, 5, "Monthly revenue targets")
header_row(ws, 5, 5, ["Month", "Target"])
random.seed(3)
for m in range(12):
    r = 6 + m
    label_cell(ws, f"E{r}", MONTHS[m], bold=True, fill=FILL_GREY)
    input_cell(ws, f"F{r}", 55000 + (m // 3) * 5000, fmt=FMT_MONEY)
label_cell(ws, "E18", "Annual", bold=True, fill=FILL_TEAL_LIGHT); formula_cell(ws, "F18", "=SUM(F6:F17)", fmt=FMT_MONEY, bold=True, fill=FILL_TEAL_LIGHT)

# ---------------------------------------------------------------- Monthly Input
ws = wb.create_sheet("Monthly Input"); hide_gridlines(ws)
set_widths(ws, {"A": 12, "B": 12, "C": 12, "D": 14, "E": 12, "F": 12, "G": 14, "H": 14, "I": 3, "J": 60})
title_band(ws, "Monthly Input", "Seven numbers per month. Leave future months blank.", 10)
header_row(ws, 4, 1, ["Month", "Leads", "Meetings", "Opportunities", "Deals won", "Deals lost", "Revenue", "New customers"])
for m in range(12):
    r = 5 + m
    label_cell(ws, f"A{r}", MONTHS[m], bold=True, fill=FILL_GREY)
    if SAMPLE and m < 8:
        leads = random.randint(95, 150); meet = int(leads * random.uniform(0.28, 0.42)); opp = int(meet * random.uniform(0.4, 0.6))
        won = int(opp * random.uniform(0.25, 0.4)); lost = int(opp * random.uniform(0.3, 0.5)); rev = won * random.randint(6500, 9500)
        vals = [leads, meet, opp, won, lost, rev, max(1, won - random.randint(0, 3))]
    else:
        vals = [None] * 7
    for j, v in enumerate(vals):
        input_cell(ws, f"{L(2+j)}{r}", v, fmt=FMT_MONEY if j == 5 else FMT_INT)
note(ws, "J4", "Leads = new leads created. Meetings = first meetings held. Opportunities = qualified deals opened. "
               "Won/lost = deals closed in the month. Revenue = closed-won value. New customers = first-time buyers.", merge_to="J9")
freeze(ws, "B5")

# ---------------------------------------------------------------- KPI Calc
ws = wb.create_sheet("KPI Calc"); hide_gridlines(ws)
set_widths(ws, {"A": 34})
for c in range(2, 16):
    ws.column_dimensions[L(c)].width = 11
title_band(ws, "KPI Calc", "Everything here is calculated from Monthly Input and Settings.", 15)
header_row(ws, 4, 1, ["KPI"] + MONTHS + ["YTD / Avg", "Benchmark"])
MI = "'Monthly Input'"
rows = [
    # name, per-month formula (col letter c, row r of input), ytd formula, fmt, benchmark ref, higher_is_better
    ("Revenue", lambda c: f"=IF({MI}!G{{r}}=\"\",\"\",{MI}!G{{r}})", "=SUM(B{r}:M{r})", FMT_MONEY, None, True),
    ("Revenue target", lambda c: f"=Settings!F{{sr}}", "=SUMPRODUCT((B5:M5<>\"\")*(B6:M6))", FMT_MONEY, None, True),
    ("Revenue attainment", lambda c: f"=IF({c}5=\"\",\"\",IF({c}6>0,{c}5/{c}6,0))", "=IF(N6>0,N5/N6,0)", FMT_PCT, "Settings!$C$15", True),
    ("Leads", lambda c: f"=IF({MI}!B{{r}}=\"\",\"\",{MI}!B{{r}})", "=IF(COUNT(B8:M8)>0,AVERAGE(B8:M8),0)", FMT_INT, "Settings!$C$11", True),
    ("Meetings", lambda c: f"=IF({MI}!C{{r}}=\"\",\"\",{MI}!C{{r}})", "=IF(COUNT(B9:M9)>0,AVERAGE(B9:M9),0)", FMT_INT, "Settings!$C$12", True),
    ("Opportunities", lambda c: f"=IF({MI}!D{{r}}=\"\",\"\",{MI}!D{{r}})", "=IF(COUNT(B10:M10)>0,AVERAGE(B10:M10),0)", FMT_INT, None, True),
    ("Deals won", lambda c: f"=IF({MI}!E{{r}}=\"\",\"\",{MI}!E{{r}})", "=SUM(B11:M11)", FMT_INT, None, True),
    ("Deals lost", lambda c: f"=IF({MI}!F{{r}}=\"\",\"\",{MI}!F{{r}})", "=SUM(B12:M12)", FMT_INT, None, False),
    ("New customers", lambda c: f"=IF({MI}!H{{r}}=\"\",\"\",{MI}!H{{r}})", "=SUM(B13:M13)", FMT_INT, None, True),
    ("Lead → meeting rate", lambda c: f"=IF({c}8=\"\",\"\",IF({c}8>0,{c}9/{c}8,0))", "=IF(SUM(B8:M8)>0,SUM(B9:M9)/SUM(B8:M8),0)", FMT_PCT, "Settings!$C$13", True),
    ("Meeting → opportunity rate", lambda c: f"=IF({c}9=\"\",\"\",IF({c}9>0,{c}10/{c}9,0))", "=IF(SUM(B9:M9)>0,SUM(B10:M10)/SUM(B9:M9),0)", FMT_PCT, "Settings!$C$14", True),
    ("Win rate", lambda c: f"=IF({c}11=\"\",\"\",IF({c}11+{c}12>0,{c}11/({c}11+{c}12),0))", "=IF(SUM(B11:M12)>0,SUM(B11:M11)/(SUM(B11:M11)+SUM(B12:M12)),0)", FMT_PCT, "Settings!$C$9", True),
    ("Average deal size", lambda c: f"=IF({c}11=\"\",\"\",IF({c}11>0,{c}5/{c}11,0))", "=IF(SUM(B11:M11)>0,SUM(B5:M5)/SUM(B11:M11),0)", FMT_MONEY, "Settings!$C$10", True),
    ("Revenue per lead", lambda c: f"=IF({c}8=\"\",\"\",IF({c}8>0,{c}5/{c}8,0))", "=IF(SUM(B8:M8)>0,SUM(B5:M5)/SUM(B8:M8),0)", FMT_MONEY, None, True),
    ("MoM revenue growth", None, "", FMT_PCT, None, True),
    ("Run-rate (annualised YTD)", None, "=IF(COUNT(B5:M5)>0,N5/COUNT(B5:M5)*12,0)", FMT_MONEY, None, True),
    ("Status vs benchmark", None, "", None, None, True),
]
for i, (name, fn, ytd, fmt, bref, hib) in enumerate(rows):
    r = 5 + i
    label_cell(ws, f"A{r}", name, bold=True, fill=FILL_GREY)
    for m in range(12):
        c = L(2 + m)
        ir = 5 + m; sr = 6 + m
        if name == "MoM revenue growth":
            if m == 0:
                f = '=""'
            else:
                pc = L(1 + m)
                f = f'=IF(OR({c}5="",{pc}5=""),"",IF({pc}5>0,{c}5/{pc}5-1,0))'
        elif name == "Run-rate (annualised YTD)":
            f = f'=IF({c}5="","",SUM($B$5:{c}5)/COUNT($B$5:{c}5)*12)'
        elif name == "Status vs benchmark":
            f = (f'=IF({c}7="","",IF({c}7>=Settings!$C$15,"GREEN",IF({c}7>=Settings!$C$15*(1-Settings!$C$17),"AMBER","RED")))')
        else:
            f = fn(c).replace("{r}", str(ir)).replace("{sr}", str(sr))
        formula_cell(ws, f"{c}{r}", f, fmt=fmt or "General", align=CENTER if name == "Status vs benchmark" else RIGHT)
    if ytd:
        formula_cell(ws, f"N{r}", ytd.replace("{r}", str(r)), fmt=fmt or "General", bold=True, fill=FILL_TEAL_LIGHT)
    else:
        label_cell(ws, f"N{r}", "", fill=FILL_TEAL_LIGHT)
    if bref:
        formula_cell(ws, f"O{r}", f"={bref}", fmt=fmt)
    else:
        label_cell(ws, f"O{r}", "—", align=CENTER)
# status colours
ws.conditional_formatting.add("B21:M21", CellIsRule(operator="equal", formula=['"GREEN"'], fill=PatternFill("solid", fgColor=GREEN_FILL), font=Font(name="Arial", size=10, bold=True, color="1B7F3B")))
ws.conditional_formatting.add("B21:M21", CellIsRule(operator="equal", formula=['"AMBER"'], fill=PatternFill("solid", fgColor=AMBER_FILL), font=Font(name="Arial", size=10, bold=True, color="9A6B00")))
ws.conditional_formatting.add("B21:M21", CellIsRule(operator="equal", formula=['"RED"'], fill=PatternFill("solid", fgColor=RED_FILL), font=Font(name="Arial", size=10, bold=True, color="B42318")))
# attainment colour scale on row 7
ws.conditional_formatting.add("B7:M7", FormulaRule(formula=['AND(B7<>"",B7>=Settings!$C$15)'], fill=PatternFill("solid", fgColor=GREEN_FILL)))
ws.conditional_formatting.add("B7:M7", FormulaRule(formula=['AND(B7<>"",B7<Settings!$C$15*(1-Settings!$C$17))'], fill=PatternFill("solid", fgColor=RED_FILL)))
freeze(ws, "B5")

# ---------------------------------------------------------------- Dashboard
ws = wb.create_sheet("Dashboard"); hide_gridlines(ws)
set_widths(ws, {"A": 3, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13, "H": 13, "I": 13, "J": 3, "K": 26, "L": 13, "M": 13, "N": 13})
title_band(ws, "Sales KPI Dashboard", "Year-to-date view. No inputs on this tab.", 14)
K = "'KPI Calc'"
kpi_tile(ws, "B", 4, "YTD REVENUE", f"={K}!N5", FMT_MONEY)
kpi_tile(ws, "D", 4, "YTD ATTAINMENT", f"={K}!N7", FMT_PCT)
kpi_tile(ws, "F", 4, "WIN RATE (YTD)", f"={K}!N16", FMT_PCT)
kpi_tile(ws, "H", 4, "AVG DEAL SIZE", f"={K}!N17", FMT_MONEY)
kpi_tile(ws, "K", 4, "RUN-RATE (ANNUALISED)", f"={K}!N20", FMT_MONEY, 2)
kpi_tile(ws, "M", 4, "NEW CUSTOMERS YTD", f"={K}!N13", FMT_INT, 2)
section(ws, 7, 2, "Revenue vs target", 6)
header_row(ws, 8, 2, ["Month", "Revenue", "Target", "Attainment", "Status"])
for m in range(12):
    r = 9 + m
    c = L(2 + m)
    label_cell(ws, f"B{r}", MONTHS[m], bold=True, fill=FILL_GREY)
    formula_cell(ws, f"C{r}", f"=IF({K}!{c}5=\"\",0,{K}!{c}5)", fmt=FMT_MONEY)
    formula_cell(ws, f"D{r}", f"={K}!{c}6", fmt=FMT_MONEY)
    formula_cell(ws, f"E{r}", f"=IF({K}!{c}7=\"\",\"\",{K}!{c}7)", fmt=FMT_PCT)
    formula_cell(ws, f"F{r}", f"=IF({K}!{c}21=\"\",\"\",{K}!{c}21)", align=CENTER)
for word, fill, fc in (("GREEN", GREEN_FILL, "1B7F3B"), ("AMBER", AMBER_FILL, "9A6B00"), ("RED", RED_FILL, "B42318")):
    ws.conditional_formatting.add("F9:F20", CellIsRule(operator="equal", formula=[f'"{word}"'], fill=PatternFill("solid", fgColor=fill), font=Font(name="Arial", size=10, bold=True, color=fc)))
ch = BarChart(); ch.type = "col"; ch.title = "Revenue vs target"; ch.height, ch.width = 8, 16
ch.add_data(Reference(ws, min_col=3, min_row=8, max_row=20), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, min_row=9, max_row=20))
ch.series[0].graphicalProperties.solidFill = "0E9F8E"
ln = LineChart(); ln.add_data(Reference(ws, min_col=4, min_row=8, max_row=20), titles_from_data=True)
ln.series[0].smooth = False; ln.series[0].graphicalProperties.line.solidFill = "1F2A44"; ln.series[0].graphicalProperties.line.width = 25000
ch += ln
ws.add_chart(ch, "B22")
section(ws, 7, 11, "Funnel (YTD totals)", 4)
header_row(ws, 8, 11, ["Stage", "Count", "Conversion", "Benchmark"])
funnel = [("Leads", f"=SUM({K}!B8:M8)", "", ""), ("Meetings", f"=SUM({K}!B9:M9)", "=IF(L9>0,L10/L9,0)", "=Settings!$C$13"),
          ("Opportunities", f"=SUM({K}!B10:M10)", "=IF(L10>0,L11/L10,0)", "=Settings!$C$14"),
          ("Deals won", f"=SUM({K}!B11:M11)", "=IF(L11>0,L12/L11,0)", "=Settings!$C$9")]
for i, (n, cnt, conv, b) in enumerate(funnel):
    r = 9 + i
    label_cell(ws, f"K{r}", n, bold=True, fill=FILL_GREY)
    formula_cell(ws, f"L{r}", cnt, fmt=FMT_INT)
    if conv:
        formula_cell(ws, f"M{r}", conv, fmt=FMT_PCT); formula_cell(ws, f"N{r}", b, fmt=FMT_PCT)
    else:
        label_cell(ws, f"M{r}", "—", align=CENTER); label_cell(ws, f"N{r}", "—", align=CENTER)
fc = BarChart(); fc.type = "bar"; fc.title = "Funnel"; fc.height, fc.width = 7, 12; fc.legend = None
fc.add_data(Reference(ws, min_col=12, min_row=8, max_row=12), titles_from_data=True)
fc.set_categories(Reference(ws, min_col=11, min_row=9, max_row=12))
fc.series[0].graphicalProperties.solidFill = "1F2A44"
fc.y_axis.scaling.orientation = "minMax"; fc.x_axis.scaling.orientation = "maxMin"
ws.add_chart(fc, "K14")
note(ws, "K30", "Conversion = this stage ÷ previous stage. Compare with the benchmark column to see where the funnel leaks.", merge_to="N31")

for s in wb.worksheets:
    s.sheet_properties.tabColor = TEAL if s.title in ("Settings", "Monthly Input") else NAVY
wb["Start Here"].sheet_properties.tabColor = "F2B134"
wb.save(OUT); print("saved", OUT)
