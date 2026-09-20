"""QuotaKit — Sales Pipeline & Deal Tracker with Weighted Forecast (Excel + Google Sheets)."""
import sys, random, datetime as dt
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.styles import PatternFill, Font
from qk_style import *

SAMPLE = "--sample" in sys.argv
OUT = sys.argv[-1] if sys.argv[-1].endswith(".xlsx") else "out.xlsx"
NDEALS = 300
NOWN = 20
NSTG = 8

wb = Workbook(); wb.remove(wb.active)
instructions_sheet(
    wb, "Sales Pipeline & Deal Tracker",
    steps=[
        "Settings tab: name your pipeline stages (up to 8) and give each a win probability. Set the forecast start month and list deal owners.",
        "Deals tab: one row per deal — name, account, owner, stage, value, expected close date, created date, next step.",
        "Probability, weighted value, age and 'overdue' flags calculate automatically on the Deals tab.",
        "Dashboard shows open and weighted pipeline, stage breakdown, a 6-month weighted forecast, win rate and your top 10 open deals.",
        "Review weekly: move stages, update close dates, and the forecast rolls forward on its own.",
    ],
    tabs=[
        ("Settings", "Stages + probabilities (the last two must be Closed Won and Closed Lost), forecast start month, currency, owners."),
        ("Deals", "Your deal list, up to 300 rows. Grey columns calculate."),
        ("Dashboard", "Tiles, stage table and chart, monthly weighted forecast, top 10 open deals, overdue count."),
    ],
    tips=[
        "Keep 'Closed Won' and 'Closed Lost' as the last two stages so win-rate and open-pipeline maths stay correct. You can rename the others freely.",
        "Weighted value = deal value × stage probability. It is a forecast, not a promise — calibrate probabilities from your own history every quarter.",
        "A deal is 'overdue' when its expected close date is in the past and it is still open. Overdue deals are the first thing to clean up.",
        "Age = days since the created date, using today's date. It updates every time you open the file.",
    ],
)

# ---------------------------------------------------------------- Settings
ws = wb.create_sheet("Settings"); hide_gridlines(ws)
set_widths(ws, {"A": 3, "B": 22, "C": 16, "D": 3, "E": 30, "F": 16, "G": 3, "H": 24, "I": 3, "J": 56})
title_band(ws, "Settings", "Edit the yellow cells.", 10)
section(ws, 4, 2, "Pipeline stages")
header_row(ws, 5, 2, ["Stage", "Win probability"])
stages = [("Lead", 0.05), ("Qualified", 0.15), ("Discovery", 0.30), ("Proposal", 0.50), ("Negotiation", 0.75), ("Verbal yes", 0.90),
          ("Closed Won", 1.0), ("Closed Lost", 0.0)]
for i, (n, p) in enumerate(stages):
    r = 6 + i
    input_cell(ws, f"B{r}", n); input_cell(ws, f"C{r}", p, fmt=FMT_PCT0)
note(ws, "J5", "Up to 8 stages. Keep Closed Won and Closed Lost in the last two rows — the dashboard treats row 12 as won and row 13 as lost.", merge_to="J7")
section(ws, 4, 5, "General")
label_cell(ws, "E5", "Currency symbol / code", bold=True, fill=FILL_GREY); input_cell(ws, "F5", "$")
label_cell(ws, "E6", "Forecast start month (1st of month)", bold=True, fill=FILL_GREY); input_cell(ws, "F6", dt.date(2026, 9, 1), fmt="mmm-yyyy")
label_cell(ws, "E7", "Won stage row", bold=True, fill=FILL_GREY); formula_cell(ws, "F7", "=B12", align=LEFT)
label_cell(ws, "E8", "Lost stage row", bold=True, fill=FILL_GREY); formula_cell(ws, "F8", "=B13", align=LEFT)
note(ws, "J9", "Forecast start month: the Dashboard forecast shows six months starting from this month. Enter the first of the month, e.g. 01-Sep-2026.", merge_to="J11")
section(ws, 4, 8, "Deal owners")
header_row(ws, 5, 8, ["Owner"])
owners = ["Aisha K.", "Ben C.", "Chloe M.", "Diego A.", "Emma W.", "Farhan A."]
for i in range(NOWN):
    input_cell(ws, f"H{6+i}", owners[i] if (SAMPLE and i < len(owners)) else None)

# ---------------------------------------------------------------- Deals
ws = wb.create_sheet("Deals"); hide_gridlines(ws)
heads = ["Deal name", "Account", "Owner", "Stage", "Value", "Expected close", "Created", "Next step", "Notes",
         "Probability", "Weighted value", "Age (days)", "Open?", "Overdue?", "Close month"]
widths = [28, 22, 14, 16, 14, 14, 14, 26, 26, 12, 14, 11, 8, 10, 12]
title_band(ws, "Deals", "One row per deal. Yellow = you type. Grey = calculates. Up to 300 deals.", len(heads))
header_row(ws, 4, 1, heads, widths=widths)
for c in range(10, 16):
    ws.cell(row=4, column=c).fill = FILL_NAVY
add_list_validation(ws, f"C5:C{4+NDEALS}", f"=Settings!$H$6:$H${5+NOWN}", "Owner from Settings.")
add_list_validation(ws, f"D5:D{4+NDEALS}", f"=Settings!$B$6:$B$13", "Stage from Settings.")
random.seed(5)
accts = ["Northwind", "Acme", "Globex", "Initech", "Umbrella", "Stark", "Wayne", "Hooli", "Vandelay", "Pied Piper", "Soylent", "Wonka", "Dunder Mifflin", "Cyberdyne", "Massive Dynamic", "Tyrell", "Weyland", "Oscorp"]
prods = ["Annual licence", "Onboarding", "Expansion", "Renewal", "Pilot", "Enterprise plan", "Support upgrade", "Training pack"]
today = dt.date(2026, 9, 13)
for i in range(NDEALS):
    r = 5 + i
    if SAMPLE and i < 60:
        stg = random.choices(range(8), weights=[10, 12, 12, 10, 6, 3, 14, 10])[0]
        created = today - dt.timedelta(days=random.randint(5, 180))
        close = created + dt.timedelta(days=random.randint(20, 150)) if stg >= 6 else today + dt.timedelta(days=random.randint(-12, 130))
        a = random.choice(accts)
        input_cell(ws, f"A{r}", f"{a} — {random.choice(prods)}"); input_cell(ws, f"B{r}", a)
        input_cell(ws, f"C{r}", random.choice(owners)); input_cell(ws, f"D{r}", stages[stg][0])
        input_cell(ws, f"E{r}", random.choice([4500, 8000, 12000, 18000, 25000, 40000, 60000, 95000]), fmt=FMT_MONEY)
        input_cell(ws, f"F{r}", close, fmt=FMT_DATE); input_cell(ws, f"G{r}", created, fmt=FMT_DATE)
        input_cell(ws, f"H{r}", random.choice(["Send proposal", "Demo scheduled", "Follow up call", "Legal review", "Awaiting PO", "Intro meeting", ""]))
        input_cell(ws, f"I{r}")
    else:
        for c, f in zip("ABCDEFGHI", [None, None, None, None, FMT_MONEY, FMT_DATE, FMT_DATE, None, None]):
            input_cell(ws, f"{c}{r}", fmt=f)
    formula_cell(ws, f"J{r}", f'=IF(D{r}="","",IFERROR(INDEX(Settings!$C$6:$C$13,MATCH(D{r},Settings!$B$6:$B$13,0)),0))', fmt=FMT_PCT0, fill=FILL_GREY)
    formula_cell(ws, f"K{r}", f'=IF(D{r}="","",N(E{r})*J{r})', fmt=FMT_MONEY, fill=FILL_GREY)
    formula_cell(ws, f"L{r}", f'=IF(G{r}="","",TODAY()-G{r})', fmt="0", fill=FILL_GREY)
    formula_cell(ws, f"M{r}", f'=IF(D{r}="","",IF(OR(D{r}=Settings!$B$12,D{r}=Settings!$B$13),"No","Yes"))', align=CENTER, fill=FILL_GREY)
    formula_cell(ws, f"N{r}", f'=IF(M{r}="","",IF(AND(M{r}="Yes",F{r}<>"",F{r}<TODAY()),"OVERDUE",""))', align=CENTER, fill=FILL_GREY)
    formula_cell(ws, f"O{r}", f'=IF(F{r}="","",DATE(YEAR(F{r}),MONTH(F{r}),1))', fmt="mmm-yyyy", fill=FILL_GREY)
ws.conditional_formatting.add(f"N5:N{4+NDEALS}", CellIsRule(operator="equal", formula=['"OVERDUE"'], fill=PatternFill("solid", fgColor=RED_FILL), font=Font(name="Arial", size=10, bold=True, color="B42318")))
ws.conditional_formatting.add(f"D5:D{4+NDEALS}", FormulaRule(formula=['$D5=Settings!$B$12'], fill=PatternFill("solid", fgColor=GREEN_FILL)))
ws.conditional_formatting.add(f"D5:D{4+NDEALS}", FormulaRule(formula=['$D5=Settings!$B$13'], fill=PatternFill("solid", fgColor=RED_FILL)))
freeze(ws, "B5")
ws.auto_filter.ref = f"A4:O{4+NDEALS}"

# ---------------------------------------------------------------- Dashboard
ws = wb.create_sheet("Dashboard"); hide_gridlines(ws)
set_widths(ws, {"A": 3, "B": 18, "C": 12, "D": 14, "E": 14, "F": 14, "G": 3, "H": 30, "I": 14, "J": 14, "K": 14, "L": 3})
title_band(ws, "Pipeline Dashboard", "Live view of the Deals tab. No inputs here.", 12)
D = "Deals"
def rng(c): return f"{D}!${c}$5:${c}${4+NDEALS}"
WON, LOST = "Settings!$B$12", "Settings!$B$13"
kpi_tile(ws, "B", 4, "OPEN PIPELINE", f'=SUMIF({rng("M")},"Yes",{rng("E")})', FMT_MONEY, 2)
kpi_tile(ws, "D", 4, "WEIGHTED PIPELINE", f'=SUMIF({rng("M")},"Yes",{rng("K")})', FMT_MONEY, 2)
kpi_tile(ws, "F", 4, "OPEN DEALS", f'=COUNTIF({rng("M")},"Yes")', "0", 1)
kpi_tile(ws, "H", 4, "WIN RATE (CLOSED DEALS)", f'=IF(COUNTIF({rng("D")},{WON})+COUNTIF({rng("D")},{LOST})>0,COUNTIF({rng("D")},{WON})/(COUNTIF({rng("D")},{WON})+COUNTIF({rng("D")},{LOST})),0)', FMT_PCT, 1)
kpi_tile(ws, "I", 4, "WON VALUE", f'=SUMIF({rng("D")},{WON},{rng("E")})', FMT_MONEY, 2)
kpi_tile(ws, "K", 4, "OVERDUE", f'=COUNTIF({rng("N")},"OVERDUE")', "0", 1)
section(ws, 7, 2, "Pipeline by stage", 5)
header_row(ws, 8, 2, ["Stage", "Deals", "Value", "Weighted", "Avg age (days)"])
for i in range(NSTG):
    r = 9 + i
    formula_cell(ws, f"B{r}", f'=IF(Settings!B{6+i}="","",Settings!B{6+i})', align=LEFT, fill=FILL_GREY)
    formula_cell(ws, f"C{r}", f'=IF(B{r}="","",COUNTIF({rng("D")},B{r}))', fmt="0")
    formula_cell(ws, f"D{r}", f'=IF(B{r}="","",SUMIF({rng("D")},B{r},{rng("E")}))', fmt=FMT_MONEY)
    formula_cell(ws, f"E{r}", f'=IF(B{r}="","",SUMIF({rng("D")},B{r},{rng("K")}))', fmt=FMT_MONEY)
    formula_cell(ws, f"F{r}", f'=IF(B{r}="","",IFERROR(AVERAGEIF({rng("D")},B{r},{rng("L")}),0))', fmt="0")
label_cell(ws, "B17", "Total", bold=True, fill=FILL_TEAL_LIGHT)
for c in "CDE":
    formula_cell(ws, f"{c}17", f"=SUM({c}9:{c}16)", fmt="0" if c == "C" else FMT_MONEY, bold=True, fill=FILL_TEAL_LIGHT)
label_cell(ws, "F17", "", fill=FILL_TEAL_LIGHT)
ch = BarChart(); ch.type = "bar"; ch.title = "Value by stage (open + closed)"; ch.height, ch.width = 8, 14
ch.add_data(Reference(ws, min_col=4, min_row=8, max_col=5, max_row=16), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, min_row=9, max_row=16))
ch.series[0].graphicalProperties.solidFill = "1F2A44"; ch.series[1].graphicalProperties.solidFill = "0E9F8E"
ch.x_axis.scaling.orientation = "maxMin"
ws.add_chart(ch, "B19")
section(ws, 7, 8, "Weighted forecast by expected close month (open deals)", 4)
header_row(ws, 8, 8, ["Month", "Open deals", "Value", "Weighted"])
for i in range(6):
    r = 9 + i
    formula_cell(ws, f"H{r}", f"=DATE(YEAR(Settings!$F$6),MONTH(Settings!$F$6)+{i},1)", fmt="mmm-yyyy", fill=FILL_GREY, align=LEFT)
    formula_cell(ws, f"I{r}", f'=COUNTIFS({rng("O")},H{r},{rng("M")},"Yes")', fmt="0")
    formula_cell(ws, f"J{r}", f'=SUMIFS({rng("E")},{rng("O")},H{r},{rng("M")},"Yes")', fmt=FMT_MONEY)
    formula_cell(ws, f"K{r}", f'=SUMIFS({rng("K")},{rng("O")},H{r},{rng("M")},"Yes")', fmt=FMT_MONEY, bold=True)
label_cell(ws, "H15", "6-month total", bold=True, fill=FILL_TEAL_LIGHT)
for c in "IJK":
    formula_cell(ws, f"{c}15", f"=SUM({c}9:{c}14)", fmt="0" if c == "I" else FMT_MONEY, bold=True, fill=FILL_TEAL_LIGHT)
label_cell(ws, "H16", "Open deals with close date before the forecast window", fill=FILL_GREY)
formula_cell(ws, "K16", f'=SUMIFS({rng("K")},{rng("O")},"<"&Settings!$F$6,{rng("M")},"Yes")', fmt=FMT_MONEY)
ws.conditional_formatting.add("K9:K14", DataBarRule(start_type="min", end_type="max", color="0E9F8E"))
section(ws, 19, 8, "Top 10 open deals by weighted value", 4)
header_row(ws, 20, 8, ["Deal", "Stage", "Value", "Weighted"])
# rank key column P on Deals (hidden helper): weighted + row/1e6 for open deals, else -1
dws = wb["Deals"]
dws.column_dimensions["P"].width = 12
dws["P4"] = "rank key"; dws["P4"].font = F_NOTE
for i in range(NDEALS):
    r = 5 + i
    formula_cell(dws, f"P{r}", f'=IF(M{r}="Yes",N(K{r})+ROW()/1000000,-1)', fmt="0.000000"); dws[f"P{r}"].font = F_NOTE
KEY = rng("P")
for k in range(10):
    r = 21 + k
    formula_cell(ws, f"H{r}", f'=IFERROR(IF(LARGE({KEY},{k+1})<0,"",INDEX({rng("A")},MATCH(LARGE({KEY},{k+1}),{KEY},0))),"")', align=LEFT)
    formula_cell(ws, f"I{r}", f'=IF(H{r}="","",INDEX({rng("D")},MATCH(LARGE({KEY},{k+1}),{KEY},0)))', align=LEFT)
    formula_cell(ws, f"J{r}", f'=IF(H{r}="","",INDEX({rng("E")},MATCH(LARGE({KEY},{k+1}),{KEY},0)))', fmt=FMT_MONEY)
    formula_cell(ws, f"K{r}", f'=IF(H{r}="","",INDEX({rng("K")},MATCH(LARGE({KEY},{k+1}),{KEY},0)))', fmt=FMT_MONEY, bold=True)
note(ws, "B36", "Weighted pipeline is value × stage probability across open deals. Compare the 6-month weighted total with your target for the same period to see the gap you still need to source.", merge_to="F38")

for s in wb.worksheets:
    s.sheet_properties.tabColor = TEAL if s.title in ("Settings", "Deals") else NAVY
wb["Start Here"].sheet_properties.tabColor = "F2B134"
wb.save(OUT); print("saved", OUT)
