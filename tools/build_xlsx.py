"""Builds a maintenance-forecast workbook for one trade. Each workbook covers four buildings being
managed; the Overview sheet scopes a maintenance contract tender for any 3-5 year span.

Usage: python3 tools/build_xlsx.py fire      -> files/fire-maintenance-forecast.xlsx

A new trade (plumbing, lifts...) is a new entry in TRADES below: building names and a few example rows. The layout is shared.
"""
import sys
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule

TRADES = {
    "fire": {
        "file": "fire-maintenance-forecast.xlsx",
        "title": "Fire Maintenance Forecast",
        "buildings": ["Building 1", "Building 2", "Building 3", "Building 4"],
        # task, building index (1-4), frequency, first due year, qty/scope
        "examples": [
            ("Fire door inspection", 1, "Annual", 2027, 12),
            ("Fire booster tank service", 2, "5-yearly", 2029, 1),
            ("Fire indicator panel test", 1, "Quarterly", 2027, 1),
            ("Hydrant flow test", 2, "5-yearly", 2028, 4),
            ("Emergency light discharge test", 3, "6-monthly", 2027, 30),
            ("Annual Fire Safety Statement", 4, "Annual", 2027, 1),
        ],
    },
}

FREQS = [("Monthly", 12, 1), ("Quarterly", 4, 1), ("6-monthly", 2, 1), ("Annual", 1, 1), ("2-yearly", 1, 2),
         ("3-yearly", 1, 3), ("5-yearly", 1, 5), ("7-yearly", 1, 7), ("10-yearly", 1, 10), ("15-yearly", 1, 15)]
N, R0, R1 = 15, 6, 205
FONT = "Arial"
def f(**k): return Font(name=FONT, size=k.pop("size", 10), **k)
HDR = PatternFill("solid", fgColor="1F3A5F"); HF = f(bold=True, color="FFFFFF")
INP = PatternFill("solid", fgColor="FFF9DB"); BLUE = f(color="0000FF")
GREY = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="D0D0D0"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
ORANGE = PatternFill("solid", bgColor="F4B183", fgColor="F4B183")
RED = PatternFill("solid", bgColor="FF9999", fgColor="FF9999")

def title(ws, t, sub):
    ws["A1"] = t; ws["A1"].font = f(bold=True, size=14, color="1F3A5F")
    ws["A2"] = sub; ws["A2"].font = f(italic=True, color="666666")

def hdr(ws, row, c0, c1):
    for col in range(c0, c1 + 1):
        x = ws.cell(row, col); x.fill = HDR; x.font = HF; x.border = BOX
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def build(key):
    T = TRADES[key]
    wb = Workbook()
    g = wb.active; g.title = "Guide"
    w = wb.create_sheet("Overview"); r = wb.create_sheet("Data")
    st = wb.create_sheet("Settings")
    YC = 10                      # first year column on Register (J)
    Y0, Y1 = L(YC), L(YC + N - 1)  # J, X
    TOT, INW, SEQ = L(YC + N), L(YC + N + 1), L(YC + N + 2)   # Y, Z, AA

    # ---- Settings
    title(st, "Settings", "Blue cells on yellow are inputs. Everything else in the workbook reads from here.")
    st["A3"] = "Forecast start year"; st["A3"].font = f(bold=True)
    st["B3"] = 2027; st["B3"].font = BLUE; st["B3"].fill = INP; st["B3"].border = BOX
    st["A7"] = "Buildings (rename the placeholders)"; st["A7"].font = f(bold=True)
    st["A8"] = "All"; st["A8"].font = f(italic=True, color="666666")
    for i, n in enumerate(T["buildings"]):
        x = st.cell(9 + i, 1, n); x.font = BLUE; x.fill = INP; x.border = BOX
    st["D2"] = "Frequencies"; st["D2"].font = f(bold=True)
    for i, h in enumerate(["Frequency", "Services per year", "Cycle (years)"]): st.cell(3, 4 + i, h)
    hdr(st, 3, 4, 6)
    for i in range(20):
        for j in range(3):
            x = st.cell(4 + i, 4 + j); x.fill = INP; x.font = BLUE; x.border = BOX
            if i < len(FREQS): x.value = FREQS[i][j]
    st["D25"] = "Add a custom frequency in a spare row (e.g. 'Annual x2' = 2 per year, cycle 1). Cycle 5 = once every 5 years."
    st["D25"].font = f(italic=True, color="666666", size=9)
    for col, wd in zip("ABCDEF", [46, 6, 3, 18, 18, 14]): st.column_dimensions[col].width = wd

    # ---- Register
    title(r, T["title"] + " - Data", "One row per task per building. Fill the yellow columns; the year grid fills itself. Rows 6-205 are live.")
    heads = ["#", "Task", "Building", "Frequency", "First / anchor due year", "Qty / scope", "Notes", "Services / yr", "Cycle (yrs)"]
    for i, h in enumerate(heads): r.cell(5, 1 + i, h)
    for k in range(N): r.cell(5, YC + k, f"=Settings!$B$3+{k}")
    for off, h in ((N, "15-yr services"), (N + 1, "In window"), (N + 2, "Window seq")): r.cell(5, YC + off, h)
    hdr(r, 5, 1, YC + N + 2); r.row_dimensions[5].height = 42
    r["I4"] = "Services due:"; r["I4"].font = f(bold=True); r["I4"].alignment = Alignment(horizontal="right")
    for k in range(N + 1):
        col = L(YC + k); x = r[f"{col}4"]; x.value = f"=SUM({col}{R0}:{col}{R1})"
        x.font = f(bold=True); x.alignment = Alignment(horizontal="center")
    ex = T["examples"]
    for row in range(R0, R1 + 1):
        i = row - R0
        r[f"A{row}"] = f'=IF(B{row}="","",ROW()-{R0 - 1})'
        r[f"H{row}"] = f'=IFERROR(INDEX(Settings!$E$4:$E$23,MATCH($D{row},Settings!$D$4:$D$23,0)),"")'
        r[f"I{row}"] = f'=IFERROR(INDEX(Settings!$F$4:$F$23,MATCH($D{row},Settings!$D$4:$D$23,0)),"")'
        for k in range(N):
            col = L(YC + k); x = r[f"{col}{row}"]
            x.value = (f'=IF(OR($B{row}="",$E{row}="",$H{row}="",$I{row}=""),0,IF({col}$5<$E{row},0,'
                       f'IF(MOD({col}$5-$E{row},$I{row})=0,$H{row},0)))')
            x.number_format = "0;;"; x.alignment = Alignment(horizontal="center"); x.font = f(); x.border = BOX
        r[f"{TOT}{row}"] = f"=SUM({Y0}{row}:{Y1}{row})"
        r[f"{INW}{row}"] = f'=SUMIFS({Y0}{row}:{Y1}{row},${Y0}$5:${Y1}$5,">="&Overview!$C$4,${Y0}$5:${Y1}$5,"<="&Overview!$C$6)'
        r[f"{SEQ}{row}"] = (f'=IF(AND($B{row}<>"",{INW}{row}>0,OR(Overview!$C$7="All",$C{row}=Overview!$C$7)),'
                            f'COUNT({SEQ}$5:{SEQ}{row - 1})+1,"")')
        for col in (TOT, INW): r[f"{col}{row}"].number_format = "0;;"
        for col in ("A", "H", "I", TOT, INW, SEQ):
            x = r[f"{col}{row}"]; x.font = f(color="808080" if col == SEQ else "000000")
            x.fill = GREY; x.alignment = Alignment(horizontal="center"); x.border = BOX
        for col in "BCDEFG":
            x = r[f"{col}{row}"]; x.fill = INP; x.font = BLUE; x.border = BOX
        for col in "EF": r[f"{col}{row}"].alignment = Alignment(horizontal="center")
        r[f"E{row}"].number_format = "0"
        if i < len(ex):
            n, bi, fr, fy, q = ex[i]
            r[f"B{row}"], r[f"C{row}"], r[f"D{row}"] = n, f"=Settings!$A${8 + bi}", fr
            r[f"E{row}"], r[f"F{row}"] = fy, q
            r[f"G{row}"] = "EXAMPLE - overwrite or delete"
            r[f"C{row}"].font = f(color="008000")
    r.conditional_formatting.add(f"{Y0}{R0}:{Y1}{R1}", CellIsRule(operator="greaterThan", formula=["0"], fill=ORANGE))
    r.conditional_formatting.add(f"A{R0}:E{R1}", FormulaRule(
        formula=[f'AND($B{R0}<>"",OR($C{R0}="",$D{R0}="",$E{R0}=""))'], fill=RED))
    for fm, col in (("=Settings!$A$9:$A$12", "C"), ("=Settings!$D$4:$D$23", "D")):
        dv = DataValidation(type="list", formula1=fm, allow_blank=True, showErrorMessage=True,
                            error="Pick a value from the list (edit the lists in Settings).")
        r.add_data_validation(dv); dv.add(f"{col}{R0}:{col}{R1}")
    dvy = DataValidation(type="whole", operator="between", formula1="1990", formula2="2100", allow_blank=True,
                         showErrorMessage=True, error="Enter a 4-digit year.")
    r.add_data_validation(dvy); dvy.add(f"E{R0}:E{R1}")
    for col, wd in zip("ABCDEFGHI", [5, 34, 18, 13, 12, 8, 34, 10, 8]): r.column_dimensions[col].width = wd
    for k in range(N): r.column_dimensions[L(YC + k)].width = 6.5
    for col in (TOT, INW, SEQ): r.column_dimensions[col].width = 10
    r.freeze_panes = "C6"

    # ---- Window
    title(w, "Overview", "Pick a 3-5 year window (and optionally one building) to list every task that falls due - the scope for the maintenance contract tender.")
    for t, row in (("Window start year", 4), ("Years in window (3-5)", 5), ("Window end year", 6), ("Building filter", 7)):
        w[f"B{row}"] = t; w[f"B{row}"].font = f(bold=True)
    w["C4"] = "=Settings!$B$3"; w["C5"] = 5; w["C6"] = "=MIN(C4+C5-1,Settings!$B$3+14)"; w["C7"] = "All"
    for a in ("C4", "C5", "C7"): w[a].font = BLUE; w[a].fill = INP; w[a].border = BOX
    w["D4"] = "Defaults to the forecast start year - overtype to change."; w["D4"].font = f(italic=True, color="666666", size=9)
    w["C6"].border = BOX; w["C6"].font = f(bold=True)
    for a in ("C4", "C5", "C6", "C7"): w[a].alignment = Alignment(horizontal="center")
    for fm, a in ((f"=Data!${Y0}$5:${Y1}$5", "C4"), ('"3,4,5"', "C5"), ("=Settings!$A$8:$A$12", "C7")):
        dv = DataValidation(type="list", formula1=fm, showErrorMessage=True, error="Pick a value from the list.")
        w.add_data_validation(dv); dv.add(a)
    W0, W1 = 13, 212
    for i, h in enumerate(["#", "Task", "Building", "Frequency", "Qty / scope"]): w.cell(12, 1 + i, h)
    for k in range(5): w.cell(12, 6 + k, f'=IF($C$4+{k}<=$C$6,$C$4+{k},"")')
    w["K12"] = "Total services"; w["L12"] = "Data row"
    hdr(w, 12, 1, 12); w.row_dimensions[12].height = 32
    for row in range(W0, W1 + 1):
        n = row - W0 + 1
        w[f"A{row}"] = f'=IF({n}<=MAX(Data!${SEQ}${R0}:${SEQ}${R1}),{n},"")'
        w[f"L{row}"] = f'=IF($A{row}="","",MATCH($A{row},Data!${SEQ}${R0}:${SEQ}${R1},0))'
        for col, src in (("B", "B"), ("C", "C"), ("D", "D"), ("E", "F")):
            w[f"{col}{row}"] = f'=IF($A{row}="","",INDEX(Data!${src}${R0}:${src}${R1},$L{row}))'
        for k in range(5):
            col = L(6 + k)
            w[f"{col}{row}"] = (f'=IF(OR($A{row}="",{col}$12=""),"",'
                                f'INDEX(Data!${Y0}${R0}:${Y1}${R1},$L{row},{col}$12-Settings!$B$3+1))')
            w[f"{col}{row}"].number_format = "0;;"; w[f"{col}{row}"].alignment = Alignment(horizontal="center")
        w[f"K{row}"] = f'=IF($A{row}="","",SUM(F{row}:J{row}))'
        for col in range(1, 12):
            x = w.cell(row, col); x.border = BOX; x.font = f()
        for col in "AEK": w[f"{col}{row}"].alignment = Alignment(horizontal="center")
        w[f"L{row}"].font = f(color="808080", size=8)
    w.conditional_formatting.add(f"F{W0}:J{W1}", CellIsRule(operator="greaterThan", formula=["0"], fill=ORANGE))
    for col, wd in zip("ABCDEFGHIJKL", [5, 34, 18, 13, 11, 9, 9, 9, 9, 9, 12, 10]): w.column_dimensions[col].width = wd
    w.freeze_panes = "A13"

    # ---- Guide
    title(g, T["title"] + " - 15 years, 4 buildings", "Placeholder building names and example tasks are included - overwrite them.")
    lines = [("How to use", None),
        ("1. Settings", "Set the forecast start year and rename the 4 buildings. Add custom frequencies if needed."),
        ("2. Data", "Add one row per task per building: task, building, frequency and the year it is (or was) first due. Qty / scope and notes are optional but help a tender."),
        ("", "The 15-year grid fills itself: a number is how many services fall due that year. 'First / anchor due year' can be earlier than the start year - the cycle is projected forward from it."),
        ("", "Adding a task = type in the next empty yellow row. To remove one, clear its yellow cells. Rows turn red if building, frequency or year is missing."),
        ("3. Overview", "Choose a start year (it defaults to the forecast start year) and 3, 4 or 5 years, optionally for one building. It lists every task due in that window, year by year - the scope for the maintenance contract tender."),
        ("Colour key", None),
        ("Blue on yellow", "Input cells - the only cells you can edit."), ("Black / grey", "Formulas - locked."),
        ("Green", "Linked from Settings (building names in the example rows)."),
        ("Protection", None),
        ("", "Every sheet is protected so formulas can't be overwritten by accident; only the yellow cells can be edited. There is no password: use Review > Unprotect Sheet (Excel) if you need to change a formula or add rows beyond 200."),
        ("Assumptions", None),
        ("", "Sub-annual frequencies (monthly, quarterly, 6-monthly) count as that many services in every year from the first due year."),
        ("", "Example tasks are illustrative only. Years are whatever basis you choose (calendar or financial) - be consistent."),
        ("", "This is a task forecast, not a budget: there are no costs in it.")]
    row = 4
    for a, b in lines:
        g[f"A{row}"] = a
        if b is None:
            g[f"A{row}"].font = f(bold=True, size=11, color="1F3A5F")
        else:
            g[f"A{row}"].font = f(bold=True); g[f"A{row}"].alignment = Alignment(vertical="top")
            g[f"B{row}"] = b; g[f"B{row}"].font = f(); g[f"B{row}"].alignment = Alignment(wrap_text=True, vertical="top")
        row += 1
    g.column_dimensions["A"].width = 20; g.column_dimensions["B"].width = 110
    for ws in wb: ws.sheet_view.showGridLines = False
    for ws in wb:
        for row in ws.iter_rows():
            for x in row:
                if x.fill is not None and x.fill.fgColor is not None and x.fill.fgColor.rgb in ("00FFF9DB", "FFFFF9DB"):
                    x.protection = Protection(locked=False)
        ws.protection.sheet = True
        ws.protection.formatColumns = False   # False = allowed: let people resize columns and rows
        ws.protection.formatRows = False
    wb.calculation.fullCalcOnLoad = True
    out = Path(__file__).resolve().parent.parent / "files" / T["file"]
    wb.save(out); print("wrote", out)

if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "fire")
