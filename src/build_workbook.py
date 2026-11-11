"""Build the Mini #11 workbook: 10 pivot-table techniques as live formulas.

Sheets: Guide · Data (52 SAMPLE order lines) · OrderTotals (32 orders,
each total a SUMIFS over Data) · Pivots (the 10 techniques, every cell
a live SUMIFS / COUNTIFS / AVERAGEIFS / LARGE / INDEX+MATCH formula)
· Charts (3 native openpyxl charts fed by the Pivots tables).

Honest limit, stated in the README and the Guide sheet: openpyxl
cannot create native Excel pivot tables, so each technique is taught
as a pivot-equivalent summary table built with live formulas, and
outputs/pivot_guide.md gives the under-5-step build for the real
pivot in Excel from the same Data sheet.

No post-2007 Excel function is used anywhere (no AGGREGATE, MINIFS,
MAXIFS, XLOOKUP), so no _xlfn. prefix is needed and LibreOffice
calculates every cell. All money is recomputed from the CSV in
Python first; the expected values live in src/checks.py.
"""

import csv
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.reference import Reference
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "outputs" / "pivot_tables_charts_SAMPLE.xlsx"
OUT.parent.mkdir(parents=True, exist_ok=True)

rows = list(csv.DictReader(open(ROOT / "data" / "raw" / "orders_SAMPLE.csv")))
for r in rows:
    r["qty"] = int(r["qty"])
    r["unit_price"] = int(r["unit_price"])
    r["line_total"] = int(r["line_total"])

# --- expected values, recomputed in Python (checks.py repeats this) ---
GRAND = sum(r["line_total"] for r in rows)
assert GRAND == 346800 and len(rows) == 52
assert len({r["order_id"] for r in rows}) == 32

CATEGORIES = ["Kitchen", "Furniture", "Electronics", "Stationery"]
MONTHS = ["2026-01", "2026-02", "2026-03", "2026-04", "2026-05", "2026-06"]
PRODUCTS = ["Office Chair", "Dinner Set 24pc", "Mechanical Keyboard",
            "Ceramic Mug Set", "Wireless Mouse", "Desk Lamp",
            "Notebook Set", "Laptop Stand"]
CUSTOMERS = ["Ayesha Khan", "Maryam Iqbal", "Bilal Ahmed", "Usman Tariq",
             "Omar Sheikh", "Fatima Noor", "Ali Nawaz", "Hassan Raza",
             "Zainab Ali", "Sana Malik"]
PAYMENTS = ["Cash", "JazzCash", "Easypaisa", "Bank Transfer", "Card"]
WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
            "Saturday", "Sunday"]

DATA_LAST = 2 + len(rows) - 1          # 53
ORD_LAST = 2 + 32 - 1                  # 33
D = f"Data!$K$2:$K${DATA_LAST}"        # line totals
DCAT = f"Data!$H$2:$H${DATA_LAST}"     # category column
DMON = f"Data!$C$2:$C${DATA_LAST}"     # month column
DWD = f"Data!$D$2:$D${DATA_LAST}"      # weekday column
DPROD = f"Data!$G$2:$G${DATA_LAST}"    # product column
OT_CUST = f"OrderTotals!$C$2:$C${ORD_LAST}"
OT_PAY = f"OrderTotals!$D$2:$D${ORD_LAST}"
OT_TOT = f"OrderTotals!$E$2:$E${ORD_LAST}"

PKR = '#,##0'
PCT = '0.00%'

wb = Workbook()

# ---------------------------------------------------------------- Guide
guide = wb.active
guide.title = "Guide"
guide["A1"] = "Excel Pivot Tables & Charts Pack — SAMPLE"
guide["A1"].font = Font(bold=True, size=14)
guide_rows = [
    ("What this is", "10 pivot-table techniques taught on one SAMPLE shop: "
     "the same 32 orders / 52 lines / PKR 346,800 as Mini #1, #3 and #10."),
    ("Data", "52 order lines, Jan–Jun 2026. Line Total, Month and Weekday "
     "are LIVE formulas — change a quantity or a date and every pivot moves."),
    ("OrderTotals", "One row per order; each order total is a SUMIFS over Data. "
     "Techniques 5 and 6 read this sheet so they count ORDERS, not lines."),
    ("Pivots", "All 10 techniques as pivot-equivalent summary tables. Every "
     "value cell is a live SUMIFS / COUNTIFS / AVERAGEIFS / LARGE formula."),
    ("Charts", "3 native Excel charts fed by the Pivots tables: category bar, "
     "monthly line, category-share pie."),
    ("Real pivots", "openpyxl cannot create native pivot tables — that is an "
     "honest limit, not a hidden one. outputs/pivot_guide.md builds each real "
     "pivot in Excel from the Data sheet in under 5 steps."),
    ("SAMPLE", "DailyNeeds-style fictional shop, fictional people. The data is "
     "SAMPLE; the techniques and every formula are real."),
]
for i, (a, b) in enumerate(guide_rows, start=3):
    guide.cell(row=i, column=1, value=a).font = Font(bold=True)
    guide.cell(row=i, column=2, value=b)
guide.column_dimensions["A"].width = 16
guide.column_dimensions["B"].width = 110

# ---------------------------------------------------------------- Data
data = wb.create_sheet("Data")
data.append(["OrderID", "Order Date", "Month", "Weekday", "Customer", "City",
             "Product", "Category", "Qty", "Unit Price (PKR)",
             "Line Total (PKR)", "Payment Method"])
for i, r in enumerate(rows, start=2):
    y, m, d = (int(x) for x in r["order_date"].split("-"))
    data.cell(row=i, column=1, value=int(r["order_id"]))
    c = data.cell(row=i, column=2, value=date(y, m, d))
    c.number_format = "yyyy-mm-dd"
    data.cell(row=i, column=3, value=f'=TEXT(B{i},"yyyy-mm")')
    data.cell(row=i, column=4, value=f'=TEXT(B{i},"dddd")')
    data.cell(row=i, column=5, value=r["customer"])
    data.cell(row=i, column=6, value=r["city"])
    data.cell(row=i, column=7, value=r["product"])
    data.cell(row=i, column=8, value=r["category"])
    data.cell(row=i, column=9, value=r["qty"])
    data.cell(row=i, column=10, value=r["unit_price"])
    data.cell(row=i, column=11, value=f"=I{i}*J{i}")
    data.cell(row=i, column=11).number_format = PKR
    data.cell(row=i, column=12, value=r["payment_method"])
for col, w in zip("ABCDEFGHIJKL", (9, 12, 9, 11, 15, 12, 20, 12, 6, 16, 16, 15)):
    data.column_dimensions[col].width = w

# ---------------------------------------------------------- OrderTotals
ot = wb.create_sheet("OrderTotals")
ot.append(["OrderID", "Order Date", "Customer", "Payment Method",
           "Order Total (PKR)"])
seen = {}
for r in rows:
    seen.setdefault(int(r["order_id"]), (r["order_date"], r["customer"],
                                         r["payment_method"]))
for i, oid in enumerate(sorted(seen), start=2):
    d, cust, pay = seen[oid]
    y, m, dd = (int(x) for x in d.split("-"))
    ot.cell(row=i, column=1, value=oid)
    c = ot.cell(row=i, column=2, value=date(y, m, dd))
    c.number_format = "yyyy-mm-dd"
    ot.cell(row=i, column=3, value=cust)
    ot.cell(row=i, column=4, value=pay)
    ot.cell(row=i, column=5,
           value=f"=SUMIFS(Data!$K$2:$K${DATA_LAST},Data!$A$2:$A${DATA_LAST},A{i})")
    ot.cell(row=i, column=5).number_format = PKR
for col, w in zip("ABCDE", (9, 12, 15, 15, 18)):
    ot.column_dimensions[col].width = w

# --------------------------------------------------------------- Pivots
pv = wb.create_sheet("Pivots")
pv["A1"] = "Pivot tables — every value is a live formula over Data / OrderTotals (SAMPLE)"
pv["A1"].font = Font(bold=True, size=12)


def title(cell, text):
    pv[cell] = text
    pv[cell].font = Font(bold=True)


def header(row, cols):
    for j, h in enumerate(cols, start=1):
        pv.cell(row=row, column=j, value=h).font = Font(bold=True)


# --- (1) Sales by category: A3:B9
title("A3", "(1) Sales by category")
header(4, ["Category", "Sales (PKR)"])
for j, c in enumerate(CATEGORIES):
    r = 5 + j
    pv.cell(row=r, column=1, value=c)
    pv.cell(row=r, column=2, value=f"=SUMIFS({D},{DCAT},$A{r})").number_format = PKR
pv.cell(row=9, column=1, value="Total").font = Font(bold=True)
pv.cell(row=9, column=2, value="=SUM(B5:B8)").font = Font(bold=True)
pv.cell(row=9, column=2).number_format = PKR

# --- (2) Sales by month: D3:E11
title("D3", "(2) Sales by month")
pv.cell(row=4, column=4, value="Month").font = Font(bold=True)
pv.cell(row=4, column=5, value="Sales (PKR)").font = Font(bold=True)
for j, m in enumerate(MONTHS):
    r = 5 + j
    pv.cell(row=r, column=4, value=m)
    pv.cell(row=r, column=5, value=f"=SUMIFS({D},{DMON},$D{r})").number_format = PKR
pv.cell(row=11, column=4, value="Total").font = Font(bold=True)
pv.cell(row=11, column=5, value="=SUM(E5:E10)").font = Font(bold=True)
pv.cell(row=11, column=5).number_format = PKR

# --- (4) Top 5 products: G3:I9 (reads the all-products table G15:H22)
title("G3", "(4) Top 5 products")
for col, h in ((7, "Rank"), (8, "Product"), (9, "Sales (PKR)")):
    pv.cell(row=4, column=col, value=h).font = Font(bold=True)
for j in range(5):
    r = 5 + j
    pv.cell(row=r, column=7, value=j + 1)
    pv.cell(row=r, column=9,
           value=f"=LARGE($J$15:$J$22,G{r})").number_format = PKR
    pv.cell(row=r, column=8,
           value=f"=INDEX($I$15:$I$22,MATCH(I{r},$J$15:$J$22,0))")

# --- All products (source for Top 5): I13:J23
title("I13", "All products (source for Top 5)")
pv.cell(row=14, column=9, value="Product").font = Font(bold=True)
pv.cell(row=14, column=10, value="Sales (PKR)").font = Font(bold=True)
for j, p in enumerate(PRODUCTS):
    r = 15 + j
    pv.cell(row=r, column=9, value=p)
    pv.cell(row=r, column=10,
           value=f"=SUMIFS({D},{DPROD},$I{r})").number_format = PKR
pv.cell(row=23, column=9, value="Total").font = Font(bold=True)
pv.cell(row=23, column=10, value="=SUM(J15:J22)").font = Font(bold=True)
pv.cell(row=23, column=10).number_format = PKR

# --- (3) Category x month: A13:G19
title("A13", "(3) Sales by category x month")
pv.cell(row=14, column=1, value="Category").font = Font(bold=True)
for j, m in enumerate(MONTHS):
    pv.cell(row=14, column=2 + j, value=m).font = Font(bold=True)
for j, c in enumerate(CATEGORIES):
    r = 15 + j
    pv.cell(row=r, column=1, value=c)
    for k, m in enumerate(MONTHS):
        col = get_column_letter(2 + k)
        pv.cell(row=r, column=2 + k,
               value=f"=SUMIFS({D},{DCAT},$A{r},{DMON},{col}$14)").number_format = PKR
pv.cell(row=19, column=1, value="Total").font = Font(bold=True)
for k in range(6):
    col = get_column_letter(2 + k)
    pv.cell(row=19, column=2 + k, value=f"=SUM({col}15:{col}18)").font = Font(bold=True)
    pv.cell(row=19, column=2 + k).number_format = PKR

# --- (5) Average order value by customer: A21:D33
title("A21", "(5) Average order value by customer")
header(22, ["Customer", "Orders", "Total spend (PKR)", "Avg order value (PKR)"])
for j, c in enumerate(CUSTOMERS):
    r = 23 + j
    pv.cell(row=r, column=1, value=c)
    pv.cell(row=r, column=2, value=f"=COUNTIFS({OT_CUST},$A{r})")
    pv.cell(row=r, column=3, value=f"=SUMIFS({OT_TOT},{OT_CUST},$A{r})").number_format = PKR
    pv.cell(row=r, column=4, value=f"=C{r}/B{r}").number_format = PKR
pv.cell(row=33, column=1, value="All customers").font = Font(bold=True)
pv.cell(row=33, column=2, value="=SUM(B23:B32)").font = Font(bold=True)
pv.cell(row=33, column=3, value="=SUM(C23:C32)").font = Font(bold=True)
pv.cell(row=33, column=3).number_format = PKR
pv.cell(row=33, column=4, value="=C33/B33").font = Font(bold=True)
pv.cell(row=33, column=4).number_format = PKR

# --- (6) Order count by payment method: F21:H28
title("F21", "(6) Order count by payment method")
for col, h in ((6, "Payment method"), (7, "Orders"), (8, "Share %")):
    pv.cell(row=22, column=col, value=h).font = Font(bold=True)
for j, p in enumerate(PAYMENTS):
    r = 23 + j
    pv.cell(row=r, column=6, value=p)
    pv.cell(row=r, column=7, value=f"=COUNTIFS({OT_PAY},$F{r})")
    pv.cell(row=r, column=8, value=f"=G{r}/$G$28").number_format = PCT
pv.cell(row=28, column=6, value="Total").font = Font(bold=True)
pv.cell(row=28, column=7, value="=SUM(G23:G27)").font = Font(bold=True)
pv.cell(row=28, column=8, value="=SUM(H23:H27)").font = Font(bold=True)
pv.cell(row=28, column=8).number_format = PCT

# --- (7) Sales share % by category: J25:L31
title("J25", "(7) Sales share % by category")
for col, h in ((10, "Category"), (11, "Sales (PKR)"), (12, "Share %")):
    pv.cell(row=26, column=col, value=h).font = Font(bold=True)
for j, c in enumerate(CATEGORIES):
    r = 27 + j
    pv.cell(row=r, column=10, value=c)
    pv.cell(row=r, column=11, value=f"=SUMIFS({D},{DCAT},$J{r})").number_format = PKR
    pv.cell(row=r, column=12, value=f"=K{r}/$K$31").number_format = PCT
pv.cell(row=31, column=10, value="Total").font = Font(bold=True)
pv.cell(row=31, column=11, value="=SUM(K27:K30)").font = Font(bold=True)
pv.cell(row=31, column=11).number_format = PKR
pv.cell(row=31, column=12, value="=SUM(L27:L30)").font = Font(bold=True)
pv.cell(row=31, column=12).number_format = PCT

# --- (8) Running total by month: A35:C42
title("A35", "(8) Running total by month")
header(36, ["Month", "Sales (PKR)", "Running total (PKR)"])
for j, m in enumerate(MONTHS):
    r = 37 + j
    pv.cell(row=r, column=1, value=m)
    pv.cell(row=r, column=2, value=f"=SUMIFS({D},{DMON},$A{r})").number_format = PKR
    pv.cell(row=r, column=3,
           value=f"=B{r}" if j == 0 else f"=C{r-1}+B{r}").number_format = PKR

# --- (9) Sales by weekday: E35:F44
title("E35", "(9) Sales by weekday")
pv.cell(row=36, column=5, value="Weekday").font = Font(bold=True)
pv.cell(row=36, column=6, value="Sales (PKR)").font = Font(bold=True)
for j, w in enumerate(WEEKDAYS):
    r = 37 + j
    pv.cell(row=r, column=5, value=w)
    pv.cell(row=r, column=6, value=f"=SUMIFS({D},{DWD},$E{r})").number_format = PKR
pv.cell(row=44, column=5, value="Total").font = Font(bold=True)
pv.cell(row=44, column=6, value="=SUM(F37:F43)").font = Font(bold=True)
pv.cell(row=44, column=6).number_format = PKR

# --- (10) Best vs worst month: H35:J39
title("H35", "(10) Best vs worst month")
for col, h in ((8, "Measure"), (9, "Month"), (10, "Sales (PKR)")):
    pv.cell(row=36, column=col, value=h).font = Font(bold=True)
pv.cell(row=37, column=8, value="Best month")
pv.cell(row=37, column=9, value="=INDEX($D$5:$D$10,MATCH(MAX($E$5:$E$10),$E$5:$E$10,0))")
pv.cell(row=37, column=10, value="=MAX($E$5:$E$10)").number_format = PKR
pv.cell(row=38, column=8, value="Worst month")
pv.cell(row=38, column=9, value="=INDEX($D$5:$D$10,MATCH(MIN($E$5:$E$10),$E$5:$E$10,0))")
pv.cell(row=38, column=10, value="=MIN($E$5:$E$10)").number_format = PKR
pv.cell(row=39, column=8, value="Difference (best - worst)").font = Font(bold=True)
pv.cell(row=39, column=10, value="=J37-J38").font = Font(bold=True)
pv.cell(row=39, column=10).number_format = PKR

# --- KPIs: A46:B51
title("A46", "KPIs (live formulas)")
kpis = [
    ("Grand total (PKR)", f"=SUM(Data!$K$2:$K${DATA_LAST})"),
    ("Orders", f"=COUNTA(OrderTotals!$A$2:$A${ORD_LAST})"),
    ("Average order value (PKR)", "=B47/B48"),
    ("Best month", "=INDEX(D5:D10,MATCH(MAX(E5:E10),E5:E10,0))"),
    ("Top category", "=INDEX(A5:A8,MATCH(MAX(B5:B8),B5:B8,0))"),
]
for j, (label, formula) in enumerate(kpis):
    r = 47 + j
    pv.cell(row=r, column=1, value=label)
    pv.cell(row=r, column=2, value=formula)
pv["B47"].number_format = PKR
pv["B49"].number_format = PKR

for col, w in zip("ABCDEFGHIJKL", (26, 16, 18, 13, 14, 15, 8, 22, 22, 16, 16, 12)):
    pv.column_dimensions[col].width = w

# --------------------------------------------------------------- Charts
ch = wb.create_sheet("Charts")
ch["A1"] = "Charts — native Excel charts fed by the Pivots sheet (SAMPLE)"
ch["A1"].font = Font(bold=True, size=12)

bar = BarChart()
bar.type = "col"
bar.title = "Sales by category (PKR)"
bar.y_axis.title = "PKR"
bar.x_axis.title = "Category"
bar.style = 10
bar.height = 7.5
bar.width = 14
bar.add_data(Reference(pv, min_col=2, min_row=4, max_row=8), titles_from_data=True)
bar.set_categories(Reference(pv, min_col=1, min_row=5, max_row=8))
bar.shape = 4
ch.add_chart(bar, "A3")

line = LineChart()
line.title = "Monthly sales trend (PKR)"
line.y_axis.title = "PKR"
line.x_axis.title = "Month"
line.style = 10
line.height = 7.5
line.width = 14
line.add_data(Reference(pv, min_col=5, min_row=4, max_row=10), titles_from_data=True)
line.set_categories(Reference(pv, min_col=4, min_row=5, max_row=10))
ch.add_chart(line, "A20")

pie = PieChart()
pie.title = "Sales share by category"
pie.style = 10
pie.height = 7.5
pie.width = 14
pie.add_data(Reference(pv, min_col=11, min_row=26, max_row=30), titles_from_data=True)
pie.set_categories(Reference(pv, min_col=10, min_row=27, max_row=30))
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.dataLabels.showCatName = True
ch.add_chart(pie, "A37")

wb.save(OUT)
print(f"Wrote {OUT} ({OUT.stat().st_size:,} bytes)")
print(f"Grand total PKR {GRAND:,} | sheets: {wb.sheetnames}")
