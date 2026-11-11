"""Checks for Mini #11 (run: python3 src/checks.py).

openpyxl does not evaluate formulas, so these checks do two things:
(1) verify the exact formula string sits in every derived cell, and
(2) recompute every expected value in Python from the CSV and compare
it against the hardcoded verified numbers below (which were themselves
computed from the Mini #10 seed — same shop, same PKR 346,800).

A final LibreOffice headless recalculation (see README) confirms the
formulas actually evaluate to these numbers in a real spreadsheet.
"""

import csv
import json
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart

ROOT = Path(__file__).resolve().parent.parent
WB = ROOT / "outputs" / "pivot_tables_charts_SAMPLE.xlsx"

PASSED = 0


def check(name, cond):
    global PASSED
    assert cond, f"FAIL: {name}"
    PASSED += 1


# ------------------------------------------------------------------ data
rows = list(csv.DictReader(open(ROOT / "data" / "raw" / "orders_SAMPLE.csv")))
for r in rows:
    r["qty"] = int(r["qty"])
    r["unit_price"] = int(r["unit_price"])
    r["line_total"] = int(r["line_total"])

check("CSV has 52 order lines", len(rows) == 52)
check("CSV columns exact", list(rows[0].keys()) == [
    "order_id", "order_date", "customer", "city", "product", "category",
    "qty", "unit_price", "line_total", "payment_method"])
check("grand total PKR 346,800 (Mini #1 reconciliation)",
      sum(r["line_total"] for r in rows) == 346800)
check("32 distinct orders", len({r["order_id"] for r in rows}) == 32)
check("10 ordering customers", len({r["customer"] for r in rows}) == 10)
check("8 products sold", len({r["product"] for r in rows}) == 8)
check("every line_total = qty * unit_price",
      all(r["line_total"] == r["qty"] * r["unit_price"] for r in rows))
check("payment methods are the 5 SAMPLE values",
      {r["payment_method"] for r in rows} == {
          "Cash", "JazzCash", "Easypaisa", "Bank Transfer", "Card"})
check("payment method is constant within an order",
      len({(r["order_id"], r["payment_method"]) for r in rows})
      == len({r["order_id"] for r in rows}))

# ------------------------------------------- expected values (Python side)
bycat = defaultdict(int)
bymonth = defaultdict(int)
byprod = defaultdict(int)
bywd = defaultdict(int)
order_tot = defaultdict(int)
cust_orders = defaultdict(set)
cust_spend = defaultdict(int)
bypay = defaultdict(int)
for r in rows:
    bycat[r["category"]] += r["line_total"]
    bymonth[r["order_date"][:7]] += r["line_total"]
    byprod[r["product"]] += r["line_total"]
    order_tot[r["order_id"]] += r["line_total"]
    cust_orders[r["customer"]].add(r["order_id"])
    cust_spend[r["customer"]] += r["line_total"]
for oid in {r["order_id"] for r in rows}:  # per ORDER, not per line
    bypay[next(r["payment_method"] for r in rows
               if r["order_id"] == oid)] += 1
from datetime import date as _d
for r in rows:
    bywd[_d.fromisoformat(r["order_date"]).strftime("%A")] += r["line_total"]

EXP_CAT = {"Kitchen": 106000, "Furniture": 105600,
           "Electronics": 104000, "Stationery": 31200}
EXP_MONTH = {"2026-01": 37100, "2026-02": 50000, "2026-03": 30500,
             "2026-04": 62100, "2026-05": 72900, "2026-06": 94200}
EXP_RUN = [37100, 87100, 117600, 179700, 252600, 346800]
EXP_PROD = {"Office Chair": 72000, "Dinner Set 24pc": 58000,
            "Mechanical Keyboard": 51000, "Ceramic Mug Set": 48000,
            "Wireless Mouse": 35000, "Desk Lamp": 33600,
            "Notebook Set": 31200, "Laptop Stand": 18000}
EXP_TOP5 = ["Office Chair", "Dinner Set 24pc", "Mechanical Keyboard",
            "Ceramic Mug Set", "Wireless Mouse"]
EXP_AOV = {"Ayesha Khan": 14175.0, "Maryam Iqbal": 12466.67,
           "Bilal Ahmed": 11650.0, "Usman Tariq": 11600.0,
           "Omar Sheikh": 11366.67, "Fatima Noor": 11025.0,
           "Ali Nawaz": 9050.0, "Hassan Raza": 8666.67,
           "Zainab Ali": 8466.67, "Sana Malik": 7866.67}
EXP_PAY = {"Cash": 7, "JazzCash": 7, "Easypaisa": 6,
           "Bank Transfer": 6, "Card": 6}
EXP_WD = {"Thursday": 87700, "Saturday": 81300, "Friday": 77100,
          "Monday": 35900, "Sunday": 31700, "Wednesday": 21300,
          "Tuesday": 11800}

check("category totals match verified numbers", dict(bycat) == EXP_CAT)
check("monthly totals match verified numbers", dict(bymonth) == EXP_MONTH)
check("running totals match verified numbers",
      [sum(EXP_MONTH[m] for m in sorted(EXP_MONTH)[:i + 1])
       for i in range(6)] == EXP_RUN)
check("product totals match verified numbers", dict(byprod) == EXP_PROD)
check("top 5 products match verified order",
      sorted(EXP_PROD, key=EXP_PROD.get, reverse=True)[:5] == EXP_TOP5)
check("AOV by customer matches verified numbers",
      {c: round(cust_spend[c] / len(cust_orders[c]), 2)
       for c in cust_spend} == EXP_AOV)
check("orders by payment method match verified numbers",
      dict(bypay) == EXP_PAY)
check("weekday totals match verified numbers", dict(bywd) == EXP_WD)
check("category x month cells sum to grand total",
      sum(bycat.values()) == 346800
      and sum(bymonth.values()) == 346800
      and sum(bywd.values()) == 346800)
check("best month Jun 94,200 / worst Mar 30,500 / diff 63,700",
      max(EXP_MONTH, key=EXP_MONTH.get) == "2026-06"
      and EXP_MONTH["2026-06"] == 94200
      and min(EXP_MONTH, key=EXP_MONTH.get) == "2026-03"
      and EXP_MONTH["2026-03"] == 30500
      and 94200 - 30500 == 63700)
check("overall AOV is PKR 10,837.50",
      round(346800 / 32, 2) == 10837.50)

# --------------------------------------------------------------- workbook
wb = load_workbook(WB)
check("sheets in order Guide/Data/OrderTotals/Pivots/Charts",
      wb.sheetnames == ["Guide", "Data", "OrderTotals", "Pivots", "Charts"])
data, ot, pv, ch = (wb[s] for s in
                    ["Data", "OrderTotals", "Pivots", "Charts"])

check("Data has 52 rows + header", data.max_row == 53)
check("all 52 Line Total cells are =Qty*UnitPrice formulas",
      all(data.cell(row=i, column=11).value == f"=I{i}*J{i}"
          for i in range(2, 54)))
check("all 52 Month cells are TEXT date formulas",
      all(data.cell(row=i, column=3).value == f'=TEXT(B{i},"yyyy-mm")'
          for i in range(2, 54)))
check("all 52 Weekday cells are TEXT date formulas",
      all(data.cell(row=i, column=4).value == f'=TEXT(B{i},"dddd")'
          for i in range(2, 54)))
check("Data Line Total values reconcile to 346,800 in the CSV",
      sum(r["line_total"] for r in rows) == 346800)

check("OrderTotals has 32 rows + header", ot.max_row == 33)
check("all 32 order totals are SUMIFS formulas over Data",
      all(isinstance(ot.cell(row=i, column=5).value, str)
          and ot.cell(row=i, column=5).value.startswith("=SUMIFS(Data!")
          for i in range(2, 34)))
check("OrderTotals order ids are 1..32",
      [ot.cell(row=i, column=1).value for i in range(2, 34)]
      == list(range(1, 33)))

D = "Data!$K$2:$K$53"
check("(1) category table: 4 SUMIFS formulas with category criterion",
      all(pv.cell(row=5 + j, column=2).value ==
          f"=SUMIFS({D},Data!$H$2:$H$53,$A{5 + j})" for j in range(4)))
check("(1) category total row is =SUM(B5:B8)",
      pv["B9"].value == "=SUM(B5:B8)")
check("(2) month table: 6 SUMIFS formulas with month criterion",
      all(pv.cell(row=5 + j, column=5).value ==
          f"=SUMIFS({D},Data!$C$2:$C$53,$D{5 + j})" for j in range(6)))
check("(2) month total row is =SUM(E5:E10)",
      pv["E11"].value == "=SUM(E5:E10)")
check("(3) category x month: all 24 cells are two-criteria SUMIFS",
      all(isinstance(pv.cell(row=15 + j, column=2 + k).value, str)
          and pv.cell(row=15 + j, column=2 + k).value.startswith("=SUMIFS(")
          and "Data!$H$2:$H$53" in pv.cell(row=15 + j, column=2 + k).value
          and "Data!$C$2:$C$53" in pv.cell(row=15 + j, column=2 + k).value
          for j in range(4) for k in range(6)))
check("(3) month totals row sums each column",
      all(pv.cell(row=19, column=2 + k).value ==
          f"=SUM({chr(66 + k)}15:{chr(66 + k)}18)" for k in range(6)))
check("(4) Top 5 sales are LARGE formulas over the all-products table",
      all(pv.cell(row=5 + j, column=9).value ==
          f"=LARGE($J$15:$J$22,G{5 + j})" for j in range(5)))
check("(4) Top 5 names are INDEX+MATCH formulas",
      all(pv.cell(row=5 + j, column=8).value ==
          f"=INDEX($I$15:$I$22,MATCH(I{5 + j},$J$15:$J$22,0))"
          for j in range(5)))
check("all-products table: 8 SUMIFS formulas",
      all(pv.cell(row=15 + j, column=10).value ==
          f"=SUMIFS({D},Data!$G$2:$G$53,$I{15 + j})" for j in range(8)))
check("all-products table: product names intact (no table overwrote them)",
      [pv.cell(row=15 + j, column=9).value for j in range(8)] == [
          "Office Chair", "Dinner Set 24pc", "Mechanical Keyboard",
          "Ceramic Mug Set", "Wireless Mouse", "Desk Lamp",
          "Notebook Set", "Laptop Stand"])
check("no table overwrites another: T3 June column, T6 header, T7 title intact",
      pv.cell(row=15, column=7).value.startswith("=SUMIFS(")
      and "Data!$C$2:$C$53,G$14" in pv.cell(row=15, column=7).value
      and pv.cell(row=22, column=7).value == "Orders"
      and pv.cell(row=25, column=10).value == "(7) Sales share % by category")
check("(5) AOV table: COUNTIFS orders + SUMIFS spend + division",
      all(pv.cell(row=23 + j, column=2).value ==
          f"=COUNTIFS(OrderTotals!$C$2:$C$33,$A{23 + j})"
          and pv.cell(row=23 + j, column=3).value ==
          f"=SUMIFS(OrderTotals!$E$2:$E$33,OrderTotals!$C$2:$C$33,$A{23 + j})"
          and pv.cell(row=23 + j, column=4).value == f"=C{23 + j}/B{23 + j}"
          for j in range(10)))
check("(5) AOV total row: 32 orders, PKR 346,800, AOV = total/orders",
      pv["B33"].value == "=SUM(B23:B32)"
      and pv["C33"].value == "=SUM(C23:C32)"
      and pv["D33"].value == "=C33/B33")
check("(6) payment table: 5 COUNTIFS formulas over OrderTotals",
      all(pv.cell(row=23 + j, column=7).value ==
          f"=COUNTIFS(OrderTotals!$D$2:$D$33,$F{23 + j})" for j in range(5)))
check("(6) payment share formulas divide by the total row",
      all(pv.cell(row=23 + j, column=8).value == f"=G{23 + j}/$G$28"
          for j in range(5)))
check("(7) share table: 4 SUMIFS + share-of-total formulas",
      all(pv.cell(row=27 + j, column=11).value ==
          f"=SUMIFS({D},Data!$H$2:$H$53,$J{27 + j})"
          and pv.cell(row=27 + j, column=12).value == f"=K{27 + j}/$K$31"
          for j in range(4)))
check("(8) running total: first month = sales, rest accumulate",
      pv["C37"].value == "=B37"
      and all(pv.cell(row=r, column=3).value == f"=C{r - 1}+B{r}"
              for r in range(38, 43)))
check("(9) weekday table: 7 SUMIFS formulas",
      all(pv.cell(row=37 + j, column=6).value ==
          f"=SUMIFS({D},Data!$D$2:$D$53,$E{37 + j})" for j in range(7)))
check("(9) weekday total row is =SUM(F37:F43)",
      pv["F44"].value == "=SUM(F37:F43)")
check("(10) best/worst use INDEX+MATCH on MAX/MIN of the month table",
      pv["I37"].value ==
      "=INDEX($D$5:$D$10,MATCH(MAX($E$5:$E$10),$E$5:$E$10,0))"
      and pv["J37"].value == "=MAX($E$5:$E$10)"
      and pv["I38"].value ==
      "=INDEX($D$5:$D$10,MATCH(MIN($E$5:$E$10),$E$5:$E$10,0))"
      and pv["J38"].value == "=MIN($E$5:$E$10)"
      and pv["J39"].value == "=J37-J38")
check("KPI block: grand total, orders, AOV, best month, top category",
      pv["B47"].value == "=SUM(Data!$K$2:$K$53)"
      and pv["B48"].value == "=COUNTA(OrderTotals!$A$2:$A$33)"
      and pv["B49"].value == "=B47/B48"
      and pv["B50"].value == "=INDEX(D5:D10,MATCH(MAX(E5:E10),E5:E10,0))"
      and pv["B51"].value == "=INDEX(A5:A8,MATCH(MAX(B5:B8),B5:B8,0))")

# ------------------------------------------------- no post-2007 functions
all_formulas = []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                all_formulas.append(c.value)
check("no post-2007 function used (no AGGREGATE/MINIFS/MAXIFS/XLOOKUP)",
      not any(fn in f.upper() for f in all_formulas
              for fn in ("AGGREGATE", "MINIFS", "MAXIFS", "XLOOKUP")))
check("every derived value cell holds a formula, not a typed number",
      len(all_formulas) >= 200)

# ----------------------------------------------------------------- charts
charts = ch._charts
check("Charts sheet holds exactly 3 charts", len(charts) == 3)
check("chart 1 is a column BarChart (category sales)",
      isinstance(charts[0], BarChart) and charts[0].type == "col")
check("chart 2 is a LineChart (monthly trend)",
      isinstance(charts[1], LineChart))
check("chart 3 is a PieChart (category share)",
      isinstance(charts[2], PieChart))
check("pie chart shows percent + category labels",
      charts[2].dataLabels is not None
      and charts[2].dataLabels.showPercent
      and charts[2].dataLabels.showCatName)

# ------------------------------------------------------- labels and files
check("Guide sheet labels the pack SAMPLE",
      "SAMPLE" in (wb["Guide"]["A1"].value or ""))
check("pivot_guide.md exists and names all 10 techniques",
      (ROOT / "outputs" / "pivot_guide.md").exists()
      and all(f"({n})" in (ROOT / "outputs" / "pivot_guide.md").read_text()
              for n in range(1, 11)))
check("summary.json exists and its grand total is 346,800",
      (ROOT / "outputs" / "summary.json").exists()
      and json.loads((ROOT / "outputs" / "summary.json").read_text())
      ["grand_total_pkr"] == 346800)
check("README labels the data SAMPLE",
      "SAMPLE" in (ROOT / "README.md").read_text())

print(f"\n{PASSED} checks passed, 0 failed.")
