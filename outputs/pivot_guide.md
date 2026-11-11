# Pivot guide — how to build each real pivot in Excel

The workbook's Pivots sheet teaches each technique as a
pivot-equivalent summary table built with live formulas, because
openpyxl cannot create native pivot tables. This guide builds the
**real** pivot for each technique from the same Data sheet
(52 SAMPLE order lines, Jan–Jun 2026).

For every technique: click any cell in the Data sheet, then
**Insert → PivotTable → New Worksheet**. The steps below assume
that starting point. Drag fields into Rows / Columns / Values;
where a step says "change the value field", right-click the value
field → Summarize Values By.

## (1) Sales by category

- **Question:** which category earns the most?
- **Pivot setup:** Rows = Category · Values = Sum of Line Total
- **Workbook formula:** `SUMIFS(Data!K:K, Data!H:H, category)`
- **Build:** 1) Rows ← Category. 2) Values ← Line Total. 3) Confirm it
  says "Sum of Line Total". 4) Sort the sales column largest → smallest.
- **Expected:** Kitchen PKR 106,000 · Furniture 105,600 ·
  Electronics 104,000 · Stationery 31,200 · Total 346,800.

## (2) Sales by month

- **Question:** how did sales move month to month?
- **Pivot setup:** Rows = Month · Values = Sum of Line Total
- **Workbook formula:** `SUMIFS(Data!K:K, Data!C:C, month)`
- **Build:** 1) Rows ← Month. 2) Values ← Line Total (Sum).
  3) Sort months oldest → newest (they already are). 4) Read the trend.
- **Expected:** Jan 37,100 · Feb 50,000 · Mar 30,500 · Apr 62,100 ·
  May 72,900 · Jun 94,200.

## (3) Sales by category × month

- **Question:** which category drove each month?
- **Pivot setup:** Rows = Category · Columns = Month ·
  Values = Sum of Line Total
- **Workbook formula:** `SUMIFS(Data!K:K, Data!H:H, category,
  Data!C:C, month)` — two criteria, one formula.
- **Build:** 1) Rows ← Category. 2) Columns ← Month. 3) Values ←
  Line Total (Sum). 4) Add grand totals (on by default).
- **Expected:** every row and column sums to the (1) and (2) totals;
  e.g. Kitchen in June = PKR 43,500; Furniture in April = 36,000.

## (4) Top 5 products

- **Question:** which 5 products earn the most?
- **Pivot setup:** Rows = Product · Values = Sum of Line Total,
  then Top 10 filter
- **Workbook formula:** `LARGE(product totals, rank)` for the value
  and `INDEX+MATCH` for the name.
- **Build:** 1) Rows ← Product. 2) Values ← Line Total (Sum).
  3) Click the Row Labels arrow → Value Filters → Top 10.
  4) Set Top **5** by Sum of Line Total. 5) OK.
- **Expected:** 1 Office Chair 72,000 · 2 Dinner Set 24pc 58,000 ·
  3 Mechanical Keyboard 51,000 · 4 Ceramic Mug Set 48,000 ·
  5 Wireless Mouse 35,000.

## (5) Average order value by customer

- **Question:** who places the biggest orders on average?
- **Pivot setup:** Rows = Customer · Values = Average of Order Total,
  built on the OrderTotals sheet (one row per order)
- **Workbook formula:** `SUMIFS(order totals, customer) /
  COUNTIFS(orders, customer)` — average per ORDER, not per line.
- **Build:** 1) Build the pivot from the **OrderTotals** sheet, not
  Data (averaging lines would give the wrong answer). 2) Rows ←
  Customer. 3) Values ← Order Total. 4) Summarize Values By →
  Average. 5) Sort largest → smallest.
- **Expected:** Ayesha Khan PKR 14,175.00 (top) · Sana Malik
  PKR 7,866.67 (lowest) · overall average PKR 10,837.50.
- **Trap this teaches:** averaging the 52 lines gives PKR 6,669.23 —
  a meaningless number, because one order has several lines.

## (6) Order count by payment method

- **Question:** how do customers pay, by number of orders?
- **Pivot setup:** Rows = Payment Method · Values = Count of OrderID,
  built on the OrderTotals sheet
- **Workbook formula:** `COUNTIFS(OrderTotals!D:D, method)`
- **Build:** 1) Build the pivot from **OrderTotals**. 2) Rows ←
  Payment Method. 3) Values ← OrderID. 4) Summarize Values By →
  Count. 5) Add "Show Values As → % of Grand Total" for the share.
- **Expected:** Cash 7 · JazzCash 7 · Easypaisa 6 · Bank Transfer 6 ·
  Card 6 · Total 32 orders.
- **Trap this teaches:** counting the 52 Data lines would count
  multi-line orders twice — count orders on OrderTotals instead.

## (7) Sales share % by category

- **Question:** what slice of revenue is each category?
- **Pivot setup:** Rows = Category · Values = Sum of Line Total,
  shown as % of Grand Total
- **Workbook formula:** `SUMIFS(...) / grand total`, formatted as %.
- **Build:** 1) Rows ← Category. 2) Values ← Line Total (Sum).
  3) Right-click a value → Show Values As → % of Grand Total.
  4) Check the column sums to 100%.
- **Expected:** Kitchen 30.57% · Furniture 30.45% ·
  Electronics 29.99% · Stationery 9.00%.

## (8) Running total by month

- **Question:** how much had we earned by the end of each month?
- **Pivot setup:** Rows = Month · Values = Sum of Line Total,
  shown as Running Total
- **Workbook formula:** first month = its sales; each later month =
  previous running total + its sales.
- **Build:** 1) Rows ← Month. 2) Values ← Line Total (Sum).
  3) Right-click a value → Show Values As → Running Total In →
  Month. 4) Read the last row: it must equal the grand total.
- **Expected:** Jan 37,100 · Feb 87,100 · Mar 117,600 · Apr 179,700 ·
  May 252,600 · Jun 346,800 (= grand total — the check that the
  running total is right).

## (9) Sales by weekday

- **Question:** which weekday sells best?
- **Pivot setup:** Rows = Weekday · Values = Sum of Line Total
- **Workbook formula:** `SUMIFS(Data!K:K, Data!D:D, weekday)` — the
  Weekday column is itself a formula, `TEXT(date, "dddd")`.
- **Build:** 1) Rows ← Weekday. 2) Values ← Line Total (Sum).
  3) Sort largest → smallest. 4) Note the weekday order is
  alphabetical in the pivot — sort by value, not by name.
- **Expected:** Thursday 87,700 (top) · Saturday 81,300 · Friday
  77,100 · Monday 35,900 · Sunday 31,700 · Wednesday 21,300 ·
  Tuesday 11,800 (lowest).

## (10) Best vs worst month comparison

- **Question:** how far apart are the best and worst months?
- **Pivot setup:** the month pivot from (2), read with MAX / MIN
- **Workbook formula:** best month =
  `INDEX(months, MATCH(MAX(sales), sales, 0))`; worst the same with
  MIN; difference = best − worst.
- **Build:** 1) Build the month pivot from (2). 2) Outside the pivot,
  `=MAX(sales range)` and `=MIN(sales range)`. 3) `INDEX+MATCH` each
  back to its month name. 4) Subtract for the gap.
- **Expected:** Best June 2026, PKR 94,200 · Worst March 2026,
  PKR 30,500 · Difference PKR 63,700 (June earned 3.09× March).

## The three charts

- **Bar (category sales):** Insert → Column Chart from the (1) table.
- **Line (monthly trend):** Insert → Line Chart from the (2) table.
- **Pie (category share):** Insert → Pie Chart from the (7) table,
  with % data labels on. All three are already built on the Charts
  sheet — rebuild them once by hand and the skill sticks.
