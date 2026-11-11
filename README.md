# Mini #11 — Excel Pivot Tables & Charts Pack

**Status:** built and verified locally; GitHub publish pending.

**SAMPLE / DEMO data only.** Fictional shop, fictional people, fictional orders.
Never claimed as real client or real business data.

The Excel companion to the SQL packs: the same shop as Mini #1, Mini #3
and Mini #10 — 12 customers, 9 products, 32 orders, 52 order lines,
Jan–Jun 2026, PKR 346,800 total — now taught as pivot tables. Ten
techniques, each as a pivot-equivalent summary table built with live
SUMIFS / COUNTIFS / AVERAGEIFS / LARGE formulas, plus three native
Excel charts and a step-by-step guide for building each real pivot
in Excel.

## Demo dataset

The exact same shop as Mini #1 / Mini #3 / Mini #10 — no new rows
invented: 32 orders · 52 order lines · PKR 346,800, identical to all
three earlier packs, so everything reconciles. Schema:
`data/README.md` · Source: `data/raw/orders_SAMPLE.csv` · Regenerate:
`python3 src/generate_data.py` (same tables as the Mini #10 seed).

One attribute is new in this pack only: `payment_method`, assigned
deterministically by order (Cash, JazzCash, Easypaisa, Bank Transfer,
Card, cycling by order_id) so technique 6 has a field to pivot on.
It is SAMPLE demo data and changes no money total.

## What the pack teaches

| # | Technique | Pivot setup | Verified result |
|---|---|---|---|
| 1 | Sales by category | Rows = Category, Sum of Line Total | Kitchen PKR 106,000 (top) |
| 2 | Sales by month | Rows = Month, Sum of Line Total | Jun 94,200 (best), Mar 30,500 (worst) |
| 3 | Sales by category × month | Rows = Category, Columns = Month | Every row/column reconciles; Kitchen Jun 43,500 |
| 4 | Top 5 products | Rows = Product, Top 5 by Sum | Office Chair PKR 72,000 |
| 5 | Average order value by customer | Rows = Customer, Average of Order Total | Ayesha Khan PKR 14,175.00 |
| 6 | Order count by payment method | Rows = Payment Method, Count of OrderID | Cash 7, JazzCash 7, others 6 |
| 7 | Sales share % by category | Rows = Category, Sum shown as % of total | Kitchen 30.57% |
| 8 | Running total by month | Rows = Month, Sum shown as Running Total | Jun 346,800 = grand total |
| 9 | Sales by weekday | Rows = Weekday, Sum of Line Total | Thursday PKR 87,700 (top) |
| 10 | Best vs worst month | Month pivot read with MAX / MIN | Jun 94,200 vs Mar 30,500, gap 63,700 |

## What is inside

| File | What it does |
|---|---|
| `outputs/pivot_tables_charts_SAMPLE.xlsx` | The deliverable workbook (Guide + Data + OrderTotals + Pivots + Charts) |
| `outputs/pivot_guide.md` | How to build each real pivot in Excel, under 5 steps each |
| `outputs/summary.json` | Every verified number, machine-readable |
| `data/raw/orders_SAMPLE.csv` | The 52 SAMPLE order lines the workbook reads |
| `data/README.md` | Data dictionary + totals |
| `src/generate_data.py` | Regenerates the sample CSV (deterministic, same shop as Mini #10) |
| `src/build_workbook.py` | Builds the workbook with live formulas + 3 native charts |
| `src/checks.py` | Checks as code — 61 checks, all passing |
| `requirements.txt` | Pinned Python dependency (`openpyxl==3.1.2`) |

## Verified numbers

| KPI | Value |
|---|---|
| Orders / order lines | 32 / 52 (Jan–Jun 2026, SAMPLE) |
| Grand total | PKR 346,800 (= Mini #1 / #3 / #10) |
| Average order value | PKR 10,837.50 |
| Top category | Kitchen, PKR 106,000 (30.57%) |
| Top product | Office Chair, PKR 72,000 |
| Best month / worst month | Jun 2026, PKR 94,200 / Mar 2026, PKR 30,500 |
| Best weekday | Thursday, PKR 87,700 |
| Top customer by average order | Ayesha Khan, PKR 14,175.00 |
| Checks | 61 passing (`python3 src/checks.py`) + LibreOffice headless recalculation |

Verification: `python3 src/checks.py` (all checks pass) plus a
LibreOffice headless recalculation of the workbook — every KPI and
pivot-table value evaluates to the same numbers Python computed
from the CSV (see `outputs/summary.json`).

## How to run it

1. Install the one dependency: `pip install -r requirements.txt`.
2. Open `outputs/pivot_tables_charts_SAMPLE.xlsx` — that is the
   deliverable.
3. To regenerate the sample data: `python3 src/generate_data.py`.
4. To rebuild the workbook: `python3 src/build_workbook.py`.
5. To check it: `python3 src/checks.py`.
6. To practise on your own data: replace the CSV (same columns)
   and re-run steps 4–5 only — do NOT re-run step 3, it regenerates
   the sample data and would overwrite your file.

## Two traps this pack teaches

- **Average per order, not per line (technique 5).** Averaging the
  52 Data lines gives PKR 6,669.23 — meaningless, because one order
  has several lines. The right average reads the OrderTotals sheet
  (one row per order): PKR 10,837.50.
- **Count orders, not lines (technique 6).** Counting the 52 Data
  lines by payment method would count multi-line orders twice.
  Count the 32 OrderTotals rows instead: 7 + 7 + 6 + 6 + 6 = 32.

## Limits

- **SAMPLE data only.** The shop, the people and the orders are
  fictional; the techniques and the formulas are real.
- **No native pivot tables in the file.** openpyxl cannot create
  them — an honest limit, stated on the Guide sheet. Every technique
  is a live-formula summary table instead, and `pivot_guide.md`
  builds each real pivot in Excel in under 5 steps.
- **No post-2007 Excel function is used** (no AGGREGATE, MINIFS,
  MAXIFS, XLOOKUP), so no `_xlfn.` prefix is needed and LibreOffice
  calculates every cell. (Lesson 2026-11-08: those functions need
  the `_xlfn.` prefix in openpyxl output or LibreOffice shows
  #NAME? / #VALUE! — this pack avoids them entirely.)
- **No cached values.** openpyxl writes formulas without results,
  so a viewer that does not calculate may show 0 until Excel or
  Sheets recalculates on open. The formulas are verified by
  `src/checks.py` and by the LibreOffice recalculation, and every
  expected result is listed above and in `outputs/summary.json`.
- **Payment method is a SAMPLE attribute** added for this pack
  only (see above) — it does not exist in Mini #1 / #3 / #10.

## Published

Not yet published on GitHub — local build only.
