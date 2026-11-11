# Mini #11 — data dictionary (SAMPLE / DEMO only)

Identical to the Mini #1 / Mini #3 / Mini #10 shop: same 32 orders,
52 order lines, Jan–Jun 2026, PKR 346,800. No new rows were invented
for this pack — the pivot questions are answered from the same shop,
flattened to one row per order line.

## data/raw/orders_SAMPLE.csv (52 rows + header)

| Column | Meaning |
|---|---|
| `order_id` | Order key, 1–32 (same as Mini #10) |
| `order_date` | Order date, Jan–Jun 2026 (`yyyy-mm-dd`) |
| `customer` | Fictional customer name (10 of the shop's 12 customers ordered) |
| `city` | Customer city |
| `product` | Product name (8 of the shop's 9 products sold; Bluetooth Speaker never sold) |
| `category` | Electronics / Furniture / Stationery / Kitchen |
| `qty` | Units on this line |
| `unit_price` | PKR actually charged on this line (may differ from list price) |
| `line_total` | `qty * unit_price` — in the workbook this is a live formula, not a typed value |
| `payment_method` | **New in this pack only.** SAMPLE attribute, deterministic cycle by order_id: Cash, JazzCash, Easypaisa, Bank Transfer, Card. Not in Mini #1 / #3 / #10; changes no money total. Exists so technique 6 has a field to pivot on. |

## Workbook sheets built from this CSV

| Sheet | Contents |
|---|---|
| `Data` | The 52 lines; Month (`TEXT(date,"yyyy-mm")`), Weekday (`TEXT(date,"dddd")`) and Line Total (`Qty*Unit Price`) are live formulas |
| `OrderTotals` | One row per order (32); each order total is a SUMIFS over Data |
| `Pivots` | All 10 techniques as live-formula summary tables + KPI block |
| `Charts` | 3 native charts fed by the Pivots tables |

## Totals (verified)

- 32 orders, 52 lines, PKR 346,800 — same as Mini #1, Mini #3, Mini #10.
- Average order value: PKR 10,837.50 (346,800 / 32).
- Monthly revenue: Jan 37,100 · Feb 50,000 · Mar 30,500 ·
  Apr 62,100 · May 72,900 · Jun 94,200.
- Category revenue: Kitchen 106,000 · Furniture 105,600 ·
  Electronics 104,000 · Stationery 31,200.
- Orders by payment method: Cash 7 · JazzCash 7 · Easypaisa 6 ·
  Bank Transfer 6 · Card 6.
