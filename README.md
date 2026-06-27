# Shopify Order Reconciliation — Multi-Store SaaS

Builds your monthly reconciliation workbook from Shopify + courier status files,
for any number of stores. Each store has its own config file with product costs
and reference values.

## Folder layout

```
reconcile.py                 ← the script (same for all stores)
stores/
  KUDBI.json                 ← KUDBI store config
  KANTHA.json                ← (when you add a new store, drop a JSON here)
  _TEMPLATE.json             ← copy this to create a new store
```

## How to run

```bash
# For KUDBI:
python reconcile.py --store KUDBI --inputs-dir ./may_inputdata

# For a new store (after creating stores/KANTHA.json):
python reconcile.py --store KANTHA --inputs-dir ./kantha_april

# Override month/year if auto-detection picks wrong:
python reconcile.py --store KUDBI --inputs-dir . --month MAY --year 2026
```

Each `--inputs-dir` folder must contain:
- `shopify_orders.csv`
- `bluedart*.xlsx` (one or more files)
- `expressfly.csv`

Output goes to `/mnt/user-data/outputs/<MONTH>_<YEAR>_<STORE>_ORDERS.xlsx`.

## Adding a new store

1. Copy `stores/_TEMPLATE.json` to `stores/<NEW_STORE>.json`
2. Set `store_name` to the new store's name
3. Fill in `product_costs` with all parent product names and their per-piece costs
4. Optionally update `sample_trible` reference values

That's it — no script changes needed.

## What's auto-filled vs manual

### Now auto-filled (was manual before)
- **I24 Totale Sales** — sum of Total from raw Shopify CSV, **before** unfulfilled orders are removed. This is your top-line gross sales.
- **I25 After Cancellation Sales** — sum of Total of only the fulfilled orders.

### New rule: paid + no tracking = delivered
If a Shopify order is marked `paid` (Financial Status) and is fulfilled, but:
- it's NOT found in the Bluedart status sheet, AND
- it's NOT found in the Expressfly status sheet,

then it's automatically marked as **Delivered** in the "Other" column of the Orders sheet,
and counted in **Other Delivered** in the Reconciliation. A synthetic row is added to
the OtherStatus sheet with Courier Name = "ASSUMED (paid, no tracking)" so you can
audit which orders were treated this way.

### Still manual (yellow-highlighted in Reconciliation)
| Cell | Field |
|---|---|
| I7  | Meta Spend (enter as negative) |
| I8  | Bluedart Shipping (negative) |
| I9  | Other Shipping (negative) |
| I11 | Refunds (negative) |
| I12 | Exchange (negative) |
| I14 | Rent (negative) |
| I15 | Lightbill (negative) |
| I16 | Salary (negative) |
| I18 | Cash Expense (negative) |

Once you fill these in, Gross, Net, ROAS, After-Cancellation ROAS, Meta agency fees,
and all shipping-average percentages recalculate automatically.

## Adding a new product (existing store)

Edit your store's JSON and add an entry under `product_costs`:

```json
"product_costs": {
  ...existing entries...,
  "New Product Name Here": 550
}
```

The "Product Name" must match the part of the Shopify Lineitem name BEFORE the
` - ` separator. Example: lineitem `Mishri Stylish Blue Real Mirror Co-ords Set
for Summer - Blue / M` → parent name `Mishri Stylish Blue Real Mirror Co-ords
Set for Summer`.

If a product appears in the data but isn't in your JSON, its cost cell in
ProductCost will be **blank and highlighted yellow** so you spot it. The script
runs successfully but with no cost — fix the JSON for accurate Product Cost
totals.
