# FBA Inventory Health Report — Claude Skill

A Claude skill that generates a comprehensive Amazon FBA Inventory Health Report from 2 CSV files.

## What it does

Upload your **FBA Inventory Report** and **Business Report (by Child ASIN)** from Amazon Seller Central, and Claude will generate a professional Excel workbook with:

| Tab | Purpose |
|-----|---------|
| **Dashboard** | KPIs, lead time assumptions, status breakdown, revenue by model |
| **📦 Restock Order** | Shopping list — what to order, how many, sea vs air |
| **Inventory Detail** | Full SKU breakdown with stock, velocity, age, sessions, CVR |
| **⚠️ Action Items** | Priority-sorted issues: dead stock, OOS, low stock, slow movers |

## Status flags

| Status | Meaning |
|--------|---------|
| 🟢 HEALTHY | Normal stock levels |
| 🔵 LOW STOCK | Less than 14 days supply |
| 🟡 SLOW | 90-180 days supply |
| 🟡 SLOW - Clear | 180+ days supply |
| 🔴 DEAD STOCK | Zero sales in 90 days |
| ⚫ OOS | Out of stock, no inbound |

## Restock recommendations

Based on configurable lead times:

| Method | When |
|--------|------|
| 🚢 Sea Freight | Enough time for ocean shipping (43+ days supply) |
| ✈️ Air Freight | Sea won't make it, air will (28-43 days) |
| 🚨 AIR URGENT | Will stock out before air arrives (<28 days) |

Default lead times (configurable):
- Production + Labeling: 6 days
- Sea Freight: 30 days
- Air Freight: 15 days
- Amazon Check-in: 7 days
- Safety Stock: 14 days
- Restock Target: 60 days cover
- Min Order: 50 units

## How to get the input files

### File 1: FBA Inventory Report
Seller Central → Reports → Fulfillment → Inventory → FBA Inventory → Request .csv Download

### File 2: Business Report
Seller Central → Reports → Business Reports → **Detail Page Sales and Traffic by Child Item** → Select 30 day range → Download

## Usage

### In Claude (with skill installed)
Just upload both CSV files and say "inventory report" or "báo cáo inventory".

### Standalone script
```bash
python scripts/generate_report.py <fba_inventory.csv> <business_report.csv> [options]

# With custom lead times
python scripts/generate_report.py inv.csv biz.csv --sea-days 35 --min-order 100

# Options:
#   --prep-days N     Production + labeling (default: 6)
#   --sea-days N      Sea freight (default: 30)
#   --air-days N      Air freight (default: 15)
#   --checkin-days N  Amazon check-in (default: 7)
#   --safety-days N   Safety stock buffer (default: 14)
#   --target-days N   Target cover after arrival (default: 60)
#   --min-order N     Min order per SKU (default: 50)
#   --output PATH     Output file path
```

## File structure

```
fba-inventory-report/
├── SKILL.md                         # Skill definition (triggers, quick start)
├── README.md                        # This file
├── scripts/
│   ├── __init__.py
│   └── generate_report.py           # Main report generator
├── references/
│   ├── report-spec.md               # Full spec: status rules, restock logic, Excel formatting
│   └── lead-times.md                # Lead time config and restock decision tree
└── assets/
    └── sample-output-screenshot.png  # (optional) Example output
```

## Customization

### Different product line
The skill parses vehicle model names (F150, Silverado, Ram, etc.) from Amazon product titles. To adapt for a different product line, modify the `parse_model_color()` function in `scripts/generate_report.py`.

### Different lead times
Pass custom values via CLI args or tell Claude when running: "tàu đi 35 ngày", "min order 100", etc.

## License

MIT
