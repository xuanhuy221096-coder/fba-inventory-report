---
name: fba-inventory-report
description: "Build a comprehensive FBA Inventory Health Report from 2 Amazon CSV files: FBA Inventory report and Business Report by Child ASIN. Generates a multi-tab Excel workbook with Dashboard, Restock Order, Inventory Detail, and Action Items — including restock qty recommendations with sea vs air shipping. Use this skill whenever the user uploads Amazon inventory files and wants an inventory report, restock recommendations, stock health check, or wants to know which items are slow/dead/OOS. Trigger on phrases like 'inventory report', 'báo cáo inventory', 'báo cáo tồn kho', 'restock report', 'stock health', 'hàng ế', 'hàng cần nhập', 'inventory health', 'FBA inventory analysis', 'check my inventory', 'làm báo cáo hàng', or whenever 2 CSV files are uploaded that look like Amazon FBA Inventory + Business Report."
---

# FBA Inventory Health Report

Generate a professional multi-tab Excel inventory report from 2 Amazon CSV files.

The report identifies stock health, flags dead/slow/OOS items, and recommends restock quantities with shipping method (sea vs air) based on lead times and sell velocity.

---

## Quick Start

1. User uploads 2 CSV files (FBA Inventory + Business Report by Child ASIN)
2. Read `/mnt/skills/public/xlsx/SKILL.md` for Excel generation best practices
3. Run the report generator script: `python {skill_path}/scripts/generate_report.py <fba_inventory.csv> <business_report.csv>`
4. Present the output file to the user with `present_files`
5. Provide a brief text summary of key findings

If the script fails or the user wants customization, read `references/report-spec.md` for the full specification and build manually.

---

## Required Inputs

| File | Source | How to identify |
|------|--------|-----------------|
| **FBA Inventory Report** | Seller Central → Reports → Fulfillment → Inventory → FBA Inventory | Has columns: `snapshot-date`, `fnsku`, `inv-age-0-to-90-days`, `units-shipped-t7` |
| **Business Report** | Seller Central → Reports → Business Reports → Detail Page Sales and Traffic by **Child Item** (30d range) | Has columns: `(Parent) ASIN`, `(Child) ASIN`, `Sessions - Total`, `Unit Session Percentage` |

If a file doesn't match either pattern, ask the user which is which.

If Business Report is aggregated by date (not by ASIN), tell user to re-download with "Detail Page Sales and Traffic by Child Item" selected.

---

## Lead Time Defaults

| Stage | Default |
|-------|---------|
| Production + Labeling | 6 days |
| Sea Freight | 30 days |
| Air Freight | 15 days |
| Amazon Check-in | 7 days |
| **Total Lead (Sea)** | **43 days** |
| **Total Lead (Air)** | **28 days** |
| Safety Stock | 14 days |
| Restock Target | 60 days cover |
| Min Order Qty | 50 units |

If the user provides different values, pass them as arguments to the script or adjust manually.

---

## Output Structure

The Excel file has 4 tabs:

1. **Dashboard** — KPIs, lead time assumptions, status breakdown, restock summary, revenue by model
2. **📦 Restock Order** — Only SKUs that need restocking, sorted by urgency, with order qty (min 50) and ship method
3. **Inventory Detail** — Full SKU breakdown with Model, Color, stock levels, velocity, age, sessions, CVR, restock recommendation
4. **⚠️ Action Items** — Priority-sorted list of items needing attention (dead stock, OOS, low stock, slow movers)

---

## File Reference

| File | Purpose | When to read |
|------|---------|-------------|
| `references/report-spec.md` | Full specification: status rules, restock logic, Excel formatting, column definitions | When building manually or customizing |
| `references/lead-times.md` | Editable lead time configuration and restock formula details | When user wants to change lead times |
| `scripts/generate_report.py` | Main report generator script | Always — run this first |

---

## Text Summary Template

After generating the report, provide a summary like:

```
📦 Inventory Health Report — [date]

Overview: [X] SKUs | [Y] units available | [Z] inbound

Status:
- 🟢 Healthy: [n]
- 🔵 Low Stock: [n]
- 🟡 Slow: [n]
- 🔴 Dead Stock: [n]
- ⚫ OOS: [n]

🚨 Urgent Actions:
- [SKU/Model] — OOS, was selling X/mo → restock ASAP
- [SKU/Model] — X days left, need air freight Y units

📦 Restock Total: [X] units
- 🚨 Air Urgent: [n] SKUs, [x] units
- ✈️ Air: [n] SKUs, [x] units
- 🚢 Sea: [n] SKUs, [x] units
```
