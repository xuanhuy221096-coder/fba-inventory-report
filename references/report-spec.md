# Report Specification — FBA Inventory Health Report

Complete specification for building the inventory report manually or customizing the output.

---

## Table of Contents

1. [Data Processing](#data-processing)
2. [Model & Color Parsing](#model--color-parsing)
3. [Derived Metrics](#derived-metrics)
4. [Status Assignment](#status-assignment)
5. [Restock Logic](#restock-logic)
6. [Excel Tab Specs](#excel-tab-specs)
7. [Formatting Standards](#formatting-standards)

---

## Data Processing

### Step 1: Read and validate

```python
import pandas as pd
inv = pd.read_csv('fba_inventory.csv')
biz = pd.read_csv('business_report.csv')
```

Validate FBA Inventory has: `sku`, `asin`, `product-name`, `available`, `units-shipped-t30`
Validate Business Report has: `(Child) ASIN`, `Sessions - Total`, `Units Ordered`

### Step 2: Clean Business Report numerics

Strip `$`, `,`, `%` from these columns before converting to numeric:
- `Units Ordered`, `Sessions - Total`, `Page Views - Total` → int
- `Ordered Product Sales` → float (strip `$` and `,`)
- `Unit Session Percentage` → float (strip `%`)

### Step 3: Clean FBA Inventory numerics

Convert all measurement columns to numeric with `errors='coerce'`, fill NaN with 0:
`available`, `inbound-quantity`, `units-shipped-t7/t30/t60/t90`, `estimated-excess-quantity`, `sell-through`, `weeks-of-cover-t30/t90`, `sales-rank`, `inv-age-*`, `estimated-storage-cost-next-month`, `Total Reserved Quantity`, `unfulfillable-quantity`, `sales-shipped-last-*-days`, `your-price`, `inbound-working`, `inbound-shipped`, `inbound-received`

### Step 4: Merge

```python
merged = inv.merge(biz, left_on='asin', right_on='(Child) ASIN', how='left', suffixes=('_inv','_biz'))
```

---

## Model & Color Parsing

Extract two columns from `product-name`:

### Color
Use regex to find text in parentheses at end of name:
```python
re.search(r'\(([^)]+)\)\s*$', name)
```
Fallback: check for `, ColorName` at end of string against known color list.

### Model
Check keywords in order of specificity (first match wins):

| Check (in order) | Model value |
|-------------------|-------------|
| `Super Duty` or `F250` or `F350` or `F450` | F250/F350/F450 Super Duty |
| `F150` + `Bronco` + `Tundra` | F150 / Bronco / Tundra |
| `F150` + `Bronco` | F150 / Bronco |
| `F150` | F150 |
| `Silverado 2500` | Silverado 2500/3500 HD |
| `Extended Tow Hook` + `Silverado 1500` | Silverado 1500 Extended |
| `Silverado 1500` | Silverado 1500 |
| `Sierra 1500` | Sierra 1500 |
| `Ram 2500` or `Ram 2500/3500` | Ram 2500/3500 |
| `Ram 1500` | Ram 1500 |
| `Colorado` or `Canyon` | Colorado / Canyon |
| `Tire Valve` | Tire Valve Stem Caps |
| Otherwise | Other |

**Adapting to other product lines:** If products are NOT tow hooks / car accessories, adapt the parsing logic. The goal is always to split the product name into a **category/model** column and a **variant/color** column.

---

## Derived Metrics

```python
# Weighted-blend daily velocity (evergreen products).
# Blends 3 lookback windows so a single unusual week/month can't distort the figure.
v7  = units-shipped-t7  / 7
v30 = units-shipped-t30 / 30
v90 = units-shipped-t90 / 90
daily_velocity = 0.5*v30 + 0.3*v7 + 0.2*v90

# Coefficient of variation across the 3 windows = how erratic this SKU sells.
# Higher CV → bigger safety buffer (see Restock Logic).
vel_cv = pstdev([v7, v30, v90]) / daily_velocity   # 0 if velocity == 0

total_fba = available + Reserved FC Transfer + Reserved Customer Order + Reserved FC Processing + Reserved Staging + unfulfillable-quantity
# NOTE: do NOT use the "Total Reserved Quantity" column — it is unreliable and
# routinely undercounts reserved units (esp. FC Transfer / FC Processing).
# Sum the 4 reserved sub-columns instead.
days_remaining = total_fba / daily_velocity  # if velocity > 0, else 9999
aged_181plus = inv-age-181-to-270-days + inv-age-271-to-365-days + inv-age-366-to-455-days + inv-age-456-plus-days
```

**Weight rationale:** v30 (50%) is the backbone — long enough to smooth, recent enough to matter. v7 (30%) catches current trend (accelerating/decelerating). v90 (20%) anchors to the long-run baseline and resists noise. Adjust weights if the product is highly stable (raise v90) or in an aggressive scaling phase (raise v7).

---

## Status Assignment

Apply rules in this order (first match wins).

**IMPORTANT — OOS and LOW STOCK are checked BEFORE SLOW.** A stocked-out SKU can carry a stale high `days_remaining` artifact; if SLOW were checked first it would mask the real OOS/LOW state. Stockout/low always wins over slow-mover.

| Priority | Condition | Status |
|----------|-----------|--------|
| 1 | `no-sale-last-6-months == 'Y'` OR (`units-shipped-t90 == 0` AND `total_fba > 0`) | 🔴 DEAD STOCK |
| 2 | `total_fba <= 0` AND `inbound-quantity <= 0` | ⚫ OOS |
| 3 | `days_remaining < 14` | 🔵 LOW STOCK |
| 4 | `days_remaining > 180` | 🟡 SLOW - Clear |
| 5 | `days_remaining > 90` | 🟡 SLOW |
| 6 | Otherwise | 🟢 HEALTHY |

Note: OOS and DEAD use `total_fba` (on-hand incl. reserved), not bare `available`, so units locked in FC Transfer/Customer Order don't get misread as zero stock.

---

## Restock Logic

See `references/lead-times.md` for configurable parameters.

### Algorithm

```
total_supply = total_fba + inbound-quantity
total_days_of_supply = total_supply / daily_velocity

# Dynamic safety stock: scales with how erratic the SKU sells.
# Steady SKUs get the base buffer; erratic ones (high CV) get more.
safety_units = daily_velocity * SAFETY_DAYS * (1 + vel_cv)

1. If DEAD STOCK → ❌ No Restock (remove/liquidate instead)
2. If velocity ≈ 0 and t90 sales ≤ 5 → ❌ No Restock (too slow)
3. If velocity ≈ 0 but t90 > 5 → use fallback velocity: units-shipped-t90 / 90

Then check timing (SEA_TOTAL = prep+sea+checkin, AIR_TOTAL = prep+air+checkin):

REORDER_POINT = SEA_TOTAL + SAFETY_DAYS
# The gate: an order is only triggered when coverage FALLS TO the reorder
# point. Above it, do nothing — ordering earlier just parks cash in inventory.

4. If total_days > REORDER_POINT → ✅ OK, no restock
   (note: "Covered {total_days}d, reorder at ≤{REORDER_POINT}d")
5. Elif total_days > AIR_TOTAL + SAFETY_DAYS → 🚢 Sea
   qty = (SEA_TOTAL + TARGET_COVER) × vel + safety_units − total_supply
   (if qty ≤ 0 → ✅ OK)
6. Elif total_days > AIR_TOTAL → ✈️ Air (sea can no longer arrive in time)
   qty = (SEA_TOTAL + TARGET_COVER) × vel + safety_units − total_supply
7. Else → 🚨 AIR URGENT
   qty = (AIR_TOTAL + TARGET_COVER) × vel + safety_units − total_supply

Order qty is an order-up-to target: cover the lead time + TARGET_COVER days of
demand, plus safety, minus what's already on hand and inbound. The trigger
(step 4) and the quantity are separate concepts — do not merge them.

# HISTORY (2026-07-19): a previous version of this spec had no step-4 gate;
# the "> SEA_TOTAL + SAFETY" branch itself placed the sea order, so every SKU
# under ~(SEA_TOTAL + TARGET_COVER + safety) days of cover got an order —
# SKUs with 87–100d coverage were being told to restock. Keep the gate.

Apply MIN_ORDER: if qty > 0 and qty < MIN_ORDER, round up to MIN_ORDER
```

---

## Excel Tab Specs

### Tab 1: Dashboard

| Row | Content |
|-----|---------|
| 1 | Title: "📦 INVENTORY HEALTH REPORT" (white on #1A5276) |
| 2 | Subtitle: snapshot date, total SKUs, available, inbound |
| 4-12 | Lead Time Assumptions table (white on #2980B9 header) |
| 14-15 | Metrics: Revenue 30d, Units Shipped 30d, Daily Velocity, Storage Cost |
| 17+ | Status breakdown (count per status) |
| Next | Restock summary (count per restock action) |
| Next | Revenue by Model (30d, descending) |

### Tab 2: 📦 Restock Order

Only SKUs where `Restock Qty > 0`.

| Row | Content |
|-----|---------|
| 1 | Title: "📦 RESTOCK ORDER LIST" (white on #1A5276) |
| 2 | Summary: count + units by urgency tier |
| 3 | Lead time reference line (italic, gray) |
| 5 | Headers (white on #2C3E50) |
| 6+ | Data rows, sorted: 🚨 AIR URGENT → ✈️ Air → 🚢 Sea |
| Last | Total row (yellow background #F7DC6F) |

**Columns:** #, Urgency, Model, Color, SKU, Ship Method, Order Qty, Current Stock, Daily Vel., Days Left, Note

**Row colors:** 🚨 = #F5B7B1 (light red), ✈️ = #D6EAF8 (light blue), 🚢 = #D5F5E3 (light green)

### Tab 3: Inventory Detail

All SKUs, sorted by Model → Color.

**Columns:** Status, Model, Color, SKU, ASIN, Price, Available, Reserved, Inbound, Total FBA, Sold 7d, Sold 30d, Sold 90d, Daily Vel., Days Left, Sell-Through, Sessions, Units Ordered, CVR%, Revenue 30d, Age 0-90, Age 91-180, Age 181+, Storage Cost, Sales Rank, Health, Restock Action, Ship Method, Restock Qty, Restock Note

**Row colors by status:**
- 🔴 DEAD STOCK: #FADBD8
- ⚫ OOS: #D5D8DC
- 🔵 LOW STOCK: #D6EAF8
- 🟡 SLOW - Clear: #FDEBD0
- 🟡 SLOW: #FEF9E7
- 🟢 HEALTHY: #D5F5E3

**Last 4 columns** (Restock) get red header (#E74C3C) to stand out.

### Tab 4: ⚠️ Action Items

Priority-sorted action list.

**Columns:** Priority, Model, Color, SKU, Issue, Available, Ship Method, Recommendation

**Sort order:** 🔴 URGENT (dead) → ⚫ URGENT (OOS) → 🔵 HIGH (low stock) → 🚨 RESTOCK (air urgent) → 🟡 SLOW

---

## Formatting Standards

- **Font:** Arial, size 9-10 for data, 14-16 for titles
- **Headers:** white text (#FFFFFF) on dark background (#2C3E50)
- **Borders:** thin borders on all data cells (color #D5D8DC)
- **Currency:** `$#,##0.00`
- **Numeric columns:** center-aligned
- **Column widths:** set manually per column (not auto-fit)
- **Freeze panes:** header rows frozen on all data tabs
- **Auto-filter:** enabled on header rows

---

## Self-Check (mandatory before saving)

`generate_report.py` runs `_verify_report(merged)` right before `wb.save()`. It raises `AssertionError` and aborts if any rule fails, so a broken report never reaches the user silently:

1. **On-hand identity** — `total_fba == available + reserved_total + unfulfillable` for every row (tolerance 0.5).
2. **No negatives** — `available`, `total_fba`, `inbound-quantity`, `reserved_total` all ≥ 0.
3. **OOS sanity** — every ⚫ OOS row has `total_fba <= 0` AND `inbound-quantity <= 0`.
4. **Status sanity** — no row with stock + velocity is labeled OOS.
5. **Restock completeness** — every row with `Restock Qty > 0` has a non-`N/A` Ship Method.

On pass it prints `✅ Self-check passed (N SKUs verified)`. If it fails, fix the logic — do NOT bypass the check.
