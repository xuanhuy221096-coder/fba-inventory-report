# Lead Time Configuration

Configurable parameters for restock calculations. These are defaults — the user can override any value.

---

## Production & Shipping

| Parameter | Variable | Default | Range | Notes |
|-----------|----------|---------|-------|-------|
| Production + Labeling | `PREP_DAYS` | 6 | 5-6 | Time to manufacture and label products |
| Sea Freight | `SEA_DAYS` | 30 | 25-30 | China → USA ocean shipping |
| Air Freight | `AIR_DAYS` | 15 | 10-15 | China → USA air shipping |
| Amazon Check-in | `AMAZON_CHECKIN` | 7 | 5-10 | Time for FBA to receive, stow, and make available |

## Derived Totals

| Parameter | Formula | Default |
|-----------|---------|---------|
| Total Lead (Sea) | PREP + SEA + CHECKIN | **43 days** |
| Total Lead (Air) | PREP + AIR + CHECKIN | **28 days** |

## Inventory Planning

| Parameter | Variable | Default | Notes |
|-----------|----------|---------|-------|
| Safety Stock | `SAFETY_DAYS` | 14 | Buffer to prevent stockout |
| Restock Target | `TARGET_COVER_DAYS` | 60 | Days of cover to have after shipment arrives |
| Min Order Qty | `MIN_ORDER` | 50 | Supplier minimum order per SKU |

---

## Restock Decision Tree

```
                    ┌─ DEAD STOCK? ──→ ❌ No Restock
                    │
                    ├─ vel ≈ 0 & t90 ≤ 5? ──→ ❌ No Restock
                    │
total_days_of_supply ┤
                    ├─ > SEA_TOTAL + SAFETY (57d) ──→ 🚢 Sea Freight
                    │
                    ├─ > AIR_TOTAL + SAFETY (42d) ──→ ✈️ Air Freight
                    │
                    └─ ≤ 42d ──→ 🚨 AIR URGENT
```

## How to override

User can say things like:
- "tàu đi mất 35 ngày" → set SEA_DAYS = 35
- "min order 100" → set MIN_ORDER = 100
- "safety stock 21 ngày" → set SAFETY_DAYS = 21
- "target cover 90 ngày" → set TARGET_COVER_DAYS = 90

Pass overrides as arguments to the script or adjust in the manual build.
