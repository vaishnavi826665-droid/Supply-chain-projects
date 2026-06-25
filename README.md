# Supply-chain-projects


# Supply Chain Projects — Sports Equipment (DTC)

Three hands-on supply chain analysis projects built using data from **ProGear Direct**, a fictional direct-to-consumer sports equipment company that ships basketballs, dumbbells, lacrosse sticks, and other gear across the US. The projects demonstrate core competencies in inventory management, purchase order tracking, fulfillment analytics, and operational problem-solving.

---

## Project 1 — Multi-Warehouse Inventory Rebalancer

**File:** `project1_inventory_rebalancer/ProGear_Direct_Inventory_Rebalancer.xlsx`

### Overview

An Excel dashboard for monitoring 10 sports equipment SKUs across three regional warehouses (Northeast, Southeast, West Coast) and flagging stock imbalances before peak back-to-school and fall seasons. The model surfaces transfer recommendations so the right inventory is in the right place when demand spikes.

### Sheets

| Sheet | Purpose |
|---|---|
| **Dashboard** | Live status view with color-coded flags (✅ OK / ⚠️ LOW / 🔴 CRITICAL) per SKU and warehouse |
| **Transfer Recommendations** | Prioritized transfer plan with estimated transit days and reasoning |
| **Raw Data** | Editable input table — update on-hand quantities here and the Dashboard refreshes automatically |

### Key Features

- Status logic flags any warehouse below 50% of minimum stock as CRITICAL
- Weeks-of-supply calculated from current on-hand vs. weekly demand
- Transfer recommendations color-coded by priority (HIGH / MEDIUM / LOW)
- Blue cells = editable inputs; black cells = formula outputs

---

## Project 2 — Order Fulfillment Cycle Time Analysis

**File:** `project2_fulfillment_analysis/ProGear_Direct_Fulfillment_Analysis.xlsx`

### Overview

A cycle time analysis of 200 eCommerce orders processed Aug–Oct 2024. Maps the full fulfillment pipeline — order placed → picked → packed → shipped → delivered — and identifies where delays accumulate, which carriers perform best, and where operational improvements would have the highest impact.

### Sheets

| Sheet | Purpose |
|---|---|
| **Summary Dashboard** | KPI cards, warehouse performance table, carrier comparison, and key findings |
| **Order Data** | Full 200-row order log with dates, per-stage lags, and on-time flags |
| **Stage Breakdown** | Per-stage average, median, and max lag with bottleneck identification |

### Key Findings

- Shipping lag (packed → shipped) is the largest contributor to total cycle time
- FedEx Ground delivers the lowest average cycle time across all five carriers
- Pick lag spikes during back-to-school surge in late August — a seasonal staffing opportunity

---

## Project 3 — Purchase Order & Inbound Shipment Tracker

**File:** `project3_po_tracker/ProGear_Direct_PO_Inbound_Tracker.xlsx`

### Overview

An end-to-end PO and inbound shipment tracker covering 30 purchase orders across 5 suppliers and 3 warehouses. Tracks every order from PO creation through supplier shipment, warehouse receipt, and 3-way match (PO → receipt → expected qty). Flags discrepancies automatically and generates a dedicated follow-up report for the supply chain team.

### Sheets

| Sheet | Purpose |
|---|---|
| **PO Dashboard** | KPI summary, supplier on-time performance table, and PO status breakdown |
| **PO Log** | Full log of all 30 POs with order dates, expected vs. actual delivery, and match status |
| **Receipt Log** | Warehouse receipt records matched back to POs — variance flagged in red |
| **Discrepancy Report** | Isolated view of all mismatches with issue type, recommended supplier action, and resolution status |

### Key Features

- Automatic 3-way match: PO quantity vs. receipt quantity vs. expected delivery date
- Discrepancy types tracked: quantity short, wrong SKU delivered
- Supplier scorecard showing on-time rate and discrepancy count per vendor
- Color-coded PO log: green = delivered & matched, blue = in transit, red = discrepancy
- Discrepancy Report pre-populated with recommended supplier contact actions

---

## Rebuilding from Source

All three Excel files are generated from Python scripts:

```bash
pip install openpyxl pandas

python project1_inventory_rebalancer/build_inventory_tracker.py
python project2_fulfillment_analysis/build_fulfillment_analysis.py
python project3_po_tracker/build_po_tracker.py
```

---

## Skills Demonstrated

- Multi-location inventory monitoring and rebalancing logic
- Purchase order creation, tracking, and 3-way matching
- Inbound shipment tracking and discrepancy resolution
- eCommerce order fulfillment cycle time analysis
- Supplier performance scorecarding
- Excel dashboard design (conditional formatting, status flags, cross-sheet formulas)
- Data simulation and aggregation with pandas
- Translating data into actionable operational recommendations
