# Sample Test Data for Monte Carlo Simulation

## Files Included

1. **Simulation.csv** - The main interface sheet (import to "Simulation" sheet)
2. **SalesHistory.csv** - Historical sales data (import to "SalesHistory" sheet)

## How to Use These Files

### Quick Import to Excel:

1. Open a new Excel workbook
2. Import `SalesHistory.csv`:
   - Data → From Text/CSV → Select `SalesHistory.csv`
   - Name the sheet **"SalesHistory"** (exact spelling)
3. Import `Simulation.csv`:
   - Create a new sheet, name it **"Simulation"**
   - Data → From Text/CSV → Select `Simulation.csv`
4. Import the VBA macro (`SharedResourceMonteCarloSimulation.bas`)
5. Run the macro!

## What This Data Tests

**Data Characteristics:**
- **90 days** of DAILY sales history for each product
- Each column = 1 day of sales
- Realistic daily demand patterns with volatility

### Line 1: The "Volatile Product" Scenario
- **Product A**: Stable demand (Avg ~3.1, StdDev ~0.15)
- **Product B**: Stable demand (Avg ~2.5, StdDev ~0.17)
- **Product C**: HIGHLY VOLATILE (Avg ~9.5, StdDev ~5.2) ⚠️
- **Line Capacity**: 10 units/DAY
- **Starting Backlog**: 5 units

**Expected Result**: Product C's extreme volatility (swings from 2.8 to 18.5 per day) will drive up the safe lead time for ALL products on Line 1, even though A and B are stable.

**What-If Test**:
- Baseline run → Expect lead time ~150-200 days
- Delete Product C row and rerun → Lead time should drop to ~60-90 days
- **This proves the "shared fate" concept!**

### Line 2: The "Moderate Demand" Scenario
- **Product D**: Moderate demand (Avg ~5.9, StdDev ~0.23)
- **Product E**: Moderate demand (Avg ~4.3, StdDev ~0.15)
- **Line Capacity**: 15 units/DAY
- **Starting Backlog**: 0 units
- **Total Avg Demand**: ~10.2 (68% utilization)

**Expected Result**: Should show reasonable lead times (60-90 days) because:
- Capacity is higher (15 vs 10)
- No backlog to start
- Both products are stable
- Total demand is well below capacity

### Line 3: The "Balanced Line" Scenario
- **Product F**: Low demand (Avg ~2.2, StdDev ~0.12)
- **Product G**: Low demand (Avg ~1.8, StdDev ~0.08)
- **Product H**: Low demand (Avg ~2.4, StdDev ~0.11)
- **Line Capacity**: 8 units/DAY
- **Starting Backlog**: 2 units
- **Total Avg Demand**: ~6.4 (80% utilization)

**Expected Result**: Moderate lead times (90-120 days) because:
- High utilization (80%)
- Small starting backlog
- All products are stable

## Expected Macro Output

After running the macro, your Simulation sheet should look like this:

```
Product   | Line | Avg  | StdDev | Cap | Backlog | Recommended LT | Material LT
----------|------|------|--------|-----|---------|----------------|------------
Product A | L1   | 3.1  | 0.15   | 10  | 5       | ~200 days      | 20
Product B | L1   | 2.5  | 0.17   | 10  | 5       | ~195 days      | 15
Product C | L1   | 9.5  | 5.20   | 10  | 5       | ~205 days      | 25
Product D | L2   | 5.9  | 0.23   | 15  | 0       | ~80 days       | 10
Product E | L2   | 4.3  | 0.15   | 15  | 0       | ~82 days       | 12
Product F | L3   | 2.2  | 0.12   | 8   | 2       | ~113 days      | 8
Product G | L3   | 1.8  | 0.08   | 8   | 2       | ~110 days      | 5
Product H | L3   | 2.4  | 0.11   | 8   | 2       | ~115 days      | 10
```

**Key Insight:**
- **Line 1 Buffer**: ~180 days (calculated from Monte Carlo)
  - Product A: 20 + 180 = **200 days**
  - Product B: 15 + 180 = **195 days**
  - Product C: 25 + 180 = **205 days**
- Each product has a **different total lead time** based on its Material LT
- All products share the **same line buffer** (180 days for Line 1)

Note: Actual buffer values will vary slightly due to Monte Carlo randomness (±15%)

## Key Observations

1. **All products on the same line share the SAME buffer** (but have different total LTs due to Material LT) ✓
2. **Line 1 has the longest buffer** (~180 days) due to Product C's volatility
3. **Line 2 has the shortest buffer** (~70 days) due to excess capacity
4. **Line 3 has moderate buffer** (~105 days) with balanced utilization
5. **Each product's total LT is unique** = Material LT + Shared Line Buffer

## Testing the "What-If" Workflow

### Experiment 1: Remove Volatile Product
1. Delete the "Product C" row
2. Rerun the macro
3. Observe: Line 1 lead time should drop from ~180 to ~75 days
4. **Business Decision**: Is Product C worth a 105-day lead time penalty?

### Experiment 2: Increase Capacity
1. Restore Product C
2. Change Line 1 capacity from 10 → 15 (for all three rows)
3. Rerun the macro
4. Observe: Lead time should drop to ~90-120 days
5. **Business Decision**: Is capacity expansion worth the investment?

### Experiment 3: Move Product to Another Line
1. Change Product C's "Line Name" from "Line 1" → "Line 2"
2. Update Line 2 capacity to handle the extra load (e.g., 15 → 20)
3. Rerun the macro
4. Observe: Line 1 improves, but Line 2 may worsen
5. **Business Decision**: Where should volatile products go?

### Experiment 4: Test Overload
1. Change Line 3 capacity from 8 → 5
2. Rerun the macro
3. Expected: **"OVERLOAD: LT=XXX days (Need 7.0 capacity vs 5)"** in red (demand 6.4 > capacity 5)
4. This demonstrates the overload detection feature with recommended capacity

## Data Characteristics

### Product C (The Troublemaker)
- Swings from 2.8 to 18.5 units per DAY
- Coefficient of Variation (CV) = 55% (extremely high!)
- Represents: Seasonal product, promotional item, or unpredictable demand

### Products A, B, D-H (Stable Products)
- CV = 5-8% (very predictable)
- Represents: Steady-state demand, mature products

### Realistic Scenario
This mirrors real manufacturing where:
- One volatile product drags down the entire line's performance
- You must decide: Isolate it, increase capacity, or discontinue it

## Validation Checks

After running, verify:
- [ ] Columns C & D are filled (Avg and StdDev calculated)
- [ ] Column G shows lead times in months
- [ ] All products on Line 1 have identical lead times
- [ ] All products on Line 2 have identical lead times
- [ ] All products on Line 3 have identical lead times
- [ ] Cells have green background (success)
- [ ] No red "CRITICAL OVERLOAD" messages
- [ ] No yellow "ERROR" messages

## Troubleshooting

**"ERROR: Missing Statistics"**
- Check that product names match EXACTLY between sheets (case-sensitive)
- Ensure no extra spaces

**Results seem wrong**
- The Monte Carlo simulation includes randomness (±10% variance is normal)
- Rerun several times to see the range

**Macro doesn't run**
- Ensure sheets are named exactly "Simulation" and "SalesHistory"
- Check that macros are enabled

## Next Steps

Once the macro works with sample data:
1. Replace with your real sales history data
2. Adjust capacity values to match your production lines
3. Use the "What-If" workflow to optimize your portfolio
4. Document your findings in a scenario comparison table

---

**Enjoy testing! This data is designed to clearly demonstrate the "shared fate" concept.**
