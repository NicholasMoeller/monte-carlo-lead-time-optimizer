# Shared Resource Monte Carlo Simulation - Safe Lead Time Calculator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![VBA](https://img.shields.io/badge/VBA-Excel-green.svg)](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
[![Version](https://img.shields.io/badge/version-1.1.0-blue.svg)](CHANGELOG.md)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](CONTRIBUTING.md)

> Calculate safe lead times for products sharing production capacity using Monte Carlo simulation and the "shared fate" phenomenon

---

## 📋 Table of Contents

- [Portfolio Highlights](#-portfolio-highlights) ⭐ **Start here for recruiters/hiring managers**
- [Overview](#overview)
- [The Problem](#the-problem)
- [The Solution](#the-solution)
- [Features](#features)
- [Quick Start](#quick-start)
- [Excel Workbook Structure](#excel-workbook-structure)
- [Installation](#installation-instructions)
- [Usage](#how-to-run-the-simulation)
- [Understanding Results](#understanding-the-results)
- [Customization](#customization-options)
- [Examples](#examples)
- [Contributing](#contributing)
- [License](#license)
- [Changelog](#changelog)

**📄 Additional Documentation:**
- [PROJECT_OVERVIEW.md](PROJECT_OVERVIEW.md) - One-page portfolio summary
- [ARCHITECTURE.md](ARCHITECTURE.md) - Technical deep-dive and system design
- [CONTRIBUTING.md](CONTRIBUTING.md) - Contribution guidelines
- [CHANGELOG.md](CHANGELOG.md) - Version history

## ✨ Features

- ⚡ **2,000 Monte Carlo simulations** for robust statistical analysis
- 📊 **Dual buffer methodologies** - Equal treatment vs. variance-weighted
- 🎯 **System buffer health indicators** with color-coded capacity alerts
- 📈 **Large order detection** with automatic LT quote calculations
- 📉 **Volatility spaghetti chart** - 100-simulation visualization
- 🔧 **Configurable parameters** - Easy customization via constants
- ✅ **Automatic validation** - Detects overloaded lines and data errors
- 🚀 **Performance optimized** - Completes in seconds for typical datasets

---

## 💼 Portfolio Highlights

This project demonstrates expertise across multiple domains:

### Technical Skills Demonstrated

**Advanced VBA Programming**
- Object-oriented design with modular architecture
- 2,000-iteration Monte Carlo simulation engine
- Dictionary-based data structures for O(n) performance
- Custom statistical functions (Percentile, NormInv)
- Dynamic array manipulation and memory optimization
- Error handling and input validation
- Excel automation and chart generation

**Statistical Analysis & Modeling**
- Normal distribution modeling with Box-Muller transform
- Percentile-based risk analysis (95th percentile)
- Variance-weighted allocation algorithms
- Coefficient of Variation for volatility detection
- Time-series simulation over 365-day horizons

**Software Engineering Best Practices**
- Comprehensive documentation with function headers
- Configurable constants for maintainability
- Semantic versioning (v1.1.0)
- Git workflow with meaningful commit messages
- MIT License for open-source distribution
- Extensive input validation and user-friendly error messages

**Data Visualization**
- 100-simulation spaghetti chart with Excel ChartObjects
- Color-coded health indicators (Green/Orange/Red)
- Multi-series line charts with conditional formatting
- Automated chart generation and worksheet management

### Business Impact

**Problem Solved**: Traditional lead time calculations fail when products share production capacity. This tool models the "shared fate" phenomenon where volatile products affect all items on the same line.

**Quantifiable Results**:
- ⏱️ **Reduces planning time**: From hours of manual calculation to seconds
- 📊 **Improves accuracy**: 2,000 simulations vs. single-point estimates
- 💰 **Enables optimization**: Identifies 47-day lead time reductions
- 🎯 **Risk quantification**: 95th percentile protection against demand spikes
- 📈 **Capacity planning**: System buffer metrics guide investment decisions

### Key Innovations

1. **Variance-Weighted Buffer Allocation**: Penalizes volatile products proportionally, incentivizing demand predictability
2. **Large Order Detection**: Automatically calculates 2-sigma thresholds and extended lead times
3. **System Buffer Health**: Real-time capacity alerts with color-coded indicators
4. **Dual Methodology**: Side-by-side comparison of equal vs. risk-based approaches
5. **Volatility Visualization**: Spaghetti charts prove uncertainty to stakeholders

### Complex Problems Solved

**Challenge 1: Shared Resource Contention**
- Traditional methods model products independently
- **Solution**: Aggregate demand simulation with queue backlog tracking

**Challenge 2: Demand Uncertainty**
- Single-point forecasts hide variability
- **Solution**: Monte Carlo simulation with 2,000 scenarios

**Challenge 3: Volatile Product Penalty**
- Equal treatment allows volatile products to hurt stable ones
- **Solution**: Variance-weighted buffer allocation by standard deviation

**Challenge 4: Overload Detection**
- Users don't know when capacity is mathematically insufficient
- **Solution**: Automatic detection with recommended capacity calculations

**Challenge 5: Stakeholder Buy-In**
- Hard to justify longer lead times without proof
- **Solution**: Visual spaghetti chart showing 100 possible futures

---

## Overview

This VBA macro calculates **Safe Lead Times** for products that share production capacity. It uses Monte Carlo simulation to account for demand variability and the "shared fate" phenomenon where one product's demand spike affects all products on the same production line.

## The Problem

When multiple products (A, B, C) share the same production line:
- If any product spikes in demand, it consumes shared capacity
- This creates a backlog that delays ALL products on that line
- Traditional single-product lead time calculations are insufficient

## The Solution

This macro:
1. Groups products by their shared production line
2. Simulates aggregate demand across all products on each line
3. Calculates a "Line Safety Buffer" using 2,000 Monte Carlo simulations
4. Applies the same Safe Lead Time to every product sharing that line

---

## 🚀 Quick Start

1. **Download** the `SharedResourceMonteCarloSimulation.bas` file
2. **Import** into your Excel workbook (Alt+F11 → File → Import)
3. **Create** two sheets: "Simulation" and "SalesHistory"
4. **Populate** with your product and sales data (see template in `sample_data/`)
5. **Run** the macro: Press Alt+F8 → Select "RunSharedResourceSimulation" → Run
6. **Review** results in columns G-M and check the "Volatility Chart" sheet

See [Installation Instructions](#installation-instructions) for detailed setup.

---

## Excel Workbook Structure

### Sheet 1: "Simulation" (The Interface)

This is your main input/output sheet.

| Column | Header | Description | Input/Output |
|--------|--------|-------------|--------------|
| A | Product | Product name (e.g., "Prod A") | Input |
| B | Line Name | Production line identifier (e.g., "Line 1") | Input |
| C | Avg Demand | Average demand (calculated from history) | Output |
| D | Std Dev | Standard deviation (calculated from history) | Output |
| E | Line Capacity | Daily production capacity (e.g., 10) | Input |
| F | Line Start Backlog | Current backlog for the line | Input |
| G | Queuing LT (Equal) | **Equal Treatment**: Safe LT with uniform buffer | Output |
| H | Varying LT (Risk-Based) | **Risk-Based**: Safe LT with variance-weighted buffer | Output |
| I | Longest Material LT (Cold Start) | Product-specific material lead time floor (days) | Input |
| J | Toxic Threshold (System Buffer) | Spare capacity available (Capacity - Total Avg Demand) | Output |
| K | Max Safe Order Qty | Maximum order size before creating unrecoverable backlog | Output |
| L | Large Order Quantity | Statistical threshold (Avg + 2σ) - 95th percentile spike | Output |
| M | Large Order LT Quote | Recommended lead time for orders exceeding threshold | Output |

**Color Coding for System Health (Columns J & K):**
- 🔴 **RED (Critical)**: System Buffer ≤ 0 - Line is underwater, cannot recover
- 🟠 **ORANGE (Fragile)**: 0 < System Buffer < 2 - Very little spare capacity, high risk
- 🟢 **GREEN (Healthy)**: System Buffer ≥ 2 - Good spare capacity, can handle spikes

**Important Notes:**
- Enter the **same** Line Capacity value for ALL rows on the same line
- Enter the **same** Line Start Backlog for ALL rows on the same line
- **Longest Material LT (Cold Start)** is product-specific (can be different for each product)
- The macro provides **TWO methodologies** side-by-side for comparison

---

## Two Buffer Allocation Methodologies

The macro calculates **two different approaches** to allocating safety buffer:

### **Column G: Queuing LT - Equal Treatment (Option 1)**
- 🟢 **Green background** (normal) / **Pink background** (overload)
- **Uniform Buffer**: Every product on the line gets the SAME buffer
- **Use Case**: When fairness to all customers is the priority
- **Formula**: `Longest Material LT (Cold Start) + Shared Line Buffer`

### **Column H: Varying LT - Risk-Based (Option 2)**
- 🟢 **Green background** (normal) / **Pink background** (overload)
- **Variance-Weighted Buffer**: Volatile products get LARGER buffers
- **Use Case**: When minimizing risk and penalizing unpredictable demand
- **Formula**: `Longest Material LT (Cold Start) + (StdDev / Total StdDev) × Line Buffer`

---

**Example Results:**

Assume Line 1 has total buffer of 180 days from Monte Carlo simulation:

| Product | Longest Material LT | StdDev | Queuing LT (Equal) | Varying LT (Risk-Based) |
|---------|---------------------|--------|--------------------|-----------------------|
| Prod A  | 20                  | 0.12   | 🟢 **Recommended LT 82 days**<br>(Shared: 62) | 🟢 **Recommended LT 24 days**<br>(Risk: 4) |
| Prod B  | 15                  | 0.13   | 🟢 **Recommended LT 77 days**<br>(Shared: 62) | 🟢 **Recommended LT 20 days**<br>(Risk: 5) |
| Prod C  | 25                  | 5.60   | 🟢 **Recommended LT 87 days**<br>(Shared: 62) | 🟢 **Recommended LT 196 days**<br>(Risk: 171) |

**Interpretation:**
- **Equal Treatment**: All products get ~62 days of buffer (same for everyone)
- **Risk-Based**: Product C gets 171 days (high volatility), while A & B get only 4-5 days (stable demand)

**Which to use?**
- Use **Column G (Queuing LT - Equal)** if treating all customers fairly is important
- Use **Column H (Varying LT - Risk-Based)** if managing risk and incentivizing stable demand patterns

---

## System Buffer & Large Order Detection

### **Toxic Threshold (System Buffer) - Column J**

The **System Buffer** shows how much spare capacity your line has before it becomes overloaded:

**Formula:** `System Buffer = Line Capacity - Total Average Demand`

**What it means:**
- This is the "bank account" of spare capacity shared by ALL products on the line
- If System Buffer = 2, any demand spike totaling more than 2 units will create an unrecoverable backlog
- **RED** (≤0): Line is already underwater - lead times will grow indefinitely
- **ORANGE** (<2): Fragile - very vulnerable to demand spikes
- **GREEN** (≥2): Healthy - can absorb moderate demand fluctuations

**Example:**
- Line Capacity: 10 units/day
- Total Avg Demand (A+B+C): 8.5 units/day
- System Buffer: 1.5 units/day 🟠 (ORANGE - Fragile)

### **Max Safe Order Qty - Column K**

The **Max Safe Order Qty** tells you the maximum order size each product can accept without creating a backlog that cannot be recovered:

**Formula:** `Max Safe Order Qty = Product Average Demand + System Buffer`

**What it means:**
- Any order larger than this value will consume the entire system buffer
- Orders exceeding this threshold are "toxic" - they guarantee delayed lead times for ALL products on the line
- This is the **Recovery Time Method** - orders beyond this size create a backlog you cannot clear in one period

**Example:**
- Product A Average: 3 units/day
- System Buffer: 1.5 units/day
- Max Safe Order: 4.5 units 🟠
- **Interpretation**: If Product A receives an order >4.5 units, the line cannot recover, and lead times will extend for Products A, B, AND C (shared fate)

---

## Large Order Detection & LT Quoting

### **Large Order Quantity - Column L**

The **Large Order Quantity** identifies when an order is statistically "large" (a demand spike):

**Formula:** `Large Order Quantity = Avg Demand + (2 × Std Dev)`

**What it means:**
- Uses the **Statistical Method** (2-sigma approach)
- Represents the **95th percentile** of demand distribution
- Any order exceeding this quantity is a statistical anomaly
- Only 5% of orders should be this large or larger

**Example:**
- Product C Average: 9.7 units/day
- Product C Std Dev: 5.7 units/day
- Large Order Quantity: 9.7 + (2 × 5.7) = **21.1 units**
- **Interpretation**: Orders >21 units are "large orders" requiring special LT consideration

### **Large Order LT Quote - Column M**

The **Large Order LT Quote** provides the recommended lead time for orders exceeding the quantity threshold:

**Formula:** `Large Order LT = Material LT + Queue Buffer + (Excess Qty / Capacity)`

Where:
- **Material LT**: Product-specific material procurement time
- **Queue Buffer**: Variance-weighted buffer from Monte Carlo simulation
- **Excess Qty**: (Large Order Quantity - Avg Demand)
- **Extra Days**: Excess Qty ÷ Line Capacity

**What it means:**
- Accounts for **normal queue buffer** (volatility)
- PLUS **extra production time** needed for the large quantity
- Provides realistic lead time that won't create backlogs

**Example:**
- Product C Material LT: 25 days
- Product C Queue Buffer: 171 days (high volatility)
- Large Order Quantity: 21.1 units
- Excess Qty: 21.1 - 9.7 = 11.4 units
- Extra Days: 11.4 ÷ 10 capacity = 1.1 days
- **Large Order LT**: 25 + 171 + 1.1 = **197.1 days**

**When to use:**
- For any order exceeding the **Large Order Quantity**
- Ensures you have capacity to produce the excess quantity
- Prevents creating backlogs that delay other customers

---

### Sheet 2: "SalesHistory" (Raw Data)

This sheet contains historical sales data for statistical analysis.

| Column | Header | Description |
|--------|--------|-------------|
| A | Product Name | Product identifier (must match Sheet 1) |
| B → | Day 1, Day 2, ... | DAILY sales data |

**Example:**

```
Product Name | Day-1 | Day-2 | Day-3 | Day-4 | Day-5 | Day-6 | ...
-------------|-------|-------|-------|-------|-------|-------|----
Prod A       | 8.2   | 9.1   | 7.5   | 8.8   | 10.2  | 8.5   | ...
Prod B       | 2.1   | 2.5   | 1.9   | 2.3   | 2.7   | 2.4   | ...
Prod C       | 4.5   | 12.3  | 3.8   | 15.2  | 5.1   | 3.9   | ...
```

**Important:** Each column represents ONE DAY of sales. You need at least 30-60 days of history.

---

## Units of Measure - IMPORTANT!

All calculations use **DAILY** time periods:

| Field | Units | Example |
|-------|-------|---------|
| **Sales History** | Units per day | "8.2" = 8.2 widgets sold on that day |
| **Avg Demand** | Units per day | "3.1" = average of 3.1 units/day |
| **Line Capacity** | Units per day | "10" = can produce 10 units/day |
| **Line Start Backlog** | Units | "5" = 5 units currently in backlog |
| **Recommended LT** | **DAYS** | "52" = need 52 days of lead time |

**Key Relationship:**
- If Line Capacity = 10 units/day
- And Recommended LT = 52 days
- Then you need: **520 units** of buffer inventory (10 × 52)

**Simulation Horizon:**
- The macro simulates **365 days** (1 year) by default
- Change `NUM_DAYS` constant in code to adjust (e.g., 90 days = 1 quarter)
- Each simulation runs 2,000 iterations of 365-day forecasts

**Your sales history data should be DAILY:**
- Each column represents one day of sales
- You need at least 30-60 days of history for good statistics
- More history = better averages and standard deviations

---

## Installation Instructions

### Step 1: Import the VBA Module

1. Open your Excel workbook
2. Press `Alt + F11` to open the VBA Editor
3. Go to **File → Import File...**
4. Select `SharedResourceMonteCarloSimulation.bas`
5. Click **OK**

### Step 2: Enable Macros

1. Close the VBA Editor
2. Save your workbook as `.xlsm` (Excel Macro-Enabled Workbook)
3. When you open the file, click **Enable Content** in the security warning

### Step 3: Create Your Worksheets

1. Create a sheet named **"Simulation"** with the column headers shown above
2. Create a sheet named **"SalesHistory"** with your historical sales data

---

## How to Run the Simulation

### Method 1: From the VBA Editor
1. Press `Alt + F11` to open the VBA Editor
2. Press `F5` or go to **Run → Run Sub/UserForm**
3. Select `RunSharedResourceSimulation`
4. Click **Run**

### Method 2: Create a Button (Recommended)
1. Go to **Developer tab → Insert → Button (Form Control)**
2. Draw a button on your "Simulation" sheet
3. In the "Assign Macro" dialog, select `RunSharedResourceSimulation`
4. Click **OK** and label your button "Run Simulation"

### Method 3: Keyboard Shortcut
1. In the VBA Editor, go to **Tools → Macros**
2. Select `RunSharedResourceSimulation`
3. Click **Options**
4. Assign a shortcut key (e.g., `Ctrl+Shift+R`)

---

## Understanding the Results

### Normal Results (Green Background)

When the simulation completes successfully, Column G will show:
- **Recommended LT** (e.g., "52.3") in DAYS
- **Green background** indicating a valid result
- **All products on the same line will have the SAME lead time**

**Example:**
```
Product  | Line Name | Recommended LT
---------|-----------|----------------
Prod A   | Line 1    | 52.3 days (green)
Prod B   | Line 1    | 52.3 days (green)
Prod C   | Line 1    | 52.3 days (green)
```

### Critical Overload (Red Background)

If the sum of average demands exceeds line capacity:
- Column G will show **"OVERLOAD: LT=X.X days (Need Y.Y capacity vs Z.Z)"**
- **Red background with white text**
- This line is mathematically unsustainable

**What the message tells you:**
- **LT=X.X days**: What your lead time would be with current overloaded conditions
- **Need Y.Y capacity**: Recommended capacity to handle demand sustainably (avg demand + 10% buffer)
- **vs Z.Z**: Your current capacity

**Example:**
```
OVERLOAD: LT=180 days
(Need 13.2 capacity vs 10)
```
This means: With capacity of 10 units/day, you'd need a 180-day lead time. To fix it, increase capacity to at least 13.2 units/day.

**What to do:**
- Increase line capacity (Column E) to the recommended level
- Remove high-demand products from the line
- Split products across multiple lines

### Error Messages (Yellow Background)

If data is missing or invalid:
- Column G will show error details (e.g., "ERROR: Missing Statistics")
- **Yellow background**

**Common causes:**
- Product not found in SalesHistory sheet
- Invalid or missing capacity value
- Missing sales data

---

## The "What-If" Workflow: Build Your Optimal Portfolio

This macro is designed for iterative portfolio optimization:

### Baseline Scenario
1. Add all your products to the Simulation sheet
2. Run the macro
3. Observe the lead times for each line

**Example Result:**
```
Product  | Line Name | Avg Demand | Std Dev | Capacity | Recommended LT
---------|-----------|------------|---------|----------|---------------
Prod A   | Line 1    | 3.2        | 0.5     | 10       | 62 days
Prod B   | Line 1    | 2.8        | 0.4     | 10       | 62 days
Prod C   | Line 1    | 3.5        | 4.2     | 10       | 62 days (HIGH VOLATILITY!)
```

**Analysis:** Product C has very high volatility (Std Dev = 4.2), driving up the lead time for the entire line.

### Optimization Scenario
1. Delete the row for Product C (or move it to another line)
2. Run the macro again
3. Compare the new lead times

**New Result:**
```
Product  | Line Name | Avg Demand | Std Dev | Capacity | Recommended LT
---------|-----------|------------|---------|----------|---------------
Prod A   | Line 1    | 3.2        | 0.5     | 10       | 21 days
Prod B   | Line 1    | 2.8        | 0.4     | 10       | 21 days
```

**Mathematical Proof:** Removing Product C reduced the lead time from **62 days to 21 days** for Products A and B!

### Decision Framework

Now you can make data-driven decisions:
- Is Product C profitable enough to justify hurting A and B's lead times?
- Should we dedicate a separate line to Product C?
- Can we improve Product C's demand predictability?

---

## Technical Details

### Monte Carlo Simulation Parameters

- **Number of Simulations:** 2,000
- **Forecast Horizon:** 365 days (1 year)
- **Risk Level:** 95th percentile (you can modify this in the code)
- **Demand Distribution:** Normal distribution using mean and standard deviation

### The "Shared Fate" Algorithm

For each simulation:
1. Start with the initial backlog
2. For each day (1-365):
   - Generate random demand for **each product** on the line
   - Sum all demands into `Total_Line_Demand`
   - Update queue: `Queue = Queue + Total_Line_Demand - Capacity`
   - Track the maximum queue reached
3. Record the maximum queue from this simulation
4. Repeat 2,000 times
5. Calculate the 95th percentile of all maximum queues
6. Convert to lead time: `Safe_LT = 95th_Percentile_Queue / Capacity`

### Key Formula

```
Safe Lead Time (days) = PERCENTILE(Max_Queues, 0.95) / Line_Capacity
```

This represents: "How many days of capacity are needed to absorb 95% of worst-case scenarios?"

---

## Code Features

### Performance Optimizations
- `Application.ScreenUpdating = False` - Prevents screen flicker
- `Application.Calculation = xlCalculationManual` - Speeds up processing
- Dictionary-based grouping for O(n) performance

### Error Handling
- Validates worksheet existence
- Checks for valid capacity values
- Detects overloaded lines automatically
- Handles missing sales data gracefully

### Visual Feedback
- **Green cells:** Valid results
- **Red cells:** Critical overload
- **Yellow cells:** Data errors
- Completion message with execution time

---

## Customization Options

You can modify the code to adjust:

### Change the Risk Level
In `ProcessSingleLine`, line:
```vba
percentile95 = Percentile(maxQueues, 0.95)
```
Change `0.95` to:
- `0.90` for 90th percentile (less conservative)
- `0.99` for 99th percentile (more conservative)

### Change Simulation Count
In `ProcessSingleLine`, line:
```vba
Const NUM_SIMULATIONS As Long = 2000
```
Change to:
- `1000` for faster results (less accurate)
- `5000` for more precision (slower)

### Change Forecast Horizon
In `ProcessSingleLine`, line:
```vba
Const NUM_DAYS As Long = 365
```
Change to your desired planning horizon:
- `90` = 1 quarter
- `180` = 6 months
- `365` = 1 year (default)
- `730` = 2 years

---

## Troubleshooting

### "Required worksheets not found"
- Ensure you have sheets named exactly "Simulation" and "SalesHistory"
- Check for extra spaces in sheet names

### "ERROR: Missing Statistics"
- Verify product names match exactly between both sheets (case-sensitive)
- Ensure SalesHistory has at least 2 months of data for each product

### "CRITICAL OVERLOAD"
- Sum of average demands exceeds capacity
- Either increase capacity or remove products from the line

### Macro runs slowly
- Reduce `NUM_SIMULATIONS` from 2000 to 1000
- Reduce `NUM_MONTHS` if you have a shorter planning horizon
- Ensure you're not running other Excel processes

### Results seem incorrect
- Check that Line Capacity and Line Start Backlog are entered correctly
- Verify sales history data is realistic (no negative values, outliers)
- Ensure all products on the same line have the same capacity value

---

## Advanced Usage

### Scenario Analysis Table

Create a separate sheet to track multiple scenarios:

| Scenario | Products on Line 1 | Lead Time | Notes |
|----------|-------------------|-----------|-------|
| Baseline | A, B, C | 62 days | High volatility from C |
| Remove C | A, B | 21 days | Much better! |
| Add capacity | A, B, C | 48 days | Increased capacity to 15 |

### Capacity Planning

Use the macro to determine optimal capacity:
1. Start with current capacity
2. If overloaded, incrementally increase capacity
3. Rerun until lead times are acceptable
4. Calculate ROI of capacity investment vs. improved lead times

---

## 📊 Examples

### Example 1: Basic Setup

**Scenario**: Three products sharing Line 1 with capacity of 10 units/day

| Product | Avg Demand | Std Dev | Material LT |
|---------|-----------|---------|-------------|
| Prod A  | 3.2       | 0.5     | 20 days     |
| Prod B  | 2.8       | 0.4     | 15 days     |
| Prod C  | 3.5       | 4.2     | 25 days     |

**Results**:
- System Buffer: **0.5 units/day** 🟠 (Orange - Fragile)
- Equal Treatment LT: **82 days** (all products get same buffer)
- Risk-Based LT: **196 days** for Prod C (high volatility penalty)

**Insight**: Product C's volatility is driving up lead times for the entire line.

### Example 2: Optimization

**Action**: Move Product C to Line 2 (dedicated capacity)

**New Results for Line 1**:
- System Buffer: **4.0 units/day** 🟢 (Green - Healthy)
- Equal Treatment LT: **35 days** (dramatically improved!)
- Products A and B can now quote much shorter lead times

**Business Impact**: 47-day reduction in lead time by isolating volatile product.

---

## 🤝 Contributing

We welcome contributions! Here's how you can help:

### Reporting Issues
- Check [existing issues](../../issues) first
- Provide Excel version and OS details
- Include sample data (anonymized) if possible
- Describe expected vs. actual behavior

### Suggesting Enhancements
- Open an issue with the `enhancement` label
- Explain the use case and benefits
- Provide examples if applicable

### Submitting Pull Requests
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Make your changes with clear comments
4. Test thoroughly with sample data
5. Update documentation (README, CHANGELOG)
6. Commit with descriptive messages (`git commit -m 'Add feature: ...'`)
7. Push to your branch (`git push origin feature/AmazingFeature`)
8. Open a Pull Request

### Code Standards
- Use meaningful variable names
- Add comments for complex logic
- Follow existing code style
- Update function headers when modifying functions
- Test with various data scenarios

See [CONTRIBUTING.md](CONTRIBUTING.md) for detailed guidelines.

---

## 📝 Support and Feedback

### Getting Help
- 📖 Read the [documentation](#overview)
- 🔍 Check [troubleshooting section](#troubleshooting)
- 💬 Open a [GitHub issue](../../issues)
- 📧 Review code comments in the VBA module

### Feedback
Your feedback helps improve this tool:
- ⭐ Star this repo if you find it useful
- 🐛 Report bugs via GitHub issues
- 💡 Suggest features or improvements
- 📣 Share your success stories

---

## 📜 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.

Copyright (c) 2025 Monte Carlo Lead Time Calculator Team

Permission is granted to use, modify, and distribute this software for commercial and non-commercial purposes.

---

## 📚 Changelog

See [CHANGELOG.md](CHANGELOG.md) for a detailed version history.

**Latest Version**: 1.1.0 (2025-12-05)

---

## 🙏 Acknowledgments

- Built with Excel VBA for maximum accessibility
- Inspired by real-world supply chain challenges
- Community-driven development and improvements

---

**Made with ❤️ for supply chain professionals and operations managers**
