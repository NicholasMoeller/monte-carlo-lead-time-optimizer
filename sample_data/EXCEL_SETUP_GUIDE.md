# Excel Workbook Setup Guide

## Step-by-Step Setup (5 minutes)

### Part 1: Create the Workbook Structure

1. **Open Excel** and create a new blank workbook

2. **Create Sheet 1: "SalesHistory"**
   - Right-click on "Sheet1" → Rename to **"SalesHistory"**
   - Go to: Data tab → Get Data → From File → From Text/CSV
   - Select `SalesHistory.csv`
   - Click "Load"
   - Your sheet should look like this:

   ```
   Product Name | Jan-2024 | Feb-2024 | Mar-2024 | ...
   -------------|----------|----------|----------|----
   Product A    | 3.2      | 3.0      | 3.3      | ...
   Product B    | 2.5      | 2.3      | 2.7      | ...
   ...
   ```

3. **Create Sheet 2: "Simulation"**
   - Right-click on sheet tabs → Insert → Worksheet
   - Rename to **"Simulation"**
   - Go to: Data tab → Get Data → From File → From Text/CSV
   - Select `Simulation.csv`
   - Click "Load"
   - Your sheet should look like this:

   ```
   Product   | Line Name | Avg Demand | Std Dev | Line Capacity | Line Start Backlog | Recommended LT
   ----------|-----------|------------|---------|---------------|-------------------|---------------
   Product A | Line 1    |            |         | 10            | 5                 |
   Product B | Line 1    |            |         | 10            | 5                 |
   ...
   ```

4. **Verify Sheet Names**
   - Look at the bottom tabs
   - You should see: "Simulation" and "SalesHistory" (exact spelling, no spaces)

### Part 2: Import the VBA Macro

5. **Open VBA Editor**
   - Press `Alt + F11` (Windows) or `Fn + Option + F11` (Mac)
   - You should see the VBA Editor window

6. **Import the Module**
   - In VBA Editor: File → Import File
   - Navigate to `SharedResourceMonteCarloSimulation.bas`
   - Click "Open"
   - You should see "SharedResourceMonteCarloSimulation" appear in the Project Explorer

7. **Verify Import**
   - In the Project Explorer (left panel), expand "Modules"
   - Double-click "SharedResourceMonteCarloSimulation"
   - You should see the VBA code

8. **Close VBA Editor**
   - Close the window to return to Excel

### Part 3: Save as Macro-Enabled Workbook

9. **Save the File**
   - File → Save As
   - File name: `Monte_Carlo_Test.xlsm`
   - Save as type: **Excel Macro-Enabled Workbook (*.xlsm)**
   - Click "Save"

10. **Enable Macros** (if prompted)
    - Close and reopen the file
    - If you see a security warning: Click "Enable Content"

### Part 4: Run the Simulation

11. **Run the Macro**
    - Press `Alt + F8` (opens Macro dialog)
    - Select `RunSharedResourceSimulation`
    - Click **"Run"**

12. **Watch the Magic**
    - The macro should complete in 1-3 seconds
    - You'll see a success message: "Simulation completed successfully..."
    - Click "OK"

13. **Check the Results**
    - Go to the "Simulation" sheet
    - Columns C, D, and G should now be filled
    - Column G cells should have a **green background**
    - All products on Line 1 should show the SAME lead time

### Part 5: Create a Run Button (Optional but Recommended)

14. **Enable Developer Tab** (if not visible)
    - File → Options → Customize Ribbon
    - Check "Developer" in the right panel
    - Click "OK"

15. **Insert a Button**
    - Go to the "Simulation" sheet
    - Developer tab → Insert → Button (Form Control)
    - Draw a button somewhere (e.g., cell I2)

16. **Assign the Macro**
    - When the "Assign Macro" dialog appears
    - Select `RunSharedResourceSimulation`
    - Click "OK"

17. **Label the Button**
    - Right-click the button → Edit Text
    - Type: **"Run Simulation"**
    - Click outside the button

18. **Test the Button**
    - Click the button
    - Simulation should run!

---

## Quick Verification Checklist

After setup, verify:

- [ ] Two sheets exist: "Simulation" and "SalesHistory"
- [ ] Sheet names are spelled exactly correctly
- [ ] SalesHistory has 8 products (A through H) with 18 months of data
- [ ] Simulation has 8 products with Line Names and Capacity filled in
- [ ] File is saved as `.xlsm` (macro-enabled)
- [ ] VBA module appears in the Modules folder
- [ ] Macro runs without errors
- [ ] Results appear in columns C, D, and G
- [ ] Green background appears in column G

---

## Expected Results After Running

Your Simulation sheet should show:

### Line 1 (HIGH lead time due to Product C volatility)
- Product A: Avg ~3.1, StdDev ~0.15, **LT ~6-7 months** (green)
- Product B: Avg ~2.5, StdDev ~0.17, **LT ~6-7 months** (green)
- Product C: Avg ~9.5, StdDev ~6.20, **LT ~6-7 months** (green)

### Line 2 (LOW lead time - excess capacity)
- Product D: Avg ~5.9, StdDev ~0.23, **LT ~2-3 months** (green)
- Product E: Avg ~4.3, StdDev ~0.15, **LT ~2-3 months** (green)

### Line 3 (MODERATE lead time - balanced)
- Product F: Avg ~2.2, StdDev ~0.12, **LT ~3-4 months** (green)
- Product G: Avg ~1.8, StdDev ~0.08, **LT ~3-4 months** (green)
- Product H: Avg ~2.4, StdDev ~0.11, **LT ~3-4 months** (green)

**Key Insight**: Notice Line 1 has the longest lead time? That's because Product C is extremely volatile (StdDev 6.2). This is the "shared fate" in action!

---

## Try the "What-If" Test

1. **Delete the Product C row** (Line 1 becomes less volatile)
2. **Run the macro again**
3. **Observe**: Line 1 lead time drops from ~6-7 months to ~2-3 months!
4. **Conclusion**: Product C is costing you 4 months of lead time. Is it worth it?

---

## Troubleshooting

### "Compile Error: Can't find project or library"
- Tools → References → Uncheck any MISSING references
- Ensure "Microsoft Scripting Runtime" is checked

### "Run-time error '9': Subscript out of range"
- Check sheet names are EXACTLY "Simulation" and "SalesHistory"
- No extra spaces or typos

### Results don't appear
- Check that columns C, D, G are not protected
- Ensure macros are enabled (check security warning)

### "Type mismatch" error
- Ensure capacity values in column E are numbers (not text)
- Check that all sales history data is numeric

---

## Alternative: Copy-Paste Method

If CSV import doesn't work:

1. Open `SalesHistory.csv` in Excel
2. Select all data (Ctrl+A)
3. Copy (Ctrl+C)
4. Go to your workbook → "SalesHistory" sheet
5. Paste (Ctrl+V)
6. Repeat for Simulation.csv

---

## You're Ready!

Once you see green results in column G, you've successfully set up the Monte Carlo simulation. Now you can:
- Test with your own data
- Run "What-If" scenarios
- Optimize your product portfolio

**Happy simulating!**
