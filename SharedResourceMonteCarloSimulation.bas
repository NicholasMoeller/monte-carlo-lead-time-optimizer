Attribute VB_Name = "SharedResourceMonteCarloSimulation"
Option Explicit

' ==================================================================================
' Shared Resource Monte Carlo Simulation - Safe Lead Time Calculator
' ==================================================================================
' Purpose:  Calculate safe lead times for products sharing production capacity
'           using Monte Carlo simulation to model demand variability and the
'           "shared fate" phenomenon where volatile products affect entire lines
'
' Version:  1.1.0
' Author:   Monte Carlo Lead Time Calculator Team
' License:  MIT License
' Repository: https://github.com/NicholasMoeller/Work
' Date:     2025-12-05
'
' Key Features:
'   - 2,000 Monte Carlo simulations for robust statistical analysis
'   - Variance-weighted buffer allocation (penalizes volatile products)
'   - System buffer health indicators (capacity planning)
'   - Large order detection and LT quoting
'   - Volatility visualization with 100-simulation spaghetti chart
'
' ==================================================================================

' ==================================================================================
' CONFIGURATION CONSTANTS - Modify these to customize behavior
' ==================================================================================
Private Const MAIN_SIMULATION_COUNT As Long = 2000      ' Main simulation iterations
Private Const VOLATILITY_CHART_SIMS As Long = 100       ' Volatility chart simulations
Private Const SIMULATION_DAYS As Long = 365             ' Forecast horizon (days)
Private Const RISK_PERCENTILE As Double = 0.95          ' 95th percentile risk level
Private Const VOLATILITY_THRESHOLD As Double = 0.3      ' 30% CV for volatile products
Private Const OVERLOAD_BUFFER As Double = 1.1           ' 10% capacity buffer for overload detection
Private Const VERSION As String = "1.1.0"

' ==================================================================================

' ==================================================================================
' Main Entry Point - RunSharedResourceSimulation
' ==================================================================================
' Description:  Orchestrates the entire Monte Carlo simulation workflow
'
' Workflow:
'   1. Validates required worksheets exist (Simulation, SalesHistory)
'   2. Groups products by production line
'   3. Calculates statistical measures (avg, std dev) from sales history
'   4. Runs Monte Carlo simulation for each line to calculate safe lead times
'   5. Generates volatility spaghetti chart for visualization
'
' Prerequisites:
'   - "Simulation" sheet with columns: Product, Line Name, Line Capacity, etc.
'   - "SalesHistory" sheet with daily sales data
'
' Output:
'   - Populates columns G-M in Simulation sheet with calculated lead times
'   - Creates "Volatility Chart" sheet with visualization
'   - Displays completion message with execution time
'
' Error Handling:
'   - Validates worksheets before proceeding
'   - Provides user-friendly error messages
'   - Restores Excel settings (ScreenUpdating, Calculation) on error
' ==================================================================================
Public Sub RunSharedResourceSimulation()
    Dim startTime As Double
    startTime = Timer

    ' Optimization: Turn off screen updating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrorHandler

    ' Clear previous results
    ClearPreviousResults

    ' Step 1: Validate worksheets exist
    If Not ValidateWorksheets() Then
        MsgBox "Error: Required worksheets 'Simulation' and 'SalesHistory' not found!", vbCritical
        GoTo CleanUp
    End If

    ' Step 2: Group products by Line Name
    Dim lineGroups As Object
    Set lineGroups = GroupProductsByLine()

    If lineGroups.Count = 0 Then
        MsgBox "No products found in Simulation sheet!", vbExclamation
        GoTo CleanUp
    End If

    ' Step 3: Calculate statistical measures from sales history
    CalculateStatistics

    ' Step 4: Run Monte Carlo simulation for each line
    ProcessAllLines lineGroups

    ' Step 5: Create Volatility Spaghetti Chart
    CreateVolatilityChart

    ' Success message
    Dim elapsed As Double
    elapsed = Timer - startTime
    MsgBox "Simulation completed successfully in " & Format(elapsed, "0.00") & " seconds!" & vbCrLf & _
           "Analyzed " & lineGroups.Count & " production line(s).", vbInformation, "Monte Carlo Simulation"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error in RunSharedResourceSimulation: " & Err.Description, vbCritical
End Sub

' ==================================================================================
' Clear Previous Results
' ==================================================================================
Private Sub ClearPreviousResults()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Simulation")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow > 1 Then
        ' Clear columns C, D, G (Avg, StdDev, Recommended LT)
        ws.Range("C2:D" & lastRow).ClearContents
        ws.Range("G2:G" & lastRow).ClearContents
        ws.Range("G2:G" & lastRow).Interior.ColorIndex = xlNone
        ws.Range("G2:G" & lastRow).Font.Color = vbBlack
    End If
End Sub

' ==================================================================================
' Validate Required Worksheets Exist
' ==================================================================================
Private Function ValidateWorksheets() As Boolean
    On Error Resume Next
    Dim wsSimulation As Worksheet
    Dim wsSalesHistory As Worksheet

    Set wsSimulation = ThisWorkbook.Worksheets("Simulation")
    Set wsSalesHistory = ThisWorkbook.Worksheets("SalesHistory")

    ValidateWorksheets = (Not wsSimulation Is Nothing) And (Not wsSalesHistory Is Nothing)
    On Error GoTo 0
End Function

' ==================================================================================
' Group Products by Line Name Using Dictionary
' ==================================================================================
Private Function GroupProductsByLine() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Simulation")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim lineName As String
    Dim rowList As String

    ' Loop through each product row
    Dim productName As String
    For i = 2 To lastRow
        productName = Trim(ws.Cells(i, 1).Value) ' Column A: Product Name
        lineName = Trim(ws.Cells(i, 2).Value)    ' Column B: Line Name

        ' Skip rows with empty product name or line name
        If productName = "" Then
            ws.Cells(i, 7).Value = "ERROR: Missing Product Name"
            ws.Cells(i, 7).Interior.Color = RGB(255, 255, 0) ' Yellow
            ws.Cells(i, 7).Font.Bold = True
        ElseIf lineName = "" Then
            ws.Cells(i, 7).Value = "ERROR: Missing Line Name"
            ws.Cells(i, 7).Interior.Color = RGB(255, 255, 0) ' Yellow
            ws.Cells(i, 7).Font.Bold = True
        Else
            ' Valid row - add to dictionary
            If dict.Exists(lineName) Then
                ' Append row number to existing list
                dict(lineName) = dict(lineName) & "," & i
            Else
                ' Create new entry
                dict(lineName) = CStr(i)
            End If
        End If
    Next i

    Set GroupProductsByLine = dict
End Function

' ==================================================================================
' Calculate Average and Standard Deviation from Sales History
' ==================================================================================
Private Sub CalculateStatistics()
    Dim wsSimulation As Worksheet
    Dim wsSalesHistory As Worksheet
    Set wsSimulation = ThisWorkbook.Worksheets("Simulation")
    Set wsSalesHistory = ThisWorkbook.Worksheets("SalesHistory")

    Dim lastSimRow As Long
    lastSimRow = wsSimulation.Cells(wsSimulation.Rows.Count, 1).End(xlUp).Row

    Dim lastHistoryRow As Long
    lastHistoryRow = wsSalesHistory.Cells(wsSalesHistory.Rows.Count, 1).End(xlUp).Row

    Dim lastHistoryCol As Long
    lastHistoryCol = wsSalesHistory.Cells(1, wsSalesHistory.Columns.Count).End(xlToLeft).Column

    Dim i As Long
    Dim productName As String
    Dim historyRow As Long
    Dim salesRange As Range
    Dim avgDemand As Double
    Dim stdDev As Double

    ' Loop through each product in Simulation sheet
    For i = 2 To lastSimRow
        productName = Trim(wsSimulation.Cells(i, 1).Value) ' Column A: Product Name

        If productName <> "" Then
            ' Find product in SalesHistory
            historyRow = FindProductInHistory(productName, wsSalesHistory, lastHistoryRow)

            If historyRow > 0 Then
                ' Get sales data range (from column B onwards)
                Set salesRange = wsSalesHistory.Range(wsSalesHistory.Cells(historyRow, 2), _
                                                      wsSalesHistory.Cells(historyRow, lastHistoryCol))

                ' Calculate statistics
                avgDemand = WorksheetFunction.Average(salesRange)
                stdDev = WorksheetFunction.StDev(salesRange)

                ' Write to Simulation sheet
                wsSimulation.Cells(i, 3).Value = avgDemand ' Column C: Avg Demand
                wsSimulation.Cells(i, 4).Value = stdDev    ' Column D: Std Dev
            Else
                ' Product not found in history
                wsSimulation.Cells(i, 3).Value = "N/A"
                wsSimulation.Cells(i, 4).Value = "N/A"
            End If
        End If
    Next i
End Sub

' ==================================================================================
' Find Product Row in Sales History Sheet
' ==================================================================================
Private Function FindProductInHistory(productName As String, wsSalesHistory As Worksheet, lastRow As Long) As Long
    Dim i As Long
    For i = 2 To lastRow
        If Trim(wsSalesHistory.Cells(i, 1).Value) = productName Then
            FindProductInHistory = i
            Exit Function
        End If
    Next i
    FindProductInHistory = 0 ' Not found
End Function

' ==================================================================================
' Process All Lines (Main Simulation Loop)
' ==================================================================================
Private Sub ProcessAllLines(lineGroups As Object)
    Dim lineName As Variant
    Dim rowList As String
    Dim rows() As String

    ' Loop through each unique line
    For Each lineName In lineGroups.Keys
        rowList = lineGroups(lineName)
        rows = Split(rowList, ",")

        ' Run simulation for this line
        ProcessSingleLine CStr(lineName), rows
    Next lineName
End Sub

' ==================================================================================
' Process Single Line - The Core "Shared Fate" Simulation
' ==================================================================================
' Description:  Runs Monte Carlo simulation for a single production line to calculate
'               safe lead times accounting for shared capacity constraints
'
' Parameters:
'   lineName  - Name of the production line (e.g., "Line 1")
'   rows()    - Array of row numbers in Simulation sheet for products on this line
'
' Algorithm:
'   1. Read inputs: capacity, backlog, product statistics
'   2. Run 2,000 Monte Carlo simulations over 365 days
'   3. For each simulation, track maximum queue backlog
'   4. Calculate 95th percentile of maximum queues
'   5. Allocate buffer using variance-weighted method
'   6. Calculate system buffer and large order thresholds
'   7. Write results to columns G-M with color coding
'
' Output Columns:
'   G - Equal Treatment LT (uniform buffer allocation)
'   H - Risk-Based LT (variance-weighted buffer)
'   I - Material Lead Time (from input)
'   J - System Buffer (spare capacity indicator)
'   K - Max Safe Order Qty (before overload)
'   L - Large Order Quantity (2-sigma threshold)
'   M - Large Order LT Quote (extended lead time)
'
' Color Coding:
'   Green  - Healthy capacity (System Buffer ≥ 2)
'   Orange - Fragile capacity (0 < System Buffer < 2)
'   Red    - Critical overload (System Buffer ≤ 0)
'   Pink   - Overload warning message
' ==================================================================================
Private Sub ProcessSingleLine(lineName As String, rows() As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Simulation")

    ' Step A: Read inputs from the first row of this line group
    Dim firstRow As Long
    firstRow = CLng(rows(0))

    Dim capacity As Double
    Dim startQueue As Double

    capacity = ws.Cells(firstRow, 5).Value    ' Column E: Line Capacity
    startQueue = ws.Cells(firstRow, 6).Value  ' Column F: Line Start Backlog

    ' Validate capacity
    If capacity <= 0 Then
        FlagLineAsError rows, "ERROR: Invalid Capacity" & vbCrLf & _
                             "Capacity must be greater than 0." & vbCrLf & _
                             "Current value: " & capacity
        Exit Sub
    End If

    ' Validate start queue (should not be negative)
    If startQueue < 0 Then
        FlagLineAsError rows, "ERROR: Invalid Start Backlog" & vbCrLf & _
                             "Start backlog cannot be negative." & vbCrLf & _
                             "Current value: " & startQueue
        Exit Sub
    End If

    ' Step B: Collect product statistics for all products on this line
    Dim productCount As Long
    productCount = UBound(rows) - LBound(rows) + 1

    ReDim productAvgs(0 To productCount - 1) As Double
    ReDim productStdDevs(0 To productCount - 1) As Double
    ReDim productMaterialLTs(0 To productCount - 1) As Double

    Dim i As Long
    Dim rowNum As Long
    Dim totalAvgDemand As Double
    totalAvgDemand = 0

    For i = 0 To productCount - 1
        rowNum = CLng(rows(i))

        ' Read and validate avg and std dev
        If IsNumeric(ws.Cells(rowNum, 3).Value) And IsNumeric(ws.Cells(rowNum, 4).Value) Then
            productAvgs(i) = ws.Cells(rowNum, 3).Value
            productStdDevs(i) = ws.Cells(rowNum, 4).Value

            ' Validate non-negative demand
            If productAvgs(i) < 0 Then
                FlagLineAsError rows, "ERROR: Invalid Average Demand" & vbCrLf & _
                                     "Average demand cannot be negative." & vbCrLf & _
                                     "Product in row " & rowNum & ": " & productAvgs(i)
                Exit Sub
            End If

            ' Validate non-negative standard deviation
            If productStdDevs(i) < 0 Then
                FlagLineAsError rows, "ERROR: Invalid Standard Deviation" & vbCrLf & _
                                     "Standard deviation cannot be negative." & vbCrLf & _
                                     "Product in row " & rowNum & ": " & productStdDevs(i)
                Exit Sub
            End If

            totalAvgDemand = totalAvgDemand + productAvgs(i)

            ' Read Material Lead Time from Column I (default to 0 if not provided)
            If IsNumeric(ws.Cells(rowNum, 9).Value) Then
                productMaterialLTs(i) = ws.Cells(rowNum, 9).Value

                ' Validate non-negative material lead time
                If productMaterialLTs(i) < 0 Then
                    FlagLineAsError rows, "ERROR: Invalid Material Lead Time" & vbCrLf & _
                                         "Material LT cannot be negative." & vbCrLf & _
                                         "Product in row " & rowNum & ": " & productMaterialLTs(i)
                    Exit Sub
                End If
            Else
                productMaterialLTs(i) = 0
            End If
        Else
            FlagLineAsError rows, "ERROR: Missing Statistics" & vbCrLf & _
                                 "Average Demand and Std Dev required." & vbCrLf & _
                                 "Check row " & rowNum & " in Simulation sheet."
            Exit Sub
        End If
    Next i

    ' Critical Check: Overload detection (but still run simulation to show impact)
    Dim isOverloaded As Boolean
    isOverloaded = (totalAvgDemand > capacity)

    ' Step C: Run Monte Carlo Simulation
    Const NUM_SIMULATIONS As Long = MAIN_SIMULATION_COUNT
    Const NUM_DAYS As Long = SIMULATION_DAYS

    ReDim maxQueues(1 To NUM_SIMULATIONS) As Double

    Dim sim As Long
    Dim day As Long
    Dim currentQueue As Double
    Dim totalLineDemand As Double
    Dim prodDemand As Double
    Dim maxQueueThisSim As Double

    ' Seed randomizer
    Randomize

    For sim = 1 To NUM_SIMULATIONS
        currentQueue = startQueue
        maxQueueThisSim = currentQueue

        For day = 1 To NUM_DAYS
            totalLineDemand = 0

            ' Sum demand from all products on this line
            For i = 0 To productCount - 1
                ' Generate random demand using normal distribution
                prodDemand = NormInv(Rnd(), productAvgs(i), productStdDevs(i))
                If prodDemand < 0 Then prodDemand = 0 ' Demand cannot be negative
                totalLineDemand = totalLineDemand + prodDemand
            Next i

            ' Queue logic: Add demand, subtract capacity
            currentQueue = currentQueue + totalLineDemand - capacity
            If currentQueue < 0 Then currentQueue = 0 ' Queue cannot be negative

            ' Track maximum queue for this simulation
            If currentQueue > maxQueueThisSim Then
                maxQueueThisSim = currentQueue
            End If
        Next day

        maxQueues(sim) = maxQueueThisSim
    Next sim

    ' Step D: Calculate Risk Percentile (default: 95th percentile)
    Dim percentile95 As Double
    percentile95 = Percentile(maxQueues, RISK_PERCENTILE)

    ' Step E: Calculate Shared Line Buffer (95th percentile queue time)
    Dim sharedLineBuffer As Double
    sharedLineBuffer = percentile95 / capacity

    ' Step E2: Calculate Variance-Weighted Buffer Allocation
    ' Each product gets buffer proportional to their volatility contribution
    Dim totalStdDev As Double
    totalStdDev = 0
    For i = 0 To productCount - 1
        totalStdDev = totalStdDev + productStdDevs(i)
    Next i

    ReDim productBuffers(0 To productCount - 1) As Double
    For i = 0 To productCount - 1
        If totalStdDev > 0 Then
            ' Allocate buffer proportional to each product's StdDev contribution
            productBuffers(i) = (productStdDevs(i) / totalStdDev) * sharedLineBuffer
        Else
            ' If no volatility, split equally
            productBuffers(i) = sharedLineBuffer / productCount
        End If
    Next i

    ' Step F: Write back to ALL rows on this line (Material LT + Variance-Weighted Buffer)
    Dim totalLeadTime As Double

    If isOverloaded Then
        ' Calculate required capacity to handle the demand
        Dim requiredCapacity As Double
        requiredCapacity = totalAvgDemand * OVERLOAD_BUFFER

        ' Show overload warning for BOTH methodologies
        Dim overloadMsgEqual As String
        Dim overloadMsgRisk As String
        Dim totalLT_Equal As Double
        Dim totalLT_Risk As Double

        For i = 0 To productCount - 1
            rowNum = CLng(rows(i))

            ' Column G: Equal Treatment (Uniform Buffer)
            totalLT_Equal = productMaterialLTs(i) + sharedLineBuffer
            overloadMsgEqual = "OVERLOADED CAPACITY:" & vbCrLf & _
                              "Recommended LT " & Round(totalLT_Equal, 1) & " days" & vbCrLf & _
                              "(Equal: " & Round(sharedLineBuffer, 1) & ")" & vbCrLf & _
                              "Need " & Round(requiredCapacity, 1) & " vs " & Round(capacity, 1)
            ws.Cells(rowNum, 7).Value = overloadMsgEqual
            ws.Cells(rowNum, 7).Interior.Color = RGB(255, 182, 193) ' Light pink
            ws.Cells(rowNum, 7).Font.Color = RGB(0, 0, 0) ' Black text
            ws.Cells(rowNum, 7).Font.Bold = True
            ws.Cells(rowNum, 7).WrapText = True

            ' Column H: Risk-Based (Variance-Weighted Buffer)
            totalLT_Risk = productMaterialLTs(i) + productBuffers(i)
            overloadMsgRisk = "OVERLOADED CAPACITY:" & vbCrLf & _
                             "Recommended LT " & Round(totalLT_Risk, 1) & " days" & vbCrLf & _
                             "(Risk: " & Round(productBuffers(i), 1) & ")" & vbCrLf & _
                             "Need " & Round(requiredCapacity, 1) & " vs " & Round(capacity, 1)
            ws.Cells(rowNum, 8).Value = overloadMsgRisk
            ws.Cells(rowNum, 8).Interior.Color = RGB(255, 182, 193) ' Light pink
            ws.Cells(rowNum, 8).Font.Color = RGB(0, 0, 0) ' Black text
            ws.Cells(rowNum, 8).Font.Bold = True
            ws.Cells(rowNum, 8).WrapText = True
        Next i
    Else
        ' Normal result - show BOTH methodologies side-by-side
        Dim msgEqual As String
        Dim msgRisk As String

        For i = 0 To productCount - 1
            rowNum = CLng(rows(i))

            ' Column G: Equal Treatment (Uniform Buffer)
            totalLeadTime = productMaterialLTs(i) + sharedLineBuffer
            msgEqual = "Recommended LT " & Round(totalLeadTime, 1) & " days" & vbCrLf & _
                      "(Shared: " & Round(sharedLineBuffer, 1) & ")"
            ws.Cells(rowNum, 7).Value = msgEqual
            ws.Cells(rowNum, 7).Interior.Color = RGB(144, 238, 144) ' Light green
            ws.Cells(rowNum, 7).Font.Color = RGB(0, 0, 0) ' Black text
            ws.Cells(rowNum, 7).WrapText = True

            ' Column H: Risk-Based (Variance-Weighted Buffer)
            totalLeadTime = productMaterialLTs(i) + productBuffers(i)
            msgRisk = "Recommended LT " & Round(totalLeadTime, 1) & " days" & vbCrLf & _
                     "(Risk: " & Round(productBuffers(i), 1) & ")"
            ws.Cells(rowNum, 8).Value = msgRisk
            ws.Cells(rowNum, 8).Interior.Color = RGB(144, 238, 144) ' Light green
            ws.Cells(rowNum, 8).Font.Color = RGB(0, 0, 0) ' Black text
            ws.Cells(rowNum, 8).WrapText = True
        Next i
    End If

    ' Step G: Calculate and write System Buffer and Max Safe Order Qty for ALL products on this line
    Dim systemBuffer As Double
    Dim maxSafeOrderQty As Double
    Dim bufferColor As Long

    ' Calculate System Buffer (Toxic Threshold) for this line
    systemBuffer = capacity - totalAvgDemand

    ' Determine color based on System Buffer health
    If systemBuffer <= 0 Then
        bufferColor = RGB(255, 99, 71) ' Red (Critical - Underwater)
    ElseIf systemBuffer < 2 Then
        bufferColor = RGB(255, 165, 0) ' Orange (Fragile)
    Else
        bufferColor = RGB(144, 238, 144) ' Green (Healthy)
    End If

    ' Write System Buffer and Max Safe Order Qty for each product
    For i = 0 To productCount - 1
        rowNum = CLng(rows(i))

        ' Column J: Toxic Threshold (System Buffer)
        ws.Cells(rowNum, 10).Value = Round(systemBuffer, 2)
        ws.Cells(rowNum, 10).Interior.Color = bufferColor
        ws.Cells(rowNum, 10).Font.Color = RGB(0, 0, 0) ' Black text
        ws.Cells(rowNum, 10).Font.Bold = True

        ' Column K: Max Safe Order Qty (Product Avg + System Buffer)
        maxSafeOrderQty = productAvgs(i) + systemBuffer
        ws.Cells(rowNum, 11).Value = Round(maxSafeOrderQty, 2)
        ws.Cells(rowNum, 11).Interior.Color = bufferColor
        ws.Cells(rowNum, 11).Font.Color = RGB(0, 0, 0) ' Black text
        ws.Cells(rowNum, 11).Font.Bold = True

        ' Column L: Large Order Quantity (Statistical Method - 2 Sigma)
        ' Formula: Avg Demand + (2 × Std Dev) = 95th percentile
        Dim largeOrderThreshold As Double
        largeOrderThreshold = productAvgs(i) + (2 * productStdDevs(i))
        ws.Cells(rowNum, 12).Value = Round(largeOrderThreshold, 2)
        ws.Cells(rowNum, 12).Font.Color = RGB(0, 0, 0) ' Black text

        ' Column M: Large Order LT Quote
        ' Formula: Material LT + Queue Buffer + (Excess Qty / Capacity)
        ' Excess Qty = Large Order Quantity - Avg Demand
        Dim excessQty As Double
        Dim extraDaysNeeded As Double
        Dim largeOrderLT As Double

        excessQty = largeOrderThreshold - productAvgs(i)
        If capacity > 0 Then
            extraDaysNeeded = excessQty / capacity
        Else
            extraDaysNeeded = 0
        End If

        ' Use the variance-weighted buffer for this product
        largeOrderLT = productMaterialLTs(i) + productBuffers(i) + extraDaysNeeded

        ws.Cells(rowNum, 13).Value = Round(largeOrderLT, 1)
        ws.Cells(rowNum, 13).Font.Color = RGB(0, 0, 0) ' Black text
    Next i

    ' Auto-size columns and rows to prevent text smooshing
    ws.Columns("G:M").AutoFit
    ws.Rows.AutoFit
End Sub

' ==================================================================================
' Normal Inverse Function (Approximation of Excel's NORMINV)
' ==================================================================================
Private Function NormInv(probability As Double, mean As Double, stdDev As Double) As Double
    ' Using Box-Muller transform for normal distribution
    Dim u1 As Double, u2 As Double
    Dim z As Double

    ' Ensure probability is in valid range
    If probability <= 0 Then probability = 0.00001
    If probability >= 1 Then probability = 0.99999

    ' Simple approximation using inverse error function
    ' For better accuracy, we use Excel's WorksheetFunction
    On Error Resume Next
    z = WorksheetFunction.NormSInv(probability)
    On Error GoTo 0

    NormInv = mean + z * stdDev
End Function

' ==================================================================================
' Calculate Percentile from Array
' ==================================================================================
Private Function Percentile(arr() As Double, percentileValue As Double) As Double
    ' Sort array
    Dim sortedArr() As Double
    sortedArr = BubbleSort(arr)

    Dim n As Long
    n = UBound(sortedArr) - LBound(sortedArr) + 1

    Dim index As Long
    index = Int(percentileValue * (n - 1)) + LBound(sortedArr)

    If index > UBound(sortedArr) Then index = UBound(sortedArr)

    Percentile = sortedArr(index)
End Function

' ==================================================================================
' Simple Bubble Sort for Double Array
' ==================================================================================
Private Function BubbleSort(arr() As Double) As Double()
    Dim tempArr() As Double
    ReDim tempArr(LBound(arr) To UBound(arr))

    Dim i As Long, j As Long
    Dim temp As Double

    ' Copy array
    For i = LBound(arr) To UBound(arr)
        tempArr(i) = arr(i)
    Next i

    ' Bubble sort
    For i = LBound(tempArr) To UBound(tempArr) - 1
        For j = i + 1 To UBound(tempArr)
            If tempArr(i) > tempArr(j) Then
                temp = tempArr(i)
                tempArr(i) = tempArr(j)
                tempArr(j) = temp
            End If
        Next j
    Next i

    BubbleSort = tempArr
End Function

' ==================================================================================
' Flag Line as Overloaded (Red Warning)
' ==================================================================================
Private Sub FlagLineAsOverload(rows() As String, message As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Simulation")

    Dim i As Long
    Dim rowNum As Long

    For i = 0 To UBound(rows)
        rowNum = CLng(rows(i))
        ws.Cells(rowNum, 7).Value = message
        ws.Cells(rowNum, 7).Interior.Color = RGB(255, 0, 0) ' Red
        ws.Cells(rowNum, 7).Font.Color = RGB(255, 255, 255) ' White text
        ws.Cells(rowNum, 7).Font.Bold = True
    Next i
End Sub

' ==================================================================================
' Flag Line as Error (Yellow Warning)
' ==================================================================================
Private Sub FlagLineAsError(rows() As String, message As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Simulation")

    Dim i As Long
    Dim rowNum As Long

    For i = 0 To UBound(rows)
        rowNum = CLng(rows(i))
        ws.Cells(rowNum, 7).Value = message
        ws.Cells(rowNum, 7).Interior.Color = RGB(255, 255, 0) ' Yellow
        ws.Cells(rowNum, 7).Font.Color = RGB(0, 0, 0) ' Black text
        ws.Cells(rowNum, 7).Font.Bold = True
    Next i
End Sub

' ==================================================================================
' Create Volatility Spaghetti Chart
' ==================================================================================
' Description:  Generates a visualization showing 100 Monte Carlo simulation runs
'               for volatile products to demonstrate why longer lead times are needed
'
' Algorithm:
'   1. Identify volatile products (Coefficient of Variation > 30%)
'   2. Run 100 Monte Carlo simulations over 365 days
'   3. Track daily queue backlog for each simulation
'   4. Plot all runs as a spaghetti diagram (3 highlighted, 97 gray)
'   5. Add color legend and interpretation guide
'
' Output:
'   - Creates/updates "Volatility Chart" worksheet
'   - Line chart with 100 series (simulation runs)
'   - Explanatory text with color-coded legend
'
' Chart Elements:
'   - Red Line: Example simulation #1
'   - Orange Line: Example simulation #2
'   - Crimson Line: Example simulation #3
'   - Gray Lines: Remaining 97 simulations (semi-transparent)
'
' Purpose:
'   - Visual proof of demand uncertainty
'   - Justifies longer lead times for volatile products
'   - Shows wide spread of possible queue backlog outcomes
' ==================================================================================
Private Sub CreateVolatilityChart()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim chartSheet As Worksheet
    Dim simWs As Worksheet
    Set simWs = ThisWorkbook.Worksheets("Simulation")
    
    ' Find VOLATILE products across all lines (CV > 30%)
    Dim volatileProducts As Collection
    Set volatileProducts = New Collection

    Dim lastRow As Long
    lastRow = simWs.Cells(simWs.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim productRow As Long
    Dim capacity As Double
    Dim avgDemand As Double
    Dim stdDev As Double
    Dim coefficientOfVariation As Double
    Dim productName As String
    Dim lineName As String

    ' Collect ONLY volatile products (high Coefficient of Variation)
    For i = 2 To lastRow
        avgDemand = simWs.Cells(i, 3).Value ' Avg Demand
        stdDev = simWs.Cells(i, 4).Value ' Std Dev
        productName = simWs.Cells(i, 1).Value
        lineName = simWs.Cells(i, 2).Value

        If avgDemand > 0 Then
            coefficientOfVariation = stdDev / avgDemand

            ' Only include volatile products (CV > 30%)
            If coefficientOfVariation > VOLATILITY_THRESHOLD Then
                Dim prodData(1 To 5) As Variant
                prodData(1) = avgDemand ' Avg Demand
                prodData(2) = stdDev ' Std Dev
                prodData(3) = i ' Row number
                prodData(4) = productName ' Product Name
                prodData(5) = lineName ' Line Name
                volatileProducts.Add prodData

                ' Get capacity from first volatile product found
                If volatileProducts.Count = 1 Then
                    capacity = simWs.Cells(i, 5).Value
                End If
            End If
        End If
    Next i

    If volatileProducts.Count = 0 Then Exit Sub

    ' Calculate total demand for volatile products only
    Dim totalAvg As Double
    totalAvg = 0
    For i = 1 To volatileProducts.Count
        totalAvg = totalAvg + volatileProducts(i)(1)
    Next i
    
    ' Create or clear Volatility Chart sheet
    On Error Resume Next
    Set chartSheet = ThisWorkbook.Worksheets("Volatility Chart")
    On Error GoTo ErrorHandler
    
    If chartSheet Is Nothing Then
        Set chartSheet = ThisWorkbook.Worksheets.Add(After:=simWs)
        chartSheet.Name = "Volatility Chart"
    Else
        chartSheet.Cells.Clear
        chartSheet.ChartObjects.Delete
    End If
    
    ' Run sample Monte Carlo simulations for visualization
    Const NUM_SIMULATIONS As Long = VOLATILITY_CHART_SIMS
    Const NUM_DAYS As Long = SIMULATION_DAYS
    
    Dim simData() As Double
    ReDim simData(1 To NUM_SIMULATIONS, 1 To NUM_DAYS)
    
    Dim sim As Long
    Dim day As Long
    Dim queue As Double
    Dim demand As Double
    Dim j As Long
    
    ' Run simulations
    For sim = 1 To NUM_SIMULATIONS
        queue = 5 ' Start backlog
        
        For day = 1 To NUM_DAYS
            ' Generate total demand for VOLATILE products only
            demand = 0
            For j = 1 To volatileProducts.Count
                avgDemand = volatileProducts(j)(1)
                stdDev = volatileProducts(j)(2)
                demand = demand + NormInv(Rnd(), avgDemand, stdDev)
            Next j
            
            ' Update queue
            queue = queue + demand - capacity
            If queue < 0 Then queue = 0
            
            ' Store queue level
            simData(sim, day) = queue
        Next day
    Next sim
    
    ' Write data to chart sheet
    chartSheet.Cells(1, 1).Value = "Day"
    For sim = 1 To NUM_SIMULATIONS
        chartSheet.Cells(1, sim + 1).Value = "Sim " & sim
    Next sim
    
    ' Write day numbers and queue data
    For day = 1 To NUM_DAYS
        chartSheet.Cells(day + 1, 1).Value = day
        For sim = 1 To NUM_SIMULATIONS
            chartSheet.Cells(day + 1, sim + 1).Value = simData(sim, day)
        Next sim
    Next day
    
    ' Create the chart
    Dim chartObj As ChartObject
    Set chartObj = chartSheet.ChartObjects.Add(Left:=50, Top:=50, Width:=1200, Height:=600)
    
    With chartObj.Chart
        .ChartType = xlLine
        .SetSourceData chartSheet.Range(chartSheet.Cells(1, 1), chartSheet.Cells(NUM_DAYS + 1, NUM_SIMULATIONS + 1))
        
        ' Format chart
        .HasTitle = True

        ' Build list of volatile products for title
        Dim volatileList As String
        volatileList = ""
        For j = 1 To volatileProducts.Count
            If j > 1 Then volatileList = volatileList & ", "
            volatileList = volatileList & volatileProducts(j)(4) ' Product Name
        Next j

        .ChartTitle.Text = "Monte Carlo Volatility Analysis: Queue Backlog Over Time" & vbCrLf & _
                           "Volatile Products Only (" & volatileList & ") - 100 Simulation Runs"
        .ChartTitle.Font.Size = 14
        .ChartTitle.Font.Bold = True
        
        ' Axes
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Day"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Queue Backlog (units)"
        
        ' Make most lines semi-transparent
        For i = 1 To NUM_SIMULATIONS
            With .SeriesCollection(i)
                .Format.Line.ForeColor.RGB = RGB(180, 180, 180)
                .Format.Line.Weight = 1
                .Format.Line.Transparency = 0.7
            End With
        Next i
        
        ' Highlight the first 3 simulations in color
        If NUM_SIMULATIONS >= 1 Then
            With .SeriesCollection(1)
                .Format.Line.ForeColor.RGB = RGB(255, 68, 68) ' Red
                .Format.Line.Weight = 2.5
                .Format.Line.Transparency = 0
            End With
        End If
        
        If NUM_SIMULATIONS >= 2 Then
            With .SeriesCollection(2)
                .Format.Line.ForeColor.RGB = RGB(255, 140, 0) ' Orange
                .Format.Line.Weight = 2.5
                .Format.Line.Transparency = 0
            End With
        End If
        
        If NUM_SIMULATIONS >= 3 Then
            With .SeriesCollection(3)
                .Format.Line.ForeColor.RGB = RGB(220, 20, 60) ' Crimson
                .Format.Line.Weight = 2.5
                .Format.Line.Transparency = 0
            End With
        End If
        
        ' Remove legend (too many series)
        .HasLegend = False
        
        ' Add gridlines
        .Axes(xlValue).HasMajorGridlines = True
    End With
    
    ' Add explanatory text with color legend
    chartSheet.Cells(40, 1).Value = "CHART LEGEND & INTERPRETATION:"
    chartSheet.Cells(40, 1).Font.Bold = True
    chartSheet.Cells(40, 1).Font.Size = 12
    chartSheet.Cells(40, 1).Font.Color = RGB(0, 0, 0)

    ' Color legend
    chartSheet.Cells(42, 1).Value = "COLOR LEGEND:"
    chartSheet.Cells(42, 1).Font.Bold = True

    chartSheet.Cells(43, 1).Value = "🔴 RED LINE = Example Simulation #1 (highlighted for visibility)"
    chartSheet.Cells(43, 1).Font.Color = RGB(255, 0, 0)

    chartSheet.Cells(44, 1).Value = "🟠 ORANGE LINE = Example Simulation #2 (highlighted for visibility)"
    chartSheet.Cells(44, 1).Font.Color = RGB(255, 140, 0)

    chartSheet.Cells(45, 1).Value = "🩷 CRIMSON LINE = Example Simulation #3 (highlighted for visibility)"
    chartSheet.Cells(45, 1).Font.Color = RGB(220, 20, 60)

    chartSheet.Cells(46, 1).Value = "⬜ GRAY LINES = Remaining 97 simulations (semi-transparent to show range)"
    chartSheet.Cells(46, 1).Font.Color = RGB(128, 128, 128)

    ' Key insights
    chartSheet.Cells(48, 1).Value = "WHY THIS MATTERS:"
    chartSheet.Cells(48, 1).Font.Bold = True
    chartSheet.Cells(48, 1).Font.Color = RGB(0, 0, 0)

    ' Build volatile product list for explanation
    Dim volatileNames As String
    volatileNames = ""
    For j = 1 To volatileProducts.Count
        If j > 1 Then volatileNames = volatileNames & ", "
        volatileNames = volatileNames & volatileProducts(j)(4)
    Next j

    chartSheet.Cells(49, 1).Value = "• This chart shows 100 different possible futures for queue backlog over 365 days"
    chartSheet.Cells(50, 1).Value = "• Based on volatile product demand: " & volatileNames & " (Coefficient of Variation > 30%)"
    chartSheet.Cells(51, 1).Value = "• Line Capacity: " & capacity & " units/day | Average Demand: " & Format(totalAvg, "0.0") & " units/day"
    chartSheet.Cells(52, 1).Value = "• The WIDE SPREAD of outcomes shows high uncertainty due to demand volatility"
    chartSheet.Cells(53, 1).Value = "• Queue can grow from 0 to 300+ units depending on when demand spikes occur"
    chartSheet.Cells(54, 1).Value = "• This unpredictability is WHY we need safety buffers and longer lead times"

    ' Format text
    chartSheet.Range("A42:A54").WrapText = False
    chartSheet.Range("A49:A54").Font.Color = RGB(0, 0, 0)
    
    ' Auto-fit column
    chartSheet.Columns("A:A").ColumnWidth = 100
    
    ' Activate the chart sheet
    chartSheet.Select
    chartObj.Select
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in CreateVolatilityChart: " & Err.Description
End Sub
