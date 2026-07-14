Attribute VB_Name = "mod_BPO_Dashboards"
Option Explicit

'=====================================================================================
'  PH WORKFORCE ANALYTICS - EXECUTIVE & PAYROLL DASHBOARD BUILDER
'  ------------------------------------------------------------------------------------
'  Rebuilds "Dashboard_Executive" and builds "Dashboard_Payroll" from the workbook's
'  existing data tables (Employees, Attendance, Payroll_Engine, Attrition_Analysis,
'  Headcount_Forecasts). Two hidden calc sheets (Calc_Exec, Calc_Payroll) hold the
'  SUMIFS/COUNTIFS rollups that don't already exist in the workbook: monthly payroll
'  cost by department/site, and voluntary vs. involuntary exits by month.
'
'  HOW TO INSTALL
'    1. Open the workbook, press Alt+F11 to open the VBA editor.
'    2. File > Import File... and select this .bas file (or paste into a new Module).
'    3. Press F5, or run Sub "Build_BPO_Dashboards", to build/refresh both dashboards.
'    4. A "Refresh Dashboards" button is placed on each sheet to re-run this later.
'
'  SAFE TO RE-RUN: the macro clears and rebuilds its target sheets each time, so you
'  can re-run it anytime after the source data changes.
'
'  DESIGN NOTE: this build reads from the workbook's existing computed tables
'  (Attrition_Analysis, Headcount_Forecasts) rather than re-deriving attrition/
'  forecast logic, and avoids PivotTables/Slicers (the file already showed signs of
'  slicer/extension corruption on open) in favor of plain formulas + native charts,
'  which are far more robust to re-runs and Excel-version differences.
'=====================================================================================

' ---- THEME (assigned at runtime; RGB() cannot be used in Const declarations) ----
Private C_NAVY As Long
Private C_NAVY_DARK As Long
Private C_TEAL As Long
Private C_AMBER As Long
Private C_RED As Long
Private C_GREEN As Long
Private C_GREY As Long
Private C_GREY_LIGHT As Long
Private C_WHITE As Long

Private Const CALC_EXEC_SHEET As String = "Calc_Exec"
Private Const CALC_PAYROLL_SHEET As String = "Calc_Payroll"
Private Const N_PERIODS As Integer = 48   ' 2022-01 .. 2025-12

'=====================================================================================
' MAIN ENTRY POINT
'=====================================================================================
Sub Build_BPO_Dashboards()

    Dim startTime As Double
    startTime = Timer

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    InitTheme
    PopulateSiteHelperColumn

    EnsureCalcSheet CALC_EXEC_SHEET
    EnsureCalcSheet CALC_PAYROLL_SHEET
    PopulateCalcExec
    PopulateCalcPayroll

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild

    BuildExecutiveDashboard
    BuildPayrollDashboard

    ThisWorkbook.Sheets(CALC_EXEC_SHEET).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(CALC_PAYROLL_SHEET).Visible = xlSheetVeryHidden

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    ThisWorkbook.Sheets("Dashboard_Executive").Activate

    MsgBox "Dashboards rebuilt successfully in " & Format(Timer - startTime, "0.0") & " seconds.", _
           vbInformation, "BPO Workforce Analytics"
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    MsgBox "Dashboard build failed:" & vbNewLine & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "BPO Workforce Analytics"
End Sub

Private Sub InitTheme()
    C_NAVY = RGB(20, 33, 61)
    C_NAVY_DARK = RGB(12, 20, 38)
    C_TEAL = RGB(0, 150, 156)
    C_AMBER = RGB(240, 160, 0)
    C_RED = RGB(196, 60, 60)
    C_GREEN = RGB(46, 139, 87)
    C_GREY = RGB(110, 118, 128)
    C_GREY_LIGHT = RGB(238, 240, 243)
    C_WHITE = RGB(255, 255, 255)
End Sub

'=====================================================================================
' DATA PREP: populate the blank "Site" column in Payroll_Engine via lookup to Employees
'=====================================================================================
Private Sub PopulateSiteHelperColumn()
    Dim loPE As ListObject
    Set loPE = ThisWorkbook.Sheets("Payroll_Engine").ListObjects("tbl_PayrollEngine7")
    On Error Resume Next
    loPE.ListColumns("Site").DataBodyRange.Formula = _
        "=IFERROR(INDEX(Employees[Site],MATCH([@EmployeeID],Employees[EmployeeID],0)),""Unknown"")"
    On Error GoTo 0
End Sub

'=====================================================================================
' Ensure a hidden calc sheet exists and is empty
'=====================================================================================
Private Sub EnsureCalcSheet(shName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = shName
    Else
        ws.Visible = xlSheetVisible
        ws.Cells.Clear
    End If
End Sub

'=====================================================================================
' CALC_EXEC: monthly voluntary / involuntary exits (2022-01 .. 2025-12)
'=====================================================================================
Private Sub PopulateCalcExec()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CALC_EXEC_SHEET)

    ws.Range("A1:D1").Value = Array("MonthStart", "MonthEnd", "VoluntaryExits", "InvoluntaryExits")

    Dim i As Integer
    For i = 0 To N_PERIODS - 1
        ws.Cells(2 + i, 1).Value = DateAdd("m", i, DateSerial(2022, 1, 1))
    Next i
    Dim lastRow As Long
    lastRow = 1 + N_PERIODS
    ws.Range("A2:A" & lastRow).NumberFormat = "yyyy-mm"

    ws.Range("B2:B" & lastRow).FormulaR1C1 = "=EDATE(RC[-1],1)"

    ws.Range("C2:C" & lastRow).FormulaR1C1 = _
        "=COUNTIFS(Employees[DateExited],"">=""&RC[-2],Employees[DateExited],""<""&RC[-1],Employees[ExitReason],""Resignation"")" & _
        "+COUNTIFS(Employees[DateExited],"">=""&RC[-2],Employees[DateExited],""<""&RC[-1],Employees[ExitReason],""AWOL"")"

    ws.Range("D2:D" & lastRow).FormulaR1C1 = _
        "=COUNTIFS(Employees[DateExited],"">=""&RC[-3],Employees[DateExited],""<""&RC[-2],Employees[ExitReason],""Termination"")" & _
        "+COUNTIFS(Employees[DateExited],"">=""&RC[-3],Employees[DateExited],""<""&RC[-2],Employees[ExitReason],""End of Contract"")"

    ws.Calculate
End Sub

'=====================================================================================
' CALC_PAYROLL: monthly cost trend, department rollup, site rollup, statutory mix
'=====================================================================================
Private Sub PopulateCalcPayroll()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CALC_PAYROLL_SHEET)

    ' --- A:M monthly trend ---
    ws.Range("A1:M1").Value = Array("PayPeriod", "BasicPay", "OTPay", "NightDiffPay", "HolidayPay", _
        "GrossPay", "SSS", "PhilHealth", "PagIBIG", "WithholdingTax", "NetPay", "Headcount", "CostPerFTE")

    Dim i As Integer
    For i = 0 To N_PERIODS - 1
        ws.Cells(2 + i, 1).Value = DateAdd("m", i, DateSerial(2022, 1, 1))
    Next i
    Dim lastRow As Long
    lastRow = 1 + N_PERIODS
    ws.Range("A2:A" & lastRow).NumberFormat = "mmm-yyyy"

    ws.Range("B2:B" & lastRow).FormulaR1C1 = "=SUMIFS(tbl_PayrollEngine7[BasicPay],tbl_PayrollEngine7[PayPeriod],RC[-1])"
    ws.Range("C2:C" & lastRow).FormulaR1C1 = "=SUMIFS(tbl_PayrollEngine7[OTPay],tbl_PayrollEngine7[PayPeriod],RC[-2])"
    ws.Range("D2:D" & lastRow).FormulaR1C1 = "=SUMIFS(tbl_PayrollEngine7[NightDiffPay],tbl_PayrollEngine7[PayPeriod],RC[-3])"
    ws.Range("E2:E" & lastRow).FormulaR1C1 = "=SUMIFS(tbl_PayrollEngine7[HolidayPay],tbl_PayrollEngine7[PayPeriod],RC[-4])"
    ws.Range("F2:F" & lastRow).FormulaR1C1 = "=SUMIFS(tbl_PayrollEngine7[GrossPay],tbl_PayrollEngine7[PayPeriod],RC[-5])"
    ws.Range("G2:G" & lastRow).FormulaR1C1 = "=SUMIFS(tbl_PayrollEngine7[SSS_EmployeeShare],tbl_PayrollEngine7[PayPeriod],RC[-6])"
    ws.Range("H2:H" & lastRow).FormulaR1C1 = "=SUMIFS(tbl_PayrollEngine7[PhilHealth_EmployeeShare],tbl_PayrollEngine7[PayPeriod],RC[-7])"
    ws.Range("I2:I" & lastRow).FormulaR1C1 = "=SUMIFS(tbl_PayrollEngine7[PagIBIG_EmployeeShare],tbl_PayrollEngine7[PayPeriod],RC[-8])"
    ws.Range("J2:J" & lastRow).FormulaR1C1 = "=SUMIFS(tbl_PayrollEngine7[WithholdingTax],tbl_PayrollEngine7[PayPeriod],RC[-9])"
    ws.Range("K2:K" & lastRow).FormulaR1C1 = "=SUMIFS(tbl_PayrollEngine7[NetPay],tbl_PayrollEngine7[PayPeriod],RC[-10])"
    ws.Range("L2:L" & lastRow).FormulaR1C1 = "=COUNTIFS(tbl_PayrollEngine7[PayPeriod],RC[-11])"
    ws.Range("M2:M" & lastRow).FormulaR1C1 = "=IFERROR(RC[-7]/RC[-1],0)"

    ' --- N:Q department rollup (latest period only) ---
    Dim depts As Variant
    depts = Array("Back Office", "Customer Service", "Finance", "HR", "IT", "Sales", "Technical Support", "Workforce Management")
    ws.Range("N1:Q1").Value = Array("Department", "GrossPay_Latest", "Headcount_Latest", "CostPerFTE_Latest")
    For i = 0 To UBound(depts)
        ws.Cells(2 + i, 14).Value = depts(i)
    Next i
    Dim lastDept As Long
    lastDept = 1 + (UBound(depts) + 1)
    ws.Range("O2:O" & lastDept).FormulaR1C1 = _
        "=SUMIFS(tbl_PayrollEngine7[GrossPay],tbl_PayrollEngine7[PayPeriod],Calc_Payroll!R" & lastRow & "C1,tbl_PayrollEngine7[Department],RC[-1])"
    ws.Range("P2:P" & lastDept).FormulaR1C1 = _
        "=COUNTIFS(tbl_PayrollEngine7[PayPeriod],Calc_Payroll!R" & lastRow & "C1,tbl_PayrollEngine7[Department],RC[-2])"
    ws.Range("Q2:Q" & lastDept).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-1],0)"

    ' --- S:V site rollup (latest period only) ---
    Dim sites As Variant
    sites = Array("Manila", "Davao", "Cebu")
    ws.Range("S1:V1").Value = Array("Site", "GrossPay_Latest", "Headcount_Latest", "CostPerFTE_Latest")
    For i = 0 To UBound(sites)
        ws.Cells(2 + i, 19).Value = sites(i)
    Next i
    Dim lastSite As Long
    lastSite = 1 + (UBound(sites) + 1)
    ws.Range("T2:T" & lastSite).FormulaR1C1 = _
        "=SUMIFS(tbl_PayrollEngine7[GrossPay],tbl_PayrollEngine7[PayPeriod],Calc_Payroll!R" & lastRow & "C1,tbl_PayrollEngine7[Site],RC[-1])"
    ws.Range("U2:U" & lastSite).FormulaR1C1 = _
        "=COUNTIFS(tbl_PayrollEngine7[PayPeriod],Calc_Payroll!R" & lastRow & "C1,tbl_PayrollEngine7[Site],RC[-2])"
    ws.Range("V2:V" & lastSite).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-1],0)"

    ' --- X:Y statutory deduction mix (latest period only) ---
    ws.Range("X1").Value = "StatutoryItem": ws.Range("Y1").Value = "Amount_Latest"
    ws.Range("X2").Value = "SSS": ws.Range("Y2").Formula = "=G" & lastRow
    ws.Range("X3").Value = "PhilHealth": ws.Range("Y3").Formula = "=H" & lastRow
    ws.Range("X4").Value = "Pag-IBIG": ws.Range("Y4").Formula = "=I" & lastRow
    ws.Range("X5").Value = "Withholding Tax": ws.Range("Y5").Formula = "=J" & lastRow

    ws.Calculate
End Sub

'=====================================================================================
' GENERIC HELPERS
'=====================================================================================
Private Sub ClearSheet(ws As Worksheet)
    Dim shp As Shape
    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop
    Dim pt As Object
    On Error Resume Next
    For Each pt In ws.PivotTables
        pt.TableRange2.Clear
    Next pt
    On Error GoTo 0
    ws.Cells.Clear
    ws.Columns.ColumnWidth = 9
    ws.Rows.RowHeight = 15
End Sub

Private Sub AddRefreshButton(ws As Worksheet)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("A1").Left + ws.Range("A1:P1").Width - 150, 6, 140, 22)
    With btn
        .Fill.ForeColor.RGB = C_TEAL
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "Refresh Dashboards"
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = C_WHITE
        .OnAction = "Build_BPO_Dashboards"
    End With
End Sub

Private Sub AddKPICard(ws As Worksheet, lft As Double, tp As Double, wd As Double, ht As Double, _
                        lbl As String, valTxt As String, accent As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, lft, tp, wd, ht)
    With shp
        .Fill.ForeColor.RGB = C_WHITE
        .Line.ForeColor.RGB = RGB(222, 226, 232)
        .Line.Weight = 1
        .Adjustments.Item(1) = 0.09
        .TextFrame.MarginLeft = 9
        .TextFrame.MarginRight = 4
        .TextFrame.MarginTop = 6
        .TextFrame.MarginBottom = 4
        .TextFrame.Characters.Text = UCase(lbl) & vbNewLine & valTxt
        With .TextFrame.Characters(1, Len(lbl)).Font
            .Size = 8.5
            .Bold = True
            .Color = C_GREY
            .Name = "Calibri"
        End With
        With .TextFrame.Characters(Len(lbl) + 2, Len(valTxt)).Font
            .Size = 19
            .Bold = True
            .Color = accent
            .Name = "Calibri"
        End With
    End With
    Dim bar As Shape
    Set bar = ws.Shapes.AddShape(msoShapeRectangle, lft, tp, 4, ht)
    bar.Fill.ForeColor.RGB = accent
    bar.Line.Visible = msoFalse
End Sub

Private Sub AddSectionLabel(ws As Worksheet, lft As Double, tp As Double, txt As String)
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, lft, tp, 380, 16)
    With shp
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = txt
        .TextFrame.Characters.Font.Size = 10.5
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = C_NAVY
        .TextFrame.Characters.Font.Name = "Calibri"
    End With
End Sub

Private Sub AddLineSeries(cht As Chart, seriesName As String, catRange As Range, valRange As Range, _
                           lineColor As Long, lineWeight As Double, dashStyle As XlLineStyle)
    Dim s As Series
    Set s = cht.SeriesCollection.NewSeries
    s.Name = seriesName
    s.XValues = catRange
    s.Values = valRange
    s.ChartType = xlLine
    s.Format.Line.ForeColor.RGB = lineColor
    s.Format.Line.Weight = lineWeight
    s.Border.LineStyle = dashStyle
    s.MarkerStyle = xlMarkerStyleNone
End Sub

Private Sub AddBarSeries(cht As Chart, seriesName As String, catRange As Range, valRange As Range, fillColor As Long)
    Dim s As Series
    Set s = cht.SeriesCollection.NewSeries
    s.Name = seriesName
    s.XValues = catRange
    s.Values = valRange
    s.Format.Fill.ForeColor.RGB = fillColor
    s.Format.Line.Visible = msoFalse
End Sub

Private Sub FormatChart(cht As Chart, titleText As String, showLegend As Boolean)
    With cht
        .ChartArea.Format.Fill.ForeColor.RGB = C_WHITE
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.ForeColor.RGB = C_WHITE
        .PlotArea.Format.Line.Visible = msoFalse
        If titleText <> "" Then
            .HasTitle = True
            .ChartTitle.Text = titleText
            .ChartTitle.Font.Size = 11
        Else
            .HasTitle = False
        End If
        .HasLegend = showLegend
        If showLegend Then
            .Legend.Position = xlLegendPositionBottom
            .Legend.Font.Size = 8.5
        End If
        On Error Resume Next
        .Axes(xlValue).Format.Line.ForeColor.RGB = RGB(220, 223, 228)
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(235, 237, 240)
        .Axes(xlCategory).HasMajorGridlines = False
        On Error GoTo 0
    End With
End Sub

'=====================================================================================
' DASHBOARD_EXECUTIVE
'=====================================================================================
Private Sub BuildExecutiveDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard_Executive")
    ClearSheet ws

    Dim ce As Worksheet, aa As Worksheet, hf As Worksheet
    Set ce = ThisWorkbook.Sheets(CALC_EXEC_SHEET)
    Set aa = ThisWorkbook.Sheets("Attrition_Analysis")
    Set hf = ThisWorkbook.Sheets("Headcount_Forecasts")

    ' ---- Title banner ----
    With ws.Range("A1:T3")
        .Merge
        .Interior.Color = C_NAVY
        .Font.Color = C_WHITE
        .Font.Name = "Calibri"
        .Font.Size = 20
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .Value = "PH WORKFORCE ANALYTICS  |  EXECUTIVE DASHBOARD"
    End With
    ws.Rows("1:3").RowHeight = 12
    With ws.Range("A4:T4")
        .Merge
        .Interior.Color = C_NAVY_DARK
        .Font.Color = RGB(200, 210, 225)
        .Font.Size = 10
        .Font.Italic = True
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Value = "Data through Dec-2025  |  Sites: Manila, Davao, Cebu  |  Last refreshed: " & Format(Now, "mmm dd, yyyy hh:mm AM/PM")
    End With

    AddRefreshButton ws

    ' ---- KPI values (computed once, current as of build time) ----
    Dim attSheet As Worksheet
    Set attSheet = ThisWorkbook.Sheets("Attendance")

    Dim activeHC As Double, attrRate As Double, avgTenure As Double
    Dim absRate As Double, otHours As Double, volShare As Double, volTot As Double, involTot As Double

    activeHC = aa.Range("A2").Value
    attrRate = aa.Range("C2").Value
    avgTenure = aa.Range("E2").Value
    absRate = Application.WorksheetFunction.CountIfs(attSheet.Range("G2:G183423"), "Yes") / _
              Application.WorksheetFunction.CountA(attSheet.Range("G2:G183423"))
    otHours = Application.WorksheetFunction.Sum(attSheet.Range("J2:J183423"))
    volTot = Application.WorksheetFunction.Sum(ce.Range("C2:C49"))
    involTot = Application.WorksheetFunction.Sum(ce.Range("D2:D49"))
    volShare = volTot / (volTot + involTot)

    ' ---- KPI cards ----
    Dim cardTop As Double: cardTop = 46
    Dim cardW As Double: cardW = 122
    Dim cardH As Double: cardH = 66
    Dim gp As Double: gp = 10
    Dim ls As Double: ls = 8

    AddKPICard ws, ls + 0 * (cardW + gp), cardTop, cardW, cardH, "Active Headcount", Format(activeHC, "#,##0"), C_TEAL
    AddKPICard ws, ls + 1 * (cardW + gp), cardTop, cardW, cardH, "Attrition Rate (TTM)", Format(attrRate, "0.0%"), C_RED
    AddKPICard ws, ls + 2 * (cardW + gp), cardTop, cardW, cardH, "Avg Tenure - Active", Format(avgTenure, "0.0") & " mo", C_NAVY
    AddKPICard ws, ls + 3 * (cardW + gp), cardTop, cardW, cardH, "Absenteeism Rate", Format(absRate, "0.0%"), C_AMBER
    AddKPICard ws, ls + 4 * (cardW + gp), cardTop, cardW, cardH, "Total OT Hours (2025)", Format(otHours, "#,##0"), C_TEAL
    AddKPICard ws, ls + 5 * (cardW + gp), cardTop, cardW, cardH, "Voluntary Attrition Share", Format(volShare, "0.0%"), C_GREEN

    ' ---- Highest-risk segment callout ----
    ws.Range("B18").Value = "Highest attrition-risk segment (TTM): " & aa.Range("I2").Value & _
        "     |     Total exits (48-month trailing window): " & Format(volTot + involTot, "#,##0") & _
        "  (" & Format(volTot, "#,##0") & " voluntary / " & Format(involTot, "#,##0") & " involuntary)"
    With ws.Range("B18:T18")
        .Font.Size = 10
        .Font.Italic = True
        .Font.Color = C_GREY
    End With

    ' ---- Chart row 1 ----
    Dim chTop As Double: chTop = 128
    Dim chH As Double: chH = 210
    Dim chW As Double: chW = 384
    Dim chGap As Double: chGap = 14

    AddSectionLabel ws, ls, chTop - 16, "HEADCOUNT TREND & FORECAST (w/ 95% CI)"
    Dim c1 As Chart
    Set c1 = ws.ChartObjects.Add(ls, chTop, chW, chH).Chart
    c1.ChartType = xlLine
    AddLineSeries c1, "Actual Headcount", hf.Range("I2:I61"), hf.Range("J2:J61"), C_NAVY, 2.5, xlContinuous
    AddLineSeries c1, "Forecast", hf.Range("I2:I61"), hf.Range("K2:K61"), C_TEAL, 2.25, xlDash
    AddLineSeries c1, "Upper CI", hf.Range("I2:I61"), hf.Range("L2:L61"), RGB(190, 195, 205), 1, xlDot
    AddLineSeries c1, "Lower CI", hf.Range("I2:I61"), hf.Range("M2:M61"), RGB(190, 195, 205), 1, xlDot
    FormatChart c1, "", True
    c1.Axes(xlCategory).CategoryType = xlTimeScale
    c1.Axes(xlCategory).BaseUnit = xlMonths
    c1.Axes(xlCategory).TickLabels.NumberFormat = "mmm-yy"

    AddSectionLabel ws, ls + chW + chGap, chTop - 16, "ATTRITION RATE TREND (2022-2025)"
    Dim c2 As Chart
    Set c2 = ws.ChartObjects.Add(ls + chW + chGap, chTop, chW, chH).Chart
    c2.ChartType = xlLine
    AddLineSeries c2, "Attrition Rate", aa.Range("A7:A54"), aa.Range("C7:C54"), C_RED, 2.5, xlContinuous
    FormatChart c2, "", False
    c2.Axes(xlValue).TickLabels.NumberFormat = "0%"

    ' ---- Chart row 2 ----
    chTop = chTop + chH + 40

    AddSectionLabel ws, ls, chTop - 16, "ATTRITION RATE BY DEPARTMENT x SHIFT (TTM)"
    Dim c3 As Chart
    Set c3 = ws.ChartObjects.Add(ls, chTop, chW, chH).Chart
    c3.ChartType = xlColumnClustered
    AddBarSeries c3, "Day", aa.Range("A108:A115"), aa.Range("B108:B115"), C_TEAL
    AddBarSeries c3, "Mid", aa.Range("A108:A115"), aa.Range("D108:D115"), C_AMBER
    AddBarSeries c3, "Night", aa.Range("A108:A115"), aa.Range("F108:F115"), C_NAVY
    FormatChart c3, "", True
    c3.Axes(xlValue).TickLabels.NumberFormat = "0%"
    On Error Resume Next
    c3.Axes(xlCategory).TickLabels.Orientation = 45
    On Error GoTo 0

    AddSectionLabel ws, ls + chW + chGap, chTop - 16, "VOLUNTARY vs INVOLUNTARY EXITS BY MONTH"
    Dim c4 As Chart
    Set c4 = ws.ChartObjects.Add(ls + chW + chGap, chTop, chW, chH).Chart
    c4.ChartType = xlColumnStacked
    AddBarSeries c4, "Voluntary", ce.Range("A2:A49"), ce.Range("C2:C49"), C_RED
    AddBarSeries c4, "Involuntary / Planned", ce.Range("A2:A49"), ce.Range("D2:D49"), RGB(190, 195, 205)
    FormatChart c4, "", True
    c4.Axes(xlCategory).CategoryType = xlTimeScale
    c4.Axes(xlCategory).BaseUnit = xlMonths
    c4.Axes(xlCategory).TickLabels.NumberFormat = "mmm-yy"

    ' ---- Chart row 3: cohort retention + flight risk table ----
    chTop = chTop + chH + 40

    AddSectionLabel ws, ls, chTop - 16, "NEW-HIRE RETENTION BY COHORT (2022-2023 HIRES)"
    Dim c5 As Chart
    Set c5 = ws.ChartObjects.Add(ls, chTop, chW, chH).Chart
    c5.ChartType = xlLine
    AddLineSeries c5, "Active @ 3mo", aa.Range("A73:A96"), aa.Range("C73:C96"), RGB(190, 195, 205), 1.75, xlContinuous
    AddLineSeries c5, "Active @ 6mo", aa.Range("A73:A96"), aa.Range("D73:D96"), C_AMBER, 1.75, xlContinuous
    AddLineSeries c5, "Active @ 12mo", aa.Range("A73:A96"), aa.Range("E73:E96"), C_RED, 2.25, xlContinuous
    FormatChart c5, "", True
    c5.Axes(xlValue).TickLabels.NumberFormat = "0%"
    c5.Axes(xlCategory).CategoryType = xlTimeScale
    c5.Axes(xlCategory).BaseUnit = xlMonths
    c5.Axes(xlCategory).TickLabels.NumberFormat = "mmm-yy"

    AddSectionLabel ws, ls + chW + chGap, chTop - 16, "FLIGHT RISK WATCHLIST (Active, <6mo tenure, Night, elevated OT/lateness)"
    BuildFlightRiskTable ws, aa, 44, 10   ' topRow=44 (~chTop pts/15), leftCol=10 ("J")

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Tab.Color = C_TEAL
    ws.Range("A1").Select
End Sub

Private Sub BuildFlightRiskTable(ws As Worksheet, aa As Worksheet, topRow As Long, leftCol As Long)
    Dim nCols As Long: nCols = 7   ' EmployeeID, FullName, Department, DateHired, ShiftType, RecentOTHours, RecentLateInstances
    Dim nRows As Long: nRows = 14  ' header + 13 employees (Attrition_Analysis rows 121:135)

    Dim r As Long, c As Long
    For r = 0 To nRows - 1
        For c = 0 To nCols - 1
            ws.Cells(topRow + r, leftCol + c).Formula = _
                "=" & aa.Cells(121 + r, 1 + c).Address(False, False, xlA1, True)
        Next c
    Next r

    Dim hdrRng As Range
    Set hdrRng = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(topRow, leftCol + nCols - 1))
    With hdrRng
        .Interior.Color = C_NAVY
        .Font.Color = C_WHITE
        .Font.Bold = True
        .Font.Size = 9
    End With

    Dim bodyRng As Range
    Set bodyRng = ws.Range(ws.Cells(topRow + 1, leftCol), ws.Cells(topRow + nRows - 1, leftCol + nCols - 1))
    With bodyRng
        .Font.Size = 9
        .Borders.Color = RGB(225, 228, 232)
    End With
    ws.Range(ws.Cells(topRow, leftCol + 3), ws.Cells(topRow + nRows - 1, leftCol + 3)).NumberFormat = "mmm-yyyy"

    ' Zebra striping
    For r = 1 To nRows - 1
        If r Mod 2 = 0 Then
            ws.Range(ws.Cells(topRow + r, leftCol), ws.Cells(topRow + r, leftCol + nCols - 1)).Interior.Color = C_GREY_LIGHT
        End If
    Next r

    ' Highlight elevated OT hours (col index 5, 0-based -> leftCol+5) and late instances (leftCol+6)
    Dim otRng As Range, lateRng As Range
    Set otRng = ws.Range(ws.Cells(topRow + 1, leftCol + 5), ws.Cells(topRow + nRows - 1, leftCol + 5))
    Set lateRng = ws.Range(ws.Cells(topRow + 1, leftCol + 6), ws.Cells(topRow + nRows - 1, leftCol + 6))
    otRng.FormatConditions.AddColorScale ColorScaleType:=2
    otRng.FormatConditions(otRng.FormatConditions.Count).ColorScaleCriteria(1).FormatColor.Color = C_WHITE
    otRng.FormatConditions(otRng.FormatConditions.Count).ColorScaleCriteria(2).FormatColor.Color = C_RED
    lateRng.FormatConditions.AddColorScale ColorScaleType:=2
    lateRng.FormatConditions(lateRng.FormatConditions.Count).ColorScaleCriteria(1).FormatColor.Color = C_WHITE
    lateRng.FormatConditions(lateRng.FormatConditions.Count).ColorScaleCriteria(2).FormatColor.Color = C_RED

    ws.Range(ws.Cells(topRow, leftCol), ws.Cells(topRow + nRows - 1, leftCol + nCols - 1)).Columns.AutoFit
End Sub

Private Sub AddPieSeries(cht As Chart, catRange As Range, valRange As Range, colors As Variant)
    Dim s As Series
    Set s = cht.SeriesCollection.NewSeries
    s.XValues = catRange
    s.Values = valRange
    Dim i As Long
    For i = 1 To s.Points.Count
        s.Points(i).Format.Fill.ForeColor.RGB = colors((i - 1) Mod (UBound(colors) + 1))
        s.Points(i).Format.Line.ForeColor.RGB = C_WHITE
        s.Points(i).Format.Line.Weight = 1.5
    Next i
    s.HasDataLabels = True
    s.DataLabels.ShowPercentage = True
    s.DataLabels.ShowValue = False
    s.DataLabels.Font.Size = 8
    s.DataLabels.Font.Color = C_WHITE
    s.DataLabels.Font.Bold = True
End Sub

'=====================================================================================
' DASHBOARD_PAYROLL
'=====================================================================================
Private Sub BuildPayrollDashboard()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Dashboard_Payroll")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Dashboard_Executive"))
        ws.Name = "Dashboard_Payroll"
    End If
    ClearSheet ws

    Dim cp As Worksheet
    Set cp = ThisWorkbook.Sheets(CALC_PAYROLL_SHEET)
    Dim lastRow As Long: lastRow = 1 + N_PERIODS   ' row 49 = latest period (2025-12)
    Dim yoyRow As Long: yoyRow = lastRow - 12       ' row 37 = 2024-12

    ' ---- Title banner ----
    With ws.Range("A1:T3")
        .Merge
        .Interior.Color = C_NAVY
        .Font.Color = C_WHITE
        .Font.Name = "Calibri"
        .Font.Size = 20
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .Value = "PH WORKFORCE ANALYTICS  |  PAYROLL DASHBOARD"
    End With
    ws.Rows("1:3").RowHeight = 12
    With ws.Range("A4:T4")
        .Merge
        .Interior.Color = C_NAVY_DARK
        .Font.Color = RGB(200, 210, 225)
        .Font.Size = 10
        .Font.Italic = True
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Value = "Latest pay period: Dec-2025  |  Currency: PHP  |  Last refreshed: " & Format(Now, "mmm dd, yyyy hh:mm AM/PM")
    End With

    AddRefreshButton ws

    ' ---- KPI values ----
    Dim grossLatest As Double, netLatest As Double, costPerFTE As Double
    Dim otPct As Double, statTotal As Double, yoyGrowth As Double
    grossLatest = cp.Cells(lastRow, 6).Value
    netLatest = cp.Cells(lastRow, 11).Value
    costPerFTE = cp.Cells(lastRow, 13).Value
    otPct = cp.Cells(lastRow, 3).Value / grossLatest
    statTotal = cp.Cells(lastRow, 7).Value + cp.Cells(lastRow, 8).Value + cp.Cells(lastRow, 9).Value
    yoyGrowth = (grossLatest - cp.Cells(yoyRow, 6).Value) / cp.Cells(yoyRow, 6).Value

    ' ---- KPI cards ----
    Dim cardTop As Double: cardTop = 46
    Dim cardW As Double: cardW = 122
    Dim cardH As Double: cardH = 66
    Dim gp As Double: gp = 10
    Dim ls As Double: ls = 8

    AddKPICard ws, ls + 0 * (cardW + gp), cardTop, cardW, cardH, "Gross Payroll (Latest Mo)", "PHP " & Format(grossLatest, "#,##0,") & "K", C_TEAL
    AddKPICard ws, ls + 1 * (cardW + gp), cardTop, cardW, cardH, "Net Payroll (Latest Mo)", "PHP " & Format(netLatest, "#,##0,") & "K", C_NAVY
    AddKPICard ws, ls + 2 * (cardW + gp), cardTop, cardW, cardH, "Avg Cost per FTE", "PHP " & Format(costPerFTE, "#,##0"), C_GREEN
    AddKPICard ws, ls + 3 * (cardW + gp), cardTop, cardW, cardH, "OT Cost % of Gross", Format(otPct, "0.0%"), C_AMBER
    AddKPICard ws, ls + 4 * (cardW + gp), cardTop, cardW, cardH, "Statutory Deductions (Employee)", "PHP " & Format(statTotal, "#,##0,") & "K", C_TEAL
    AddKPICard ws, ls + 5 * (cardW + gp), cardTop, cardW, cardH, "Payroll YoY Growth", Format(yoyGrowth, "0.0%"), IIf(yoyGrowth >= 0, C_GREEN, C_RED)

    ws.Range("B18").Value = "Statutory deductions shown are the employee-side share only (SSS + PhilHealth + Pag-IBIG). " & _
        "Employer-side contributions are not included in Gross/Net Payroll figures above."
    With ws.Range("B18:T18")
        .Font.Size = 10
        .Font.Italic = True
        .Font.Color = C_GREY
    End With

    ' ---- Chart row 1 ----
    Dim chTop As Double: chTop = 128
    Dim chH As Double: chH = 210
    Dim chW As Double: chW = 384
    Dim chGap As Double: chGap = 14
    Dim trendStartRow As Long: trendStartRow = lastRow - 23   ' last 24 periods

    AddSectionLabel ws, ls, chTop - 16, "PAYROLL COST TREND BY COMPONENT (Last 24 Months)"
    Dim c1 As Chart
    Set c1 = ws.ChartObjects.Add(ls, chTop, chW, chH).Chart
    c1.ChartType = xlColumnStacked
    AddBarSeries c1, "Basic Pay", cp.Range(cp.Cells(trendStartRow, 1), cp.Cells(lastRow, 1)), cp.Range(cp.Cells(trendStartRow, 2), cp.Cells(lastRow, 2)), C_NAVY
    AddBarSeries c1, "OT Pay", cp.Range(cp.Cells(trendStartRow, 1), cp.Cells(lastRow, 1)), cp.Range(cp.Cells(trendStartRow, 3), cp.Cells(lastRow, 3)), C_AMBER
    AddBarSeries c1, "Night Diff Pay", cp.Range(cp.Cells(trendStartRow, 1), cp.Cells(lastRow, 1)), cp.Range(cp.Cells(trendStartRow, 4), cp.Cells(lastRow, 4)), C_TEAL
    AddBarSeries c1, "Holiday Pay", cp.Range(cp.Cells(trendStartRow, 1), cp.Cells(lastRow, 1)), cp.Range(cp.Cells(trendStartRow, 5), cp.Cells(lastRow, 5)), RGB(190, 195, 205)
    FormatChart c1, "", True
    c1.Axes(xlCategory).CategoryType = xlTimeScale
    c1.Axes(xlCategory).BaseUnit = xlMonths
    c1.Axes(xlCategory).TickLabels.NumberFormat = "mmm-yy"
    c1.Axes(xlValue).TickLabels.NumberFormat = "#,##0,""K"""

    AddSectionLabel ws, ls + chW + chGap, chTop - 16, "GROSS vs NET PAYROLL TREND (Last 24 Months)"
    Dim c2 As Chart
    Set c2 = ws.ChartObjects.Add(ls + chW + chGap, chTop, chW, chH).Chart
    c2.ChartType = xlLine
    AddLineSeries c2, "Gross Payroll", cp.Range(cp.Cells(trendStartRow, 1), cp.Cells(lastRow, 1)), cp.Range(cp.Cells(trendStartRow, 6), cp.Cells(lastRow, 6)), C_NAVY, 2.25, xlContinuous
    AddLineSeries c2, "Net Payroll", cp.Range(cp.Cells(trendStartRow, 1), cp.Cells(lastRow, 1)), cp.Range(cp.Cells(trendStartRow, 11), cp.Cells(lastRow, 11)), C_TEAL, 2.25, xlContinuous
    FormatChart c2, "", True
    c2.Axes(xlCategory).CategoryType = xlTimeScale
    c2.Axes(xlCategory).BaseUnit = xlMonths
    c2.Axes(xlCategory).TickLabels.NumberFormat = "mmm-yy"
    c2.Axes(xlValue).TickLabels.NumberFormat = "#,##0,""K"""

    ' ---- Chart row 2 ----
    chTop = chTop + chH + 40

    AddSectionLabel ws, ls, chTop - 16, "PAYROLL COST BY DEPARTMENT (Latest Month)"
    Dim c3 As Chart
    Set c3 = ws.ChartObjects.Add(ls, chTop, chW, chH).Chart
    c3.ChartType = xlBarClustered
    AddBarSeries c3, "Gross Pay", cp.Range("N2:N9"), cp.Range("O2:O9"), C_TEAL
    FormatChart c3, "", False
    c3.Axes(xlValue).TickLabels.NumberFormat = "#,##0,""K"""

    AddSectionLabel ws, ls + chW + chGap, chTop - 16, "PAYROLL COST BY SITE (Latest Month)"
    Dim c4 As Chart
    Set c4 = ws.ChartObjects.Add(ls + chW + chGap, chTop, chW, chH).Chart
    c4.ChartType = xlColumnClustered
    AddBarSeries c4, "Gross Pay", cp.Range("S2:S4"), cp.Range("T2:T4"), C_NAVY
    FormatChart c4, "", False
    c4.Axes(xlValue).TickLabels.NumberFormat = "#,##0,""K"""

    ' ---- Chart row 3 ----
    chTop = chTop + chH + 40

    AddSectionLabel ws, ls, chTop - 16, "STATUTORY DEDUCTIONS MIX - EMPLOYEE SHARE (Latest Month)"
    Dim c5 As Chart
    Set c5 = ws.ChartObjects.Add(ls, chTop, chW, chH).Chart
    c5.ChartType = xlPie
    AddPieSeries c5, cp.Range("X2:X5"), cp.Range("Y2:Y5"), Array(C_TEAL, C_NAVY, C_AMBER, RGB(190, 195, 205))
    FormatChart c5, "", True

    AddSectionLabel ws, ls + chW + chGap, chTop - 16, "COST PER FTE BY DEPARTMENT (Latest Month)"
    Dim c6 As Chart
    Set c6 = ws.ChartObjects.Add(ls + chW + chGap, chTop, chW, chH).Chart
    c6.ChartType = xlBarClustered
    AddBarSeries c6, "Cost per FTE", cp.Range("N2:N9"), cp.Range("Q2:Q9"), C_GREEN
    FormatChart c6, "", False
    c6.Axes(xlValue).TickLabels.NumberFormat = "#,##0"

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Tab.Color = C_NAVY
    ws.Range("A1").Select
End Sub
