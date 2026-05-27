Attribute VB_Name = "RefreshModule"
'=============================================================
'  PROMPT 3 - AUTOMATED DASHBOARD BUILDER
'  Sheets required: Data, Pivot (must already exist)
'  Dashboard sheet is auto-created/rebuilt each time
'
'  MACRO TO ASSIGN TO BUTTON: BuildDashboard
'  MACRO TO ASSIGN TO REFRESH BUTTON: RefreshDashboard
'
'  Data columns (from your file):
'  A=Order ID | B=Order Date | C=Year | D=Month | E=Day
'  F=Segment  | G=Category   | H=Sub-Category
'  I=Region   | J=Sales      | K=Profit | L=Quantity | M=Discount
'=============================================================

Option Explicit

'-------------------------------------------------------------
'  ENTRY POINT 1 - First-time build
'  Assign this to your first button: BuildDashboard
'-------------------------------------------------------------
Sub BuildDashboard()
    Call DashboardMain
    MsgBox "Dashboard created successfully!", vbInformation, "Done"
End Sub

'-------------------------------------------------------------
'  ENTRY POINT 2 - Refresh button (rebuilds in-place)
'  Assign this to your Refresh button: RefreshDashboard
'-------------------------------------------------------------
Sub RefreshDashboard()
    ' Refresh all pivot tables first
    Dim ws As Worksheet
    Dim pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            On Error Resume Next
            pt.RefreshTable
            On Error GoTo 0
        Next pt
    Next ws

    Call DashboardMain
    MsgBox "Dashboard refreshed successfully!", vbInformation, "Refreshed"
End Sub

'-------------------------------------------------------------
'  CORE LOGIC
'  Finds or creates the Dashboard sheet, clears it,
'  then rebuilds every section in-place
'-------------------------------------------------------------
Private Sub DashboardMain()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' --- Validate Pivot sheet exists ---
    Dim wsPvt As Worksheet
    Set wsPvt = Nothing
    On Error Resume Next
    Set wsPvt = ThisWorkbook.Sheets("Pivot")
    On Error GoTo 0

    If wsPvt Is Nothing Then
        MsgBox "Sheet 'Pivot' not found! Please create the Pivot sheet first.", vbCritical
        GoTo CleanUp
    End If

    ' --- Get existing Dashboard sheet OR create it once ---
    Dim wsDash As Worksheet
    Set wsDash = Nothing
    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    On Error GoTo 0

    If wsDash Is Nothing Then
        Set wsDash = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDash.Name = "Dashboard"
        wsDash.Tab.Color = RGB(94, 59, 183)
    End If

    ' --- Clear existing content and all charts ---
    wsDash.Cells.Clear
    wsDash.Cells.Interior.ColorIndex = xlNone

    Dim chtObj As ChartObject
    For Each chtObj In wsDash.ChartObjects
        chtObj.Delete
    Next chtObj

    ' --- Remove old Refresh button shape (will be re-added) ---
    Dim shp As Shape
    For Each shp In wsDash.Shapes
        If shp.Name = "btnRefreshDashboard" Then
            shp.Delete
            Exit For
        End If
    Next shp

    ' --- Rebuild all sections ---
    Call SetupDashboardLayout(wsDash)
    Call BuildHeader(wsDash)
    Call BuildKPICards(wsDash, wsPvt)
    Call CreateAllCharts(wsDash, wsPvt)
    Call AddMacroButton(wsDash)

    wsDash.Activate
    ActiveWindow.DisplayGridlines = False
    wsDash.Range("A1").Select

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'-------------------------------------------------------------
'  LAYOUT - Column widths and row heights
'-------------------------------------------------------------
Private Sub SetupDashboardLayout(ws As Worksheet)
    Dim c As Integer
    For c = 1 To 17
        ws.Columns(c).ColumnWidth = 9.5
    Next c

    ws.Rows("1:2").RowHeight = 20     ' Header
    ws.Rows("3:3").RowHeight = 5      ' Spacer
    ws.Rows("4:4").RowHeight = 16     ' KPI label row
    ws.Rows("5:5").RowHeight = 26     ' KPI value row
    ws.Rows("6:6").RowHeight = 6      ' KPI shadow row
    ws.Rows("7:7").RowHeight = 8      ' Spacer

    Dim r As Integer
    For r = 8 To 31
        ws.Rows(r).RowHeight = 15
    Next r

    ws.Range("A1:Q31").Interior.Color = RGB(248, 247, 255)
End Sub

'-------------------------------------------------------------
'  HEADER BANNER (Rows 1-2)
'-------------------------------------------------------------
Private Sub BuildHeader(ws As Worksheet)
    ws.Range("A1:Q2").Merge

    With ws.Range("A1")
        .Value = "SALES PERFORMANCE DASHBOARD"
        .Font.Name = "Calibri"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(94, 59, 183)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

'-------------------------------------------------------------
'  KPI CARDS (Rows 4-6) - 6 cards across columns B to M
'  Pulls values directly from Pivot sheet PivotTables
'  *** Adjust cell refs if your Pivot layout differs ***
'-------------------------------------------------------------
Private Sub BuildKPICards(ws As Worksheet, wsPvt As Worksheet)
    Dim purple     As Long: purple = RGB(94, 59, 183)
    Dim purpleDeep As Long: purpleDeep = RGB(60, 30, 130)
    Dim bgWhite    As Long: bgWhite = RGB(255, 255, 255)
    Dim shadow     As Long: shadow = RGB(210, 205, 230)
    Dim textDark   As Long: textDark = RGB(30, 30, 30)

    ' --- KPI Definitions ---
    Dim kpiLabels(5) As String
    Dim kpiFmlas(5)  As String
    Dim kpiFmts(5)   As String
    Dim kpiCols(5)   As Integer

    kpiLabels(0) = "Total Sales"
    kpiLabels(1) = "Total Profit"
    kpiLabels(2) = "Total Orders"
    kpiLabels(3) = "Total Qty Sold"
    kpiLabels(4) = "Avg Discount"
    kpiLabels(5) = "Profit Margin"

    ' Formulas - using direct SUM/COUNTA on Data sheet
    ' (safe fallback if Pivot layout differs)
    kpiFmlas(0) = "=SUM(Data!J2:J10000)"
    kpiFmlas(1) = "=SUM(Data!K2:K10000)"
    kpiFmlas(2) = "=COUNTA(Data!A2:A10000)"
    kpiFmlas(3) = "=SUM(Data!L2:L10000)"
    kpiFmlas(4) = "=AVERAGE(Data!M2:M10000)"
    kpiFmlas(5) = "=D5/B5"

    kpiFmts(0) = "[$$-en-US]#,##0"
    kpiFmts(1) = "[$$-en-US]#,##0"
    kpiFmts(2) = "#,##0"
    kpiFmts(3) = "#,##0"
    kpiFmts(4) = "0.0%"
    kpiFmts(5) = "0.0%"

    ' Starting column for each card (2 cols wide each, 1 col gap)
    kpiCols(0) = 2  ' B
    kpiCols(1) = 4  ' D
    kpiCols(2) = 6  ' F
    kpiCols(3) = 8  ' H
    kpiCols(4) = 10 ' J
    kpiCols(5) = 12 ' L

    Dim k As Integer
    For k = 0 To 5
        Dim c As Integer: c = kpiCols(k)

        ' --- Row 4: Purple header label ---
        With ws.Range(ws.Cells(4, c), ws.Cells(4, c + 1))
            .Merge
            .Value = kpiLabels(k)
            .Interior.Color = purple
            .Font.Bold = True
            .Font.Size = 9
            .Font.Name = "Calibri"
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).Color = purpleDeep
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeLeft).Color = purpleDeep
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeRight).Color = purpleDeep
        End With

        ' --- Row 5: White value cell (live formula) ---
        With ws.Range(ws.Cells(5, c), ws.Cells(5, c + 1))
            .Merge
            .Formula = kpiFmlas(k)
            .NumberFormat = kpiFmts(k)
            .Interior.Color = bgWhite
            .Font.Bold = True
            .Font.Size = 13
            .Font.Name = "Calibri"
            .Font.Color = textDark
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeLeft).Color = purpleDeep
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeRight).Color = purpleDeep
        End With

        ' --- Row 6: Shadow strip ---
        With ws.Range(ws.Cells(6, c), ws.Cells(6, c + 1))
            .Merge
            .Interior.Color = shadow
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).Color = purpleDeep
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeBottom).Color = purpleDeep
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeLeft).Color = purpleDeep
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeRight).Color = purpleDeep
        End With
    Next k
End Sub

'-------------------------------------------------------------
'  CREATE ALL 6 CHARTS
'  Top row: 3 charts across (rows 8-18)
'  Bottom row: 3 charts across (rows 19-31)
'  *** Adjust wsPvt.Range() refs to match your Pivot layout ***
'-------------------------------------------------------------
Private Sub CreateAllCharts(ws As Worksheet, wsPvt As Worksheet)
    Dim cPurple As Long: cPurple = RGB(94, 59, 183)
    Dim cPink   As Long: cPink = RGB(197, 90, 240)
    Dim cOrange As Long: cOrange = RGB(255, 140, 0)
    Dim cBlue   As Long: cBlue = RGB(68, 114, 196)
    Dim cWhite  As Long: cWhite = RGB(255, 255, 255)

    ' Grid anchor positions
    Dim L1 As Double: L1 = ws.Cells(8, 1).Left + 2
    Dim L2 As Double: L2 = ws.Cells(8, 6).Left + 2
    Dim L3 As Double: L3 = ws.Cells(8, 12).Left + 2
    Dim T1 As Double: T1 = ws.Cells(8, 1).Top + 2
    Dim T2 As Double: T2 = ws.Cells(19, 1).Top + 2

    Dim CW1 As Double: CW1 = ws.Cells(8, 6).Left - ws.Cells(8, 1).Left - 4
    Dim CW2 As Double: CW2 = ws.Cells(8, 12).Left - ws.Cells(8, 6).Left - 4
    Dim CW3 As Double: CW3 = ws.Cells(8, 17).Left + ws.Columns(17).Width - ws.Cells(8, 12).Left - 4
    Dim CH  As Double: CH = ws.Cells(19, 1).Top - ws.Cells(8, 1).Top - 4

    Dim cht As ChartObject
    Dim sc  As Integer
    Dim p   As Integer

    '=== CHART 1: Sales by Segment (Doughnut) - Top Left ===
    '    Pivot range: PT_Segment at A2 -> rows A3:B5
    Set cht = ws.ChartObjects.Add(Left:=L1, Top:=T1, Width:=CW1, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.Range("A3:B5")
        .chartType = xlDoughnut
        .HasTitle = True
        .chartTitle.Text = "Sales by Segment"
        Call ApplyTitleStyle(.chartTitle, cPurple)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        .SeriesCollection(1).Points(1).Interior.Color = cOrange
        .SeriesCollection(1).Points(2).Interior.Color = cPurple
        .SeriesCollection(1).Points(3).Interior.Color = cPink
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.ShowPercentage = True
        .SeriesCollection(1).DataLabels.ShowValue = False
        .SeriesCollection(1).DataLabels.Font.Size = 9
        .SeriesCollection(1).DataLabels.Font.Bold = True
        .SeriesCollection(1).DataLabels.Font.Color = cWhite
        On Error GoTo 0
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        .Legend.Font.Size = 8
        .ShowAllFieldButtons = False
    End With

    '=== CHART 2: Sales & Profit Trend (Line) - Top Centre ===
    '    Pivot range: PT_Trend at D2 -> D3:F46
    Set cht = ws.ChartObjects.Add(Left:=L2, Top:=T1, Width:=CW2, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.Range("D3:F46")
        .chartType = xlLineMarkers
        .HasTitle = True
        .chartTitle.Text = "Sales & Profit Trend"
        Call ApplyTitleStyle(.chartTitle, cPurple)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        With .SeriesCollection(1)
            .Name = "Total Sales"
            .Format.Line.ForeColor.RGB = cPurple
            .Format.Line.Weight = 2
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
        End With
        With .SeriesCollection(2)
            .Name = "Total Profit"
            .Format.Line.ForeColor.RGB = cPink
            .Format.Line.Weight = 2
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
        End With
        On Error GoTo 0
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 7
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With

    '=== CHART 3: Monthly Sales Performance (Column) - Top Right ===
    '    Pivot range: PT_Monthly at H2 -> H3:L15
    Set cht = ws.ChartObjects.Add(Left:=L3, Top:=T1, Width:=CW3, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.Range("H3:L15")
        .chartType = xlColumnClustered
        .HasTitle = True
        .chartTitle.Text = "Monthly Sales Performance"
        Call ApplyTitleStyle(.chartTitle, cPurple)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        Dim yearCols(3) As Long
        yearCols(0) = cPurple
        yearCols(1) = cPink
        yearCols(2) = cOrange
        yearCols(3) = cBlue
        For sc = 1 To .SeriesCollection.Count
            If sc <= 4 Then .SeriesCollection(sc).Interior.Color = yearCols(sc - 1)
        Next sc
        On Error GoTo 0
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 8
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With

    '=== CHART 4: Profit vs Discount (Column) - Bottom Left ===
    '    Pivot range: PT_Discount at U2 -> U3:V19
    Set cht = ws.ChartObjects.Add(Left:=L1, Top:=T2, Width:=CW1, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.Range("U3:V19")
        .chartType = xlColumnClustered
        .HasTitle = True
        .chartTitle.Text = "Profit vs Discount"
        Call ApplyTitleStyle(.chartTitle, cPurple)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        .SeriesCollection(1).Interior.Color = cPink
        .SeriesCollection(1).Name = "Total Profit"
        On Error GoTo 0
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 6
        .Axes(xlCategory).TickLabelSpacing = 1
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With

    '=== CHART 5: Sales by Category (Bar) - Bottom Centre ===
    '    Pivot range: PT_Category at O2 -> O3:P5
    Set cht = ws.ChartObjects.Add(Left:=L2, Top:=T2, Width:=CW2, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.Range("O3:P5")
        .chartType = xlBarClustered
        .HasTitle = True
        .chartTitle.Text = "Sales by Category"
        Call ApplyTitleStyle(.chartTitle, cPurple)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        .SeriesCollection(1).Points(1).Interior.Color = cPink
        .SeriesCollection(1).Points(2).Interior.Color = cPurple
        .SeriesCollection(1).Points(3).Interior.Color = cOrange
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.ShowValue = True
        .SeriesCollection(1).DataLabels.NumberFormat = "#,##0"
        .SeriesCollection(1).DataLabels.Font.Size = 8
        On Error GoTo 0
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 9
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With

    '=== CHART 6: Sales by Region (Column) - Bottom Right ===
    '    Pivot range: PT_Region at R2 -> R3:S6
    Set cht = ws.ChartObjects.Add(Left:=L3, Top:=T2, Width:=CW3, Height:=CH)
    With cht.Chart
        .SetSourceData Source:=wsPvt.Range("R3:S6")
        .chartType = xlColumnClustered
        .HasTitle = True
        .chartTitle.Text = "Sales by Region"
        Call ApplyTitleStyle(.chartTitle, cPurple)
        Call ApplyChartAreaStyle(cht.Chart)
        On Error Resume Next
        Dim regCols(3) As Long
        regCols(0) = cPurple
        regCols(1) = cPink
        regCols(2) = cOrange
        regCols(3) = cBlue
        For p = 1 To .SeriesCollection(1).Points.Count
            If p <= 4 Then .SeriesCollection(1).Points(p).Interior.Color = regCols(p - 1)
        Next p
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.ShowValue = True
        .SeriesCollection(1).DataLabels.NumberFormat = "$#,##0"
        .SeriesCollection(1).DataLabels.Font.Size = 8
        On Error GoTo 0
        .HasLegend = False
        .Axes(xlValue).MinimumScaleIsAuto = True
        .Axes(xlValue).MaximumScaleIsAuto = True
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 9
        Call RemoveGridlines(cht.Chart)
        .ShowAllFieldButtons = False
    End With
End Sub

'-------------------------------------------------------------
'  REFRESH BUTTON - placed top-right inside header
'-------------------------------------------------------------
Private Sub AddMacroButton(ws As Worksheet)
    Dim btnLeft   As Double: btnLeft = ws.Cells(1, 15).Left + 4
    Dim btnTop    As Double: btnTop = ws.Cells(1, 1).Top + 4
    Dim btnWidth  As Double: btnWidth = ws.Cells(1, 17).Left + ws.Columns(17).Width - ws.Cells(1, 15).Left - 8
    Dim btnHeight As Double: btnHeight = ws.Cells(3, 1).Top - ws.Cells(1, 1).Top - 8

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                  btnLeft, btnTop, btnWidth, btnHeight)
    With shp
        .Name = "btnRefreshDashboard"
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Solid
        .Line.ForeColor.RGB = RGB(94, 59, 183)
        .Line.Weight = 1.5
        .Line.Visible = msoTrue
        .TextFrame2.TextRange.Text = "Refresh Dashboard"
        With .TextFrame2.TextRange.Font
            .Bold = msoTrue
            .Size = 9
            .Fill.ForeColor.RGB = RGB(94, 59, 183)
        End With
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .OnAction = "RefreshDashboard"
    End With
End Sub

'-------------------------------------------------------------
'  HELPERS
'-------------------------------------------------------------
Private Sub ApplyTitleStyle(ct As chartTitle, titleColor As Long)
    ct.Font.Bold = True
    ct.Font.Color = titleColor
    ct.Font.Size = 11
    ct.Font.Name = "Calibri"
End Sub

Private Sub ApplyChartAreaStyle(cht As Chart)
    cht.ChartArea.Interior.Color = RGB(245, 242, 255)
    With cht.ChartArea.Border
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = RGB(94, 59, 183)
    End With
    cht.PlotArea.Interior.ColorIndex = xlNone
    On Error Resume Next
    cht.PlotArea.Border.LineStyle = xlNone
    On Error GoTo 0
End Sub

Private Sub RemoveGridlines(cht As Chart)
    On Error Resume Next
    cht.Axes(xlValue).HasMajorGridlines = False
    cht.Axes(xlValue).HasMinorGridlines = False
    cht.Axes(xlCategory).HasMajorGridlines = False
    cht.Axes(xlCategory).HasMinorGridlines = False
    On Error GoTo 0
End Sub

