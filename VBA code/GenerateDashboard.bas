Attribute VB_Name = "GenerateDashboard"
'==============================================================
' AUTOMATED DASHBOARD BUILDER - PURPLE THEME
' Sheets required: Data, Pivot, Dashboard
' Button: "Refresh Dashboard" linked to RefreshDashboard()
'==============================================================

Option Explicit

'--------------------------------------------------------------
' CONSTANTS - Adjust to match your actual Data/Pivot columns
'--------------------------------------------------------------
Const DATA_SHEET   As String = "Data"
Const PIVOT_SHEET  As String = "Pivot"
Const DASH_SHEET   As String = "Dashboard"

' Purple colour palette (RGB packed as Long)
Const CLR_DEEP     As Long = 4718592    ' #480060  - deep purple header
Const CLR_MID      As Long = 7340032    ' #700080  - mid-purple accent
Const CLR_LIGHT    As Long = 14540253   ' #DDA8DD  - light lavender fill
Const CLR_ACCENT   As Long = 10066329   ' #9933FF  - bright accent
Const CLR_WHITE    As Long = 16777215   ' #FFFFFF
Const CLR_LTGREY   As Long = 15921906   ' #F2F2F2  - card background
Const CLR_DKTEXT   As Long = 3947580    ' #3C003C  - dark text

'==============================================================
' MAIN ENTRY POINT  (link your button here)
'==============================================================
Public Sub RefreshDashboard()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo CleanUp

    ' 1. Refresh all pivot caches / data connections
    Call RefreshAllPivots

    ' 2. Guarantee Dashboard sheet exists and is blank
    Call PrepDashboardSheet

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(DASH_SHEET)

    ' 3. Pull KPI values from Pivot sheet
    Dim kpi(1 To 4) As Variant   ' label, value, prefix, format
    Call BuildKPIArray(kpi)

    ' 4. Draw every section
    Call DrawHeader(ws)
    Call DrawKPICards(ws, kpi)
    Call DrawTopCharts(ws)
    Call DrawBottomCharts(ws)
    Call DrawFooter(ws)

    ' 5. Final sheet tidy
    ws.Activate
    ws.Range("A1").Select
    ActiveWindow.DisplayGridlines = False

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    If Err.Number <> 0 Then
        MsgBox "Dashboard refresh failed: " & Err.Description, vbCritical
    Else
        MsgBox "Dashboard refreshed successfully!", vbInformation
    End If
End Sub

'==============================================================
' STEP 1  –  REFRESH PIVOTS & CONNECTIONS
'==============================================================
Private Sub RefreshAllPivots()
    Dim pc As PivotCache
    Dim conn As WorkbookConnection

    ' Refresh every pivot cache
    For Each pc In ThisWorkbook.PivotCaches
        On Error Resume Next
        pc.Refresh
        On Error GoTo 0
    Next pc

    ' Refresh every data connection
    For Each conn In ThisWorkbook.Connections
        On Error Resume Next
        conn.Refresh
        On Error GoTo 0
    Next conn

    ' Force full recalc
    Application.Calculate
End Sub

'==============================================================
' STEP 2  –  PREPARE DASHBOARD SHEET
'==============================================================
Private Sub PrepDashboardSheet()
    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' Remove old sheet if present
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets(DASH_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Add fresh sheet at the end
    Dim ws As Worksheet
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = DASH_SHEET

    ' Page setup
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PaperSize = xlPaperA4
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
    End With

    ' Set standard column widths and row heights for grid
    Dim i As Integer
    For i = 1 To 30
        ws.Columns(i).ColumnWidth = 8.5
    Next i
    For i = 1 To 80
        ws.Rows(i).RowHeight = 15
    Next i

    ' White background
    ws.Cells.Interior.Color = CLR_WHITE
End Sub

'==============================================================
' STEP 3  –  BUILD KPI ARRAY FROM PIVOT SHEET
' --- EDIT these cell references to match YOUR Pivot layout ---
'==============================================================
Private Sub BuildKPIArray(ByRef kpi() As Variant)
    Dim pv As Worksheet
    On Error Resume Next
    Set pv = ThisWorkbook.Sheets(PIVOT_SHEET)
    On Error GoTo 0

    ' Each KPI: Array(Label, Value, Prefix, NumberFormat)
    ' Defaults shown; replace cell refs to match your Pivot sheet
    Dim v As Variant

    ' KPI 1 – Total Revenue
    v = GetPivotValue(pv, "B2", 0) ' <-- change "B2" to your cell
    kpi(1) = Array("Total Revenue", v, "$", "#,##0")

    ' KPI 2 – Total Units Sold
    v = GetPivotValue(pv, "B3", 0)
    kpi(2) = Array("Units Sold", v, "", "#,##0")

    ' KPI 3 – Avg Order Value
    v = GetPivotValue(pv, "B4", 2)
    kpi(3) = Array("Avg Order Value", v, "$", "#,##0.00")

    ' KPI 4 – Growth %
    v = GetPivotValue(pv, "B5", 2)
    kpi(4) = Array("YoY Growth", v, "", "0.0%")
End Sub

Private Function GetPivotValue(ws As Worksheet, cellAddr As String, decimals As Integer) As Variant
    If ws Is Nothing Then
        GetPivotValue = 0
        Exit Function
    End If
    On Error Resume Next
    Dim v As Variant
    v = ws.Range(cellAddr).Value
    If IsEmpty(v) Or IsError(v) Then v = 0
    GetPivotValue = v
End Function

'==============================================================
' SECTION A  –  HEADER BANNER  (rows 1-6)
'==============================================================
Private Sub DrawHeader(ws As Worksheet)
    ' Background band
    With ws.Range("A1:AD6")
        .Merge
        .Interior.Color = CLR_DEEP
        .RowHeight = 15   ' applied to row 1 as proxy
    End With

    ' Adjust actual row heights
    Dim r As Integer
    For r = 1 To 6
        ws.Rows(r).RowHeight = 18
    Next r

    ' Title text
    With ws.Range("A1:AD6")
        .Value = "BUSINESS PERFORMANCE DASHBOARD"
        .Font.Name = "Segoe UI"
        .Font.Size = 26
        .Font.Bold = True
        .Font.Color = CLR_WHITE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Thin accent line below header
    With ws.Range("A7:AD7")
        .Interior.Color = CLR_ACCENT
        .RowHeight = 4
    End With

    ' Subtitle / timestamp row
    ws.Rows(8).RowHeight = 16
    With ws.Range("A8:AD8")
        .Merge
        .Value = "Last Refreshed: " & Format(Now(), "dd-mmm-yyyy  hh:mm AM/PM")
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = RGB(100, 0, 100)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
End Sub

'==============================================================
' SECTION B  –  KPI CARDS  (rows 10-20)
'==============================================================
Private Sub DrawKPICards(ws As Worksheet, kpi() As Variant)
    ' 4 cards across columns A-AD, evenly spaced
    Dim cardCols(1 To 4, 1 To 2) As Integer
    cardCols(1, 1) = 1: cardCols(1, 2) = 7      ' A-G
    cardCols(2, 1) = 9: cardCols(2, 2) = 15     ' I-O
    cardCols(3, 1) = 17: cardCols(3, 2) = 23    ' Q-W
    cardCols(4, 1) = 25: cardCols(4, 2) = 30    ' Y-AD

    Dim topRow As Integer: topRow = 10
    Dim botRow As Integer: botRow = 20

    ' Row heights for card area
    Dim r As Integer
    For r = topRow To botRow
        ws.Rows(r).RowHeight = IIf(r = topRow Or r = botRow, 8, 18)
    Next r

    Dim i As Integer
    For i = 1 To 4
        Dim c1 As Integer: c1 = cardCols(i, 1)
        Dim c2 As Integer: c2 = cardCols(i, 2)

        ' Card outer box
        Dim cardRng As Range
        Set cardRng = ws.Range(ws.Cells(topRow, c1), ws.Cells(botRow, c2))
        cardRng.Merge
        With cardRng
            .Interior.Color = CLR_LTGREY
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ' Coloured top stripe
        Dim stripeRng As Range
        Set stripeRng = ws.Range(ws.Cells(topRow, c1), ws.Cells(topRow + 1, c2))
        stripeRng.Merge
        stripeRng.Interior.Color = CLR_MID

        ' Label (row 12-13)
        Dim lblRng As Range
        Set lblRng = ws.Range(ws.Cells(topRow + 2, c1), ws.Cells(topRow + 3, c2))
        lblRng.Merge
        With lblRng
            .Value = UCase(CStr(kpi(i)(0)))
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = RGB(100, 0, 100)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = CLR_LTGREY
        End With

        ' Value (rows 14-17)
        Dim valRng As Range
        Set valRng = ws.Range(ws.Cells(topRow + 4, c1), ws.Cells(topRow + 7, c2))
        valRng.Merge
        Dim rawVal As Variant: rawVal = kpi(i)(1)
        Dim prefix  As String: prefix = CStr(kpi(i)(2))
        Dim fmt     As String: fmt = CStr(kpi(i)(3))
        Dim dispVal As String

        If fmt = "0.0%" Then
            dispVal = Format(CDbl(rawVal), "0.0") & "%"
        Else
            dispVal = prefix & Format(CDbl(rawVal), fmt)
        End If

        With valRng
            .Value = dispVal
            .Font.Name = "Segoe UI"
            .Font.Size = 18
            .Font.Bold = True
            .Font.Color = CLR_DEEP
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = CLR_LTGREY
        End With

        ' Border
        With cardRng.Borders
            .LineStyle = xlContinuous
            .Color = CLR_MID
            .Weight = xlThin
        End With
    Next i
End Sub

'==============================================================
' SECTION C  –  TOP CHARTS  (rows 22-50, two charts side-by-side)
'==============================================================
Private Sub DrawTopCharts(ws As Worksheet)
    ' ---- Left chart: Revenue by Category (Column chart) ----
    Call AddChart(ws, _
        chartType:=xlColumnClustered, _
        title:="Revenue by Category", _
        srcSheet:=DATA_SHEET, _
        srcRange:="A1:B20", _
        Left:=ws.Columns(1).Left + 6, _
        Top:=ws.Rows(22).Top + 4, _
        Width:=380, _
        Height:=220, _
        seriesColor:=CLR_MID)

    ' ---- Right chart: Monthly Trend (Line chart) ----
    Call AddChart(ws, _
        chartType:=xlLine, _
        title:="Monthly Sales Trend", _
        srcSheet:=DATA_SHEET, _
        srcRange:="A1:C20", _
        Left:=ws.Columns(16).Left + 6, _
        Top:=ws.Rows(22).Top + 4, _
        Width:=380, _
        Height:=220, _
        seriesColor:=CLR_ACCENT)
End Sub

'==============================================================
' SECTION D  –  BOTTOM CHARTS  (rows 54-78, three charts)
'==============================================================
Private Sub DrawBottomCharts(ws As Worksheet)
    ' ---- Left: Region breakdown (Pie) ----
    Call AddChart(ws, _
        chartType:=xlPie, _
        title:="Sales by Region", _
        srcSheet:=PIVOT_SHEET, _
        srcRange:="A1:B10", _
        Left:=ws.Columns(1).Left + 6, _
        Top:=ws.Rows(54).Top + 4, _
        Width:=240, _
        Height:=200, _
        seriesColor:=CLR_DEEP)

    ' ---- Centre: Top Products (Bar) ----
    Call AddChart(ws, _
        chartType:=xlBarClustered, _
        title:="Top Products", _
        srcSheet:=PIVOT_SHEET, _
        srcRange:="D1:E10", _
        Left:=ws.Columns(11).Left + 6, _
        Top:=ws.Rows(54).Top + 4, _
        Width:=240, _
        Height:=200, _
        seriesColor:=CLR_MID)

    ' ---- Right: Quarterly Comparison (Column) ----
    Call AddChart(ws, _
        chartType:=xlColumnClustered, _
        title:="Quarterly Comparison", _
        srcSheet:=PIVOT_SHEET, _
        srcRange:="G1:H10", _
        Left:=ws.Columns(21).Left + 6, _
        Top:=ws.Rows(54).Top + 4, _
        Width:=240, _
        Height:=200, _
        seriesColor:=CLR_ACCENT)
End Sub

'==============================================================
' GENERIC CHART BUILDER
'==============================================================
Private Sub AddChart(ws As Worksheet, _
                     chartType As XlChartType, _
                     title As String, _
                     srcSheet As String, _
                     srcRange As String, _
                     Left As Double, Top As Double, _
                     Width As Double, Height As Double, _
                     seriesColor As Long)

    ' Safely get source data range
    Dim srcWs As Worksheet
    On Error Resume Next
    Set srcWs = ThisWorkbook.Sheets(srcSheet)
    On Error GoTo 0
    If srcWs Is Nothing Then Exit Sub

    Dim dataRng As Range
    On Error Resume Next
    Set dataRng = srcWs.Range(srcRange)
    On Error GoTo 0
    If dataRng Is Nothing Then Exit Sub

    ' Add chart object
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=Left, Top:=Top, Width:=Width, Height:=Height)

    With co.Chart
        .chartType = chartType
        .SetSourceData Source:=dataRng, PlotBy:=xlColumns
        .HasTitle = True
        .chartTitle.Text = title
        .chartTitle.Font.Name = "Segoe UI"
        .chartTitle.Font.Size = 11
        .chartTitle.Font.Bold = True
        .chartTitle.Font.Color = CLR_DEEP

        ' Plot area / chart area colours
        .PlotArea.Interior.Color = RGB(252, 245, 255)
        .PlotArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color = CLR_WHITE
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.RoundedCorners = True

        ' Legend
        If .HasLegend Then
            .Legend.Font.Name = "Segoe UI"
            .Legend.Font.Size = 8
            .Legend.Font.Color = CLR_DKTEXT
        End If

        ' Series colour
        If .SeriesCollection.Count > 0 Then
            Dim s As Series
            Dim idx As Integer
            Dim tints As Variant
            tints = Array(0, 40, 80, 30, 60)  ' brightness offsets for multi-series

            For idx = 1 To .SeriesCollection.Count
                Set s = .SeriesCollection(idx)
                Dim baseR As Long, baseG As Long, baseB As Long
                baseR = (seriesColor Mod 256)
                baseG = ((seriesColor \ 256) Mod 256)
                baseB = ((seriesColor \ 65536) Mod 256)

                Dim t As Long
                t = tints((idx - 1) Mod 5)
                Dim adjR As Long, adjG As Long, adjB As Long
                adjR = WorksheetFunction.Min(255, baseR + t)
                adjG = WorksheetFunction.Min(255, baseG + t)
                adjB = WorksheetFunction.Min(255, baseB + t)

                s.Interior.Color = RGB(adjR, adjG, adjB)
                s.Format.Fill.ForeColor.RGB = RGB(adjR, adjG, adjB)

                If chartType = xlLine Or chartType = xlLineMarkers Then
                    s.Format.Line.ForeColor.RGB = seriesColor
                    s.Format.Line.Weight = 2.5
                    s.MarkerStyle = xlMarkerStyleCircle
                    s.MarkerSize = 6
                End If
            Next idx
        End If

        ' Axis formatting (where applicable)
        On Error Resume Next
        With .Axes(xlCategory)
            .TickLabels.Font.Name = "Segoe UI"
            .TickLabels.Font.Size = 8
            .TickLabels.Font.Color = CLR_DKTEXT
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(220, 200, 230)
        End With
        With .Axes(xlValue)
            .TickLabels.Font.Name = "Segoe UI"
            .TickLabels.Font.Size = 8
            .TickLabels.Font.Color = CLR_DKTEXT
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(220, 200, 230)
        End With
        On Error GoTo 0
    End With

    ' Drop-shadow effect on chart frame
    With co.ShapeRange
        .shadow.Visible = msoTrue
        .shadow.OffsetX = 3
        .shadow.OffsetY = 3
        .shadow.Blur = 6
        .shadow.ForeColor.RGB = RGB(180, 160, 200)
        .shadow.Transparency = 0.6
        .Line.ForeColor.RGB = CLR_LIGHT
        .Line.Weight = 1
    End With
End Sub

'==============================================================
' SECTION E  –  FOOTER  (rows 80-82)
'==============================================================
Private Sub DrawFooter(ws As Worksheet)
    Dim r As Integer
    For r = 80 To 82
        ws.Rows(r).RowHeight = 14
    Next r

    With ws.Range("A80:AD82")
        .Merge
        .Interior.Color = CLR_DEEP
        .Value = "CONFIDENTIAL  |  Auto-generated by Dashboard Macro  |  " & _
                 Format(Date, "MMMM YYYY")
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Color = CLR_LIGHT
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

'==============================================================
' UTILITY  –  Add Refresh Button to Dashboard sheet
' Call this ONCE manually to create the button
'==============================================================
Public Sub AddRefreshButton()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(DASH_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Run Refresh first to create Dashboard sheet, then call this again.", vbInformation
        Exit Sub
    End If

    ' Remove old button if present
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = "btnRefresh" Then shp.Delete
    Next shp

    ' Add new button
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                Left:=ws.Columns(1).Left + 10, _
                Top:=ws.Rows(9).Top + 2, _
                Width:=160, Height:=24)

    With shp
        .Name = "btnRefresh"
        .TextFrame.Characters.Text = "?  Refresh Dashboard"
        .TextFrame.Characters.Font.Name = "Segoe UI"
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = CLR_WHITE
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Fill.ForeColor.RGB = CLR_ACCENT
        .Fill.Solid
        .Line.ForeColor.RGB = CLR_DEEP
        .Line.Weight = 1
        .OnAction = "RefreshDashboard"
    End With
End Sub

