Attribute VB_Name = "Pivot"
Sub CreatePivotSheets()
    Dim wsData  As Worksheet
    Dim wsPvt   As Worksheet
    Dim pc      As PivotCache
    Dim pt      As PivotTable
    Dim lastRow As Long
    Dim sField  As String

    ' --- Validate Data sheet ---
    Set wsData = Nothing
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("Data")
    On Error GoTo 0
    If wsData Is Nothing Then
        MsgBox "Sheet named 'Data' not found!", vbCritical
        Exit Sub
    End If

    ' --- Get existing "Pivot" sheet ---
    Set wsPvt = Nothing
    On Error Resume Next
    Set wsPvt = ThisWorkbook.Sheets("Pivot")
    On Error GoTo 0
    If wsPvt Is Nothing Then
        MsgBox "Sheet named 'Pivot' not found!" & vbNewLine & _
               "Please create a sheet named 'Pivot' first.", vbCritical
        Exit Sub
    End If

    ' --- Clear all existing PivotTables and content (clean re-run) ---
    Dim existPT As PivotTable
    For Each existPT In wsPvt.PivotTables
        existPT.TableRange2.Clear
    Next existPT
    wsPvt.Cells.Clear

    ' --- Build shared PivotCache ---
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:="Data!$A$1:$M$" & lastRow)

    ' =========================================================
    ' PT_Segment â€" A1  (2 cols: Row Labels | Total Sales)
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("A2"), _
        TableName:="PT_Segment")
    With pt
        .PivotFields("Segment").Orientation = xlRowField
        .PivotFields("Segment").Position = 1
        With .PivotFields("Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
            .Name = "Total Sales"
        End With
    End With

    ' =========================================================
    ' PT_Trend â€" D1  (3 cols: Year | Month | Total Sales | Total Profit)
    ' 1-col gap (col C) separates from PT_Segment
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("D2"), _
        TableName:="PT_Trend")
    With pt
        .PivotFields("Year").Orientation = xlRowField
        .PivotFields("Year").Position = 1
        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Month").Position = 2
        With .PivotFields("Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
            .Name = "Total Sales"
        End With
        With .PivotFields("Profit")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
            .Name = "Total Profit"
        End With
    End With

    ' =========================================================
    ' PT_Monthly â€" H1  (cols: Month | 2023 | 2024 | 2025 | 2026...)
    ' 1-col gap (col G) separates from PT_Trend
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("H2"), _
        TableName:="PT_Monthly")
    With pt
        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Month").Position = 1
        .PivotFields("Year").Orientation = xlColumnField
        .PivotFields("Year").Position = 1
        With .PivotFields("Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
            .Name = "Total Sales"
        End With
    End With

    ' =========================================================
    ' PT_Category â€" N1  (2 cols: Row Labels | Total Sales)
    ' 1-col gap (col M) separates from PT_Monthly
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("O2"), _
        TableName:="PT_Category")
    With pt
        .PivotFields("Category").Orientation = xlRowField
        .PivotFields("Category").Position = 1
        With .PivotFields("Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
            .Name = "Total Sales"
        End With
    End With

    ' =========================================================
    ' PT_Region â€" Q1  (2 cols: Row Labels | Total Sales)
    ' 1-col gap (col P) separates from PT_Category
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("R2"), _
        TableName:="PT_Region")
    With pt
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Region").Position = 1
        With .PivotFields("Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
            .Name = "Total Sales"
        End With
    End With

    ' =========================================================
    ' PT_Discount â€" T1  (3 cols: Sub-Category | Total Profit | Avg Discount)
    ' 1-col gap (col S) separates from PT_Region
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("U2"), _
        TableName:="PT_Discount")
    With pt
        .PivotFields("Sub-Category").Orientation = xlRowField
        .PivotFields("Sub-Category").Position = 1
        With .PivotFields("Profit")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
            .Name = "Total Profit"
        End With
        With .PivotFields("Discount")
            .Orientation = xlDataField
            .Function = xlAverage
            .NumberFormat = "0.00%"
            .Name = "Avg Discount"
        End With
    End With

    ' =========================================================
    ' PT_Discount â€" T1  (3 cols: Sub-Category | Total Profit | Avg Discount)
    ' 1-col gap (col S) separates from PT_Region
    ' =========================================================


  Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("AB2"), _
        TableName:="PT_Quantity")
    With pt
        .PivotFields("Segment").Orientation = xlRowField
        .PivotFields("Segment").Position = 1
        With .PivotFields("Quantity")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
            .Name = "Total Quantity"
        End With
    End With




    ' =========================================================
    ' PT_Top5 â€" X1  (2 cols: Order ID | Total Sales)
    ' 1-col gap (col W) separates from PT_Discount
    ' FIX: AutoSort OUTSIDE With block, runtime field name
    ' =========================================================
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("Y2"), _
        TableName:="PT_Top5")
    With pt
        .PivotFields("Order ID").Orientation = xlRowField
        .PivotFields("Order ID").Position = 1
        With .PivotFields("Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
            .Name = "Total Sales"
        End With
    End With
    sField = pt.DataFields(1).Name
    pt.PivotFields("Order ID").AutoSort xlDescending, sField

    ' --- Write titles above each PivotTable in row 1 ---
    Call WritePivotTitles(wsPvt)

    ' --- Auto-fit all used columns ---
    wsPvt.Columns("A:Z").AutoFit

    ' --- Freeze top row so headers stay visible when scrolling ---
    wsPvt.Activate
    wsPvt.Range("A3").Select
    ActiveWindow.FreezePanes = True

End Sub


'-------------------------------------------------------------
'  HELPER - Write bold purple title above each PivotTable
'-------------------------------------------------------------
Sub WritePivotTitles(wsPvt As Worksheet)
    Dim purple  As Long: purple = RGB(94, 59, 183)
    Dim bgColor As Long: bgColor = RGB(235, 228, 255)
    Dim white   As Long: white = RGB(255, 255, 255)

    ' Define: title text, start col letter, end col letter
    Dim titles(7, 2) As String
    titles(0, 0) = "Sales by Segment": titles(0, 1) = "A": titles(0, 2) = "B"
    titles(1, 0) = "Sales & Profit Trend": titles(1, 1) = "D": titles(1, 2) = "F"
    titles(2, 0) = "Monthly Sales Performance": titles(2, 1) = "H": titles(2, 2) = "M"
    titles(3, 0) = "Sales by Category": titles(3, 1) = "O": titles(3, 2) = "P"
    titles(4, 0) = "Sales by Region": titles(4, 1) = "R": titles(4, 2) = "S"
    titles(5, 0) = "Profit vs Discount": titles(5, 1) = "U": titles(5, 2) = "W"
    titles(6, 0) = "Top 5 Sales Orders": titles(6, 1) = "Y": titles(6, 2) = "Z"
    titles(7, 0) = "Sales by Quantity": titles(7, 1) = "AB": titles(7, 2) = "AC"

    Dim i As Integer
    For i = 0 To 7
        Dim titleRange As String
        titleRange = titles(i, 1) & "1:" & titles(i, 2) & "1"

        ' Merge title cells across pivot width
        With wsPvt.Range(titleRange)
            .Merge
            .Value = titles(i, 0)
            .Font.Bold = True
            .Font.Size = 14
            .Font.Color = white
            .Font.Name = "Calibri"
            .Interior.Color = purple
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 20
        End With

        ' Bottom border under title to visually separate from pivot
        wsPvt.Range(titleRange).Borders(xlEdgeBottom).LineStyle = xlContinuous
        wsPvt.Range(titleRange).Borders(xlEdgeBottom).Color = bgColor
        wsPvt.Range(titleRange).Borders(xlEdgeBottom).Weight = xlThin
    Next i

    ' Set row 1 height
    wsPvt.Rows(1).RowHeight = 20
End Sub



