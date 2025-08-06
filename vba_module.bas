Option Explicit

' Module: SalesReportAutomation
' Description: Generates a summary report (TotalQuantity, TotalSales) by Product,
' creates a chart, formats the Report sheet, and exports it as a PDF.

Sub GenerateSalesReport()
    Dim wb As Workbook
    Dim wsData As Worksheet, wsReport As Worksheet
    Dim lastRow As Long, i As Long, r As Long
    Dim prod As String
    Dim qty As Double, price As Double, sales As Double
    Dim dictQty As Object, dictSales As Object
    Dim key As Variant

    Set wb = ThisWorkbook
    On Error Resume Next
    Set wsData = wb.Sheets("RawData")
    If wsData Is Nothing Then
        MsgBox "RawData sheet not found. Please ensure the sheet is named 'RawData'.", vbCritical
        Exit Sub
    End If
    Set wsReport = wb.Sheets("Report")
    If wsReport Is Nothing Then
        Set wsReport = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsReport.Name = "Report"
    End If
    wsReport.Cells.Clear

    Set dictQty = CreateObject("Scripting.Dictionary")
    Set dictSales = CreateObject("Scripting.Dictionary")

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        prod = Trim(wsData.Cells(i, 3).Value)
        qty = Val(wsData.Cells(i, 4).Value)
        price = Val(wsData.Cells(i, 5).Value)
        sales = qty * price
        If prod <> "" Then
            If dictQty.Exists(prod) Then
                dictQty(prod) = dictQty(prod) + qty
                dictSales(prod) = dictSales(prod) + sales
            Else
                dictQty.Add prod, qty
                dictSales.Add prod, sales
            End If
        End If
    Next i

    ' Write header
    wsReport.Range("A1").Value = "Product"
    wsReport.Range("B1").Value = "TotalQuantity"
    wsReport.Range("C1").Value = "TotalSales"
    r = 2
    For Each key In dictQty.Keys
        wsReport.Cells(r, 1).Value = key
        wsReport.Cells(r, 2).Value = dictQty(key)
        wsReport.Cells(r, 3).Value = Round(dictSales(key), 2)
        r = r + 1
    Next key

    ' Sort by TotalSales descending
    With wsReport.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsReport.Range("C2:C" & r - 1), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange wsReport.Range("A1:C" & r - 1)
        .Header = xlYes
        .Apply
    End With

    ' Format
    wsReport.Columns("A:C").AutoFit
    wsReport.Range("C2:C" & r - 1).NumberFormat = "#,##0.00"

    ' Delete existing charts
    On Error Resume Next
    Dim chObj As ChartObject
    For Each chObj In wsReport.ChartObjects
        chObj.Delete
    Next chObj
    On Error GoTo 0

    ' Create chart
    Dim chtObj As ChartObject
    Set chtObj = wsReport.ChartObjects.Add(Left:=300, Top:=10, Width:=480, Height:=300)
    chtObj.Chart.ChartType = xlColumnClustered
    chtObj.Chart.SetSourceData Source:=wsReport.Range("A1:C" & r - 1)
    chtObj.Chart.HasTitle = True
    chtObj.Chart.ChartTitle.Text = "Total Sales by Product"

    ' Export as PDF if workbook has a valid path
    Dim pdfPath As String
    If wb.Path <> "" Then
        pdfPath = wb.Path & Application.PathSeparator & "Sales_Report.pdf"
        wsReport.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End If

    MsgBox "Report generated successfully." & IIf(wb.Path <> "", vbCrLf & "PDF saved to: " & pdfPath, ""), vbInformation
End Sub