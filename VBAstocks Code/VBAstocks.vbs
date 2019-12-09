Sub VBA_Stocks():
    Dim WorksheetName As String
    Dim brand As String
    Dim total As Double
    Dim summaryRow As Integer
    Dim firstOpen As Double
    Dim finalClose As Double
    Dim yearChange As Double
    Dim percentChange As Long
    Dim lastRow As Long
    Dim firstRow As Long

    For Each ws In Worksheets
        total = 0
        summaryRow = 2
        yearChange = finalClose - firstOpen
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Volume"
        
        For I = 2 To lastRow
            If firstOpen = 0 Then
                firstOpen = Cells(i, 3).Value
            End If
            If (ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value) Then
                brand = ws.Cells(I, 1).Value
                total = total + ws.Cells(I, 7).Value
                finalClose = ws.Cells(I, 6).Value
                ws.Cells(summaryRow, 9).Value = brand
                ws.Cells(summaryRow, 10).Value = finalClose - firstOpen
                ws.Cells(summaryRow, 11).Value = ((finalClose - firstOpen) / firstOpen)
                ws.Cells(summaryRow, 12).Value = total
                summaryRow = summaryRow + 1
                total = 0
            Else
                total = total + ws.Cells(I, 7).Value
            End If
        Next I
        lastSummaryRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        For I = 2 To lastSummaryRow
            ws.Cells(I, 11).Style = "Percent"
            If ws.Cells(I, 10).Value <= 0 Then
                ws.Cells(I, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(I, 10).Interior.ColorIndex = 4
            End If
            
            If ws.Cells(I, 11).Value <= 0 Then
                ws.Cells(I, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(I, 11).Interior.ColorIndex = 4
            End If
        Next I
    Next ws
End Sub