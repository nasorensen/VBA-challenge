Attribute VB_Name = "Module1"
Sub ProcessMultipleSheets()

    Dim ws As Worksheet
    Dim i As Double
    Dim LastRow As Double
    Dim FuncLastRow As Double
    Dim Ticker As String
    Dim OPrice As Double
    Dim CPrice As Double
    Dim QChange As Double
    Dim Stock_Total As Double
    Dim StockSummary_Row As Integer
    Dim Count As Long
    Dim maxValue As Double
    Dim maxRow As String
    Dim minValue As Double
    Dim minRow As String
    Dim currentRow As Integer
    Dim columnToCheck As Integer
    Dim maxTotal As Double
    Dim maxTotalRow As String

    For Each ws In ThisWorkbook.Worksheets
        
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        Stock_Total = 0
        StockSummary_Row = 2
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
                ws.Range("I" & StockSummary_Row).Value = Ticker
                ws.Range("L" & StockSummary_Row).Value = Stock_Total
                StockSummary_Row = StockSummary_Row + 1
                Stock_Total = 0
            Else
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            End If
        Next i
        
        Count = 0
        OPrice = 0
        CPrice = 0
        StockSummary_Row = 2
        
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Count = Count + 1
                OPrice = OPrice + ws.Cells(i, 3).Value
                CPrice = CPrice + ws.Cells(i, 6).Value
                QChange = CPrice - OPrice
                ws.Cells(StockSummary_Row, 10).Value = QChange
                
                If OPrice <> 0 Then
                    ws.Cells(StockSummary_Row, 11).Value = (QChange / OPrice) * 100
                Else
                    ws.Cells(StockSummary_Row, 11).Value = 0
                End If
                
            StockSummary_Row = StockSummary_Row + 1
            OPrice = 0
            CPrice = 0
            Count = 0
            
            Else
                Count = Count + 1
                OPrice = OPrice + ws.Cells(i, 3).Value
                CPrice = CPrice + ws.Cells(i, 6).Value
            
            End If
                
        Next i
        
        For i = 2 To LastRow
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 0
            End If
        Next i
        
        For i = 2 To LastRow
            ws.Cells(i, 11).Style = "Percent"
        Next i
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        maxRow = 1
        For i = 2 To LastRow
            If ws.Cells(i, 11).Value > maxValue Then
            maxValue = ws.Cells(i, 11).Value
            maxRow = ws.Cells(i, 1).Value

            End If
        Next i
        
        For i = 2 To LastRow
            If ws.Cells(i, 11).Value < minValue Then
                minValue = ws.Cells(i, 11).Value
                minRow = ws.Cells(i, 1).Value

            End If
        Next i
        
            For i = 2 To LastRow
                If ws.Cells(i, 12).Value > maxTotal Then
                    maxTotal = ws.Cells(i, 12).Value
                    maxTotalRow = ws.Cells(i, 1).Value

                End If
        Next i
        
        ws.Range("P2").Value = maxRow
        ws.Range("Q2").Value = maxValue
        ws.Range("Q2").Style = "Percent"
        ws.Range("P3").Value = minRow
        ws.Range("Q3").Value = minValue
        ws.Range("Q3").Style = "Percent"
        ws.Range("P4").Value = maxTotalRow
        ws.Range("Q4").Value = maxTotal
        
    Next ws

End Sub

