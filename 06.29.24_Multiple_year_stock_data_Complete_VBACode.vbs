Sub QuarterlyStockAnalysisAllSheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim ticker As String
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim startRow As Long
    Dim endRow As Long
    Dim currentQuarter As Integer
    Dim nextQuarter As Integer
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String
    
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        greatestIncrease = -1
        greatestDecrease = 1
        greatestVolume = 0
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Stock"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        j = 2
        startRow = 2
        
        For i = 2 To lastRow
            currentQuarter = WorksheetFunction.RoundUp(Month(ws.Cells(i, 2).Value) / 3, 0)
            If i = lastRow Then
                nextQuarter = 0
            Else
                nextQuarter = WorksheetFunction.RoundUp(Month(ws.Cells(i + 1, 2).Value) / 3, 0)
            End If
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or currentQuarter <> nextQuarter Then
                endRow = i
                ticker = ws.Cells(startRow, 1).Value
                openPrice = ws.Cells(startRow, 3).Value
                closePrice = ws.Cells(endRow, 6).Value
                totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7)))
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ws.Cells(j, 9).Value = ticker
                ws.Cells(j, 10).Value = quarterlyChange
                ws.Cells(j, 11).Value = percentChange / 100
                ws.Cells(j, 12).Value = totalVolume
                
                'greatest increase, decrease, and volume
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    tickerIncrease = ticker
                End If
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    tickerDecrease = ticker
                End If
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    tickerVolume = ticker
                End If
                
                j = j + 1
                startRow = i + 1
            End If
        Next i
        
        'Conditional formatting to the Quarterly Change
        With ws.Range("J2:J" & j - 1)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0)
        End With
        
        'Greatest values to the specified columns
        ws.Cells(2, 16).Value = tickerIncrease
        ws.Cells(3, 16).Value = tickerDecrease
        ws.Cells(4, 16).Value = tickerVolume
        ws.Cells(2, 17).Value = Format(greatestIncrease, "0.00") & "%"
        ws.Cells(3, 17).Value = Format(greatestDecrease, "0.00") & "%"
        ws.Cells(4, 17).Value = greatestVolume
        
    Next sheetName
    
    MsgBox "Analysis complete for all sheets!"
End Sub
