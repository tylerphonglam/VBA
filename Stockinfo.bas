Attribute VB_Name = "Module1"
Sub StockInfo()
    Dim ws As Worksheet
    Dim ticker As String
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearchange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summarylastrow As Integer
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxvolTicker As String
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxvol As Double
    
    For Each ws In ThisWorkbook.Sheets
        summarylastrow = 2
        totalVolume = 0
        maxPercentIncrease = 0
        maxPercentDecrease = 0
        maxvol = 0
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                openprice = ws.Cells(i, 3).Value
            End If
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                closeprice = ws.Cells(i, 6).Value
                yearchange = closeprice - openprice
                If openprice <> 0 Then
                    percentChange = yearchange / openprice
                Else
                    percentChange = 0
                End If
                ws.Cells(summarylastrow, 9).Value = ticker
                ws.Cells(summarylastrow, 10).Value = yearchange
                ws.Cells(summarylastrow, 11).Value = percentChange
                ws.Cells(summarylastrow, 12).Value = totalVolume
                ws.Cells(summarylastrow, 11).NumberFormat = "0.00%"
                If yearchange >= 0 Then
                    ws.Cells(summarylastrow, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(summarylastrow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                If percentChange > maxPercentIncrease Then
                    maxPercentIncrease = percentChange
                    maxPercentIncreaseTicker = ticker
                End If
                If percentChange < maxPercentDecrease Then
                    maxPercentDecrease = percentChange
                    maxPercentDecreaseTicker = ticker
                End If
                If totalVolume > maxvol Then
                    maxvol = totalVolume
                    maxvolTicker = ticker
                End If
                totalVolume = 0
                summarylastrow = summarylastrow + 1
            End If
        Next i
        ws.Columns("I:L").AutoFit
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 16).Value = maxPercentIncreaseTicker
        ws.Cells(2, 17).Value = maxPercentIncrease
        ws.Cells(3, 16).Value = maxPercentDecreaseTicker
        ws.Cells(3, 17).Value = maxPercentDecrease
        ws.Cells(4, 16).Value = maxvolTicker
        ws.Cells(4, 17).Value = maxvol
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
    Next ws
End Sub


