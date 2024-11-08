# VBA_challenge
Sub QuarterlyStockAnalysis()

    Dim ws As Worksheet
    Dim lastRow As Long, summaryRow As Long
    Dim ticker As String, currentTicker As String
    Dim openPrice As Double, closePrice As Double
    Dim quarterlyChange As Double, percentChange As Double
    Dim totalVolume As Double
    
    Dim maxIncrease As Double, maxDecrease As Double, maxVolume As Double
    Dim maxIncreaseTicker As String, maxDecreaseTicker As String, maxVolumeTicker As String

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables
        summaryRow = 2
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        maxIncreaseTicker = ""
        maxDecreaseTicker = ""
        maxVolumeTicker = ""
        
        ' Find the last row in the data
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through all rows to calculate quarterly data
        For i = 2 To lastRow
            ' Check if we are still on the same ticker
            If ws.Cells(i, 1).Value <> currentTicker Then
                If currentTicker <> "" Then
                    ' Calculate quarterly change and percent change
                    quarterlyChange = closePrice - openPrice
                    If openPrice <> 0 Then
                        percentChange = (quarterlyChange / openPrice) * 100
                    Else
                        percentChange = 0
                    End If
                    
                    ' Output to summary table
                    ws.Cells(summaryRow, 8).Value = currentTicker
                    ws.Cells(summaryRow, 9).Value = quarterlyChange
                    ws.Cells(summaryRow, 10).Value = percentChange
                    ws.Cells(summaryRow, 11).Value = totalVolume
                    
                    ' Update max values
                    If percentChange > maxIncrease Then
                        maxIncrease = percentChange
                        maxIncreaseTicker = currentTicker
                    End If
                    If percentChange < maxDecrease Then
                        maxDecrease = percentChange
                        maxDecreaseTicker = currentTicker
                    End If
                    If totalVolume > maxVolume Then
                        maxVolume = totalVolume
                        maxVolumeTicker = currentTicker
                    End If
                    
                    ' Move to next summary row
                    summaryRow = summaryRow + 1
                End If
                
                ' New ticker - reset values
                currentTicker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = 0
            End If
            
            ' Update close price and accumulate volume
            closePrice = ws.Cells(i, 6).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        Next i
        
        ' Last ticker summary
        quarterlyChange = closePrice - openPrice
        If openPrice <> 0 Then
            percentChange = (quarterlyChange / openPrice) * 100
        Else
            percentChange = 0
        End If

        ' Output the last ticker's summary data
        ws.Cells(summaryRow, 8).Value = currentTicker
        ws.Cells(summaryRow, 9).Value = quarterlyChange
        ws.Cells(summaryRow, 10).Value = percentChange
        ws.Cells(summaryRow, 11).Value = totalVolume

        ' Update max values for last ticker
        If percentChange > maxIncrease Then
            maxIncrease = percentChange
            maxIncreaseTicker = currentTicker
        End If
        If percentChange < maxDecrease Then
            maxDecrease = percentChange
            maxDecreaseTicker = currentTicker
        End If
        If totalVolume > maxVolume Then
            maxVolume = totalVolume
            maxVolumeTicker = currentTicker
        End If

        ' Output the greatest increase, decrease, and total volume
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = maxIncreaseTicker
        ws.Cells(2, 16).Value = maxIncrease

        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = maxDecreaseTicker
        ws.Cells(3, 16).Value = maxDecrease

        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = maxVolumeTicker
        ws.Cells(4, 16).Value = maxVolume

        ' Formatting the output
        ws.Columns("I:K").NumberFormat = "0.00"
        ws.Columns("L:L").NumberFormat = "0"
        ws.Columns("P:P").NumberFormat = "0.00"
        
        ' Adding conditional formatting for positive and negative values in Quarterly Change
        With ws.Range("I2:I" & summaryRow)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(0, 255, 0) ' Green for positive
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0) ' Red for negative
        End With

        ' Format headers
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Quarterly Change"
        ws.Cells(1, 10).Value = "Percent Change"
        ws.Cells(1, 11).Value = "Total Stock Volume"

    Next ws

    MsgBox "Quarterly stock analysis completed on all sheets!"

End Sub

