Attribute VB_Name = "Module1"
Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim startRow As Long

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' Add headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Add column headers
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        ' Initialize the start row
        startRow = 2

        ' Initialize worksheet specific variables
        greatestIncrease = -999999
        greatestDecrease = 999999
        greatestVolume = 0

        ' Loop through each ticker symbol
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = 0

                ' Loop through all rows for the same ticker
                For j = i To lastRow
                    If ws.Cells(j, 1).Value <> ticker Then
                        Exit For
                    End If
                    closePrice = ws.Cells(j, 6).Value
                    totalVolume = totalVolume + ws.Cells(j, 7).Value
                Next j

                ' Calculate quarterly change and percent change
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If

                ' Write values to the columns
                ws.Cells(startRow, 9).Value = ticker
                ws.Cells(startRow, 10).Value = quarterlyChange
                ws.Cells(startRow, 11).Value = percentChange / 100
                ws.Cells(startRow, 12).Value = totalVolume

                ' Apply percentage format to the percent change column
                ws.Cells(startRow, 11).NumberFormat = "0.00%"

                ' Apply conditional formatting
                If quarterlyChange > 0 Then
                    ws.Cells(startRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf quarterlyChange < 0 Then
                    ws.Cells(startRow, 10).Interior.Color = RGB(255, 0, 0)
                End If

                If percentChange > 0 Then
                    ws.Cells(startRow, 11).Interior.Color = RGB(0, 255, 0)
                ElseIf percentChange < 0 Then
                    ws.Cells(startRow, 11).Interior.Color = RGB(255, 0, 0)
                End If

                ' Update greatest values
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                End If

                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If

                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If

                startRow = startRow + 1
            End If
        Next i

        ' Write the greatest values to the cells in each worksheet
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncrease / 100
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease / 100
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolume

        ' Apply percentage format to the greatest increase and decrease cells
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"

    Next ws
End Sub


