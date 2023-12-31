Attribute VB_Name = "Module1"
Sub AnalyzeStockData()
    ' Loop through each worksheet in the workbook
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Set up variables
        Dim LastRow As Long
        Dim Ticker As String
        Dim Volume As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As Double
        Dim TickerGreatestIncrease As String
        Dim TickerGreatestDecrease As String
        Dim TickerGreatestVolume As String
        
        ' Find the last row of data in the worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Add column headers for the analyzed data
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        
        ' Loop through each row of data
        Dim i As Long
        Dim SummaryRow As Long
        SummaryRow = 2 ' Start from row 2 to write data
        TotalVolume = 0
        OpenPrice = ws.Cells(2, 3).Value ' Store the initial open price
        
        For i = 2 To LastRow
            ' Check if the ticker symbol has changed
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Store the ticker symbol
                Ticker = ws.Cells(i, 1).Value
                
                ' Store the close price
                ClosePrice = ws.Cells(i, 6).Value
                
                ' Calculate the yearly change
                YearlyChange = ClosePrice - OpenPrice
                
                ' Calculate the percent change
                If OpenPrice <> 0 Then
                    PercentChange = YearlyChange / OpenPrice
                Else
                    PercentChange = 0
                End If
                
                ' Write the data to the summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = TotalVolume
                ws.Cells(SummaryRow, 11).Value = YearlyChange
                ws.Cells(SummaryRow, 12).Value = PercentChange
                
                ' Apply conditional formatting to yearly change column
                If YearlyChange > 0 Then
                    ws.Cells(SummaryRow, 11).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(SummaryRow, 11).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                
                ' Add to the summary table row
                SummaryRow = SummaryRow + 1
                
                ' Reset variables for the next ticker
                TotalVolume = 0
                OpenPrice = ws.Cells(i + 1, 3).Value ' Update the open price for the next ticker
            End If
            
            ' Add to the total volume
            Volume = ws.Cells(i, 7).Value
            TotalVolume = TotalVolume + Volume
            
            ' Check if it's the last row of data for the worksheet
            If i = LastRow Then
                ' Store the close price for the last ticker
                ClosePrice = ws.Cells(i, 6).Value
                
                ' Calculate the yearly change and percent change for the last ticker
                YearlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = YearlyChange / OpenPrice
                Else
                    PercentChange = 0
                End If
                
                ' Write the data to the summary table for the last ticker
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = TotalVolume
                ws.Cells(SummaryRow, 11).Value = YearlyChange
                ws.Cells(SummaryRow, 12).Value = PercentChange
                
                ' Apply conditional formatting to yearly change column for the last ticker
                If YearlyChange > 0 Then
                    ws.Cells(SummaryRow, 11).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(SummaryRow, 11).Interior.Color = RGB(255, 0, 0) ' Red
                End If
            End If
            
            ' Check if the percent change is the greatest increase or decrease
            If PercentChange > GreatestIncrease Then
                GreatestIncrease = PercentChange
                TickerGreatestIncrease = Ticker
            ElseIf PercentChange < GreatestDecrease Then
                GreatestDecrease = PercentChange
                TickerGreatestDecrease = Ticker
            End If
            
            ' Check if the total volume is the greatest volume
            If TotalVolume > GreatestVolume Then
                GreatestVolume = TotalVolume
                TickerGreatestVolume = Ticker
            End If
        Next i
        
        ' Write the calculated values to the summary table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = TickerGreatestIncrease
        ws.Cells(2, 17).Value = Format(GreatestIncrease, "0.00%") ' Display as percentage
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = TickerGreatestDecrease
        ws.Cells(3, 17).Value = Format(GreatestDecrease, "0.00%") ' Display as percentage
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = TickerGreatestVolume
        ws.Cells(4, 17).Value = GreatestVolume
        
        ' Apply conditional formatting to percent change column
        ws.Columns(12).NumberFormat = "0.00%" ' Format as percentage
        ws.Columns(12).FormatConditions.AddColorScale ColorScaleType:=3
        ws.Columns(12).FormatConditions(ws.Columns(12).FormatConditions.Count).SetFirstPriority
        ws.Columns(12).FormatConditions(1).ColorScaleCriteria(1).Type = _
            xlConditionValuePercentile
        ws.Columns(12).FormatConditions(1).ColorScaleCriteria(1).Value = 0
        ws.Columns(12).FormatConditions(1).ColorScaleCriteria(2).Type = _
            xlConditionValuePercentile
        ws.Columns(12).FormatConditions(1).ColorScaleCriteria(2).Value = 50
        ws.Columns(12).FormatConditions(1).ColorScaleCriteria(3).Type = _
            xlConditionValuePercentile
        ws.Columns(12).FormatConditions(1).ColorScaleCriteria(3).Value = 100
        ws.Columns(12).FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0) ' Red
        ws.Columns(12).FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = RGB(0, 255, 0) ' Green
        ws.Columns(12).FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = RGB(0, 255, 0) ' Green

        ' Auto-fit columns for better visibility
        ws.Columns.AutoFit
    Next ws
End Sub


