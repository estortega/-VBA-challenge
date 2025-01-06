Attribute VB_Name = "Module1"
Sub MultipleYearStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim summaryRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
    ' Table headers
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Quarterly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
        
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    summaryRow = 2
    
        
    Dim startRow As Long
    startRow = 2
        
    ' Variables to track greatest values
    greatestIncrease = -1
    greatestDecrease = 1
    greatestVolume = 0
    greatestIncreaseTicker = ""
    greatestDecreaseTicker = ""
    greatestVolumeTicker = ""
        
    ' Loop through each ticker
    Do While startRow <= lastRow
    ticker = ws.Cells(startRow, 1).Value
    openPrice = ws.Cells(startRow, 3).Value
    totalVolume = 0
            
    Dim currentRow As Long
    currentRow = startRow
            
    ' Loop through the rows of the same ticker
    Do While currentRow <= lastRow And ws.Cells(currentRow, 1).Value = ticker
    totalVolume = totalVolume + ws.Cells(currentRow, 7).Value
    currentRow = currentRow + 1
    Loop
            
    ' Closing price at the last row of the current ticker
    closePrice = ws.Cells(currentRow - 1, 6).Value
            
    ' Quarterly change and percentage change
    quarterlyChange = closePrice - openPrice
    If openPrice <> 0 Then
    percentageChange = (quarterlyChange / openPrice) * 100
    Else
    percentageChange = 0
    End If
            
    ' Output the results in the summary table
    ws.Cells(summaryRow, 10).Value = ticker
    ws.Cells(summaryRow, 11).Value = quarterlyChange
    ws.Cells(summaryRow, 12).Value = percentageChange
    ws.Cells(summaryRow, 13).Value = totalVolume
            
            
    ' Check for greatest % increase, % decrease, and total volume
    If percentageChange > greatestIncrease Then
    greatestIncrease = percentageChange
    greatestIncreaseTicker = ticker
    End If
    If percentageChange < greatestDecrease Then
    greatestDecrease = percentageChange
    greatestDecreaseTicker = ticker
    End If
    If totalVolume > greatestVolume Then
    greatestVolume = totalVolume
    greatestVolumeTicker = ticker
    End If
            
    summaryRow = summaryRow + 1
    startRow = currentRow
    Loop
        
    ' Output the metric results
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = greatestIncreaseTicker
    ws.Cells(2, 17).Value = greatestIncrease
        
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = greatestDecreaseTicker
    ws.Cells(3, 17).Value = greatestDecrease
        
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = greatestVolumeTicker
    ws.Cells(4, 17).Value = greatestVolume
            
    ' Apply conditional formatting for Quarterly Change
    Dim rng As Range
    lastRow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
    Set rng = ws.Range("K2:K" & lastRow)
        
    ' Positive change in green
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
    .Interior.Color = RGB(144, 238, 144)
    End With
        
    ' Negative change in red
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
    .Interior.Color = RGB(255, 99, 71)
    End With
        
    Next ws
    
End Sub




