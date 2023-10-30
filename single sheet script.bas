Sub functionalityLoop():

'Loop through all worksheets

For Each ws In ThisWorkbook.Worksheets


' declare variables for table

Dim PercentageChange As Double
Dim TotalVolume As Double
Dim SummaryRow As Integer
Dim Ticker As String
Dim Value As Double
' Set the initial values for summary table
    SummaryRow = 2
    TotalVolume = 0
   'set ticker and value headers for table
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    'label % increase, decrease and total volume
    ws.Cells(SummaryRow, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
' greatest % increase is max value of column K
ws.Cells(SummaryRow, 17).Value = Application.WorksheetFunction.Max(ws.Range("k:k"))
' greatest % idecrease in max value of column K
ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range("k:k"))
'greatest total volume
ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L:L"))

'max value ticker symbol
tickerRowNumber = Application.WorksheetFunction.Match(ws.Cells(SummaryRow, 17).Value, ws.Range("k:k"), 0)
ws.Cells(SummaryRow, 16).Value = ws.Cells(tickerRowNumber, 9).Value

'min value ticker symbol
MinTickerRowNumber = Application.WorksheetFunction.Match(ws.Cells(3, 17).Value, ws.Range("k:k"), 0)
ws.Cells(3, 16).Value = ws.Cells(MinTickerRowNumber, 9).Value

' Greatest stock volume
MaxTotalVolume = Application.WorksheetFunction.Match(ws.Cells(4, 17).Value, ws.Range("L:L"), 0)
ws.Cells(4, 16).Value = ws.Cells(MaxTotalVolume, 9).Value


Next ws

End Sub

