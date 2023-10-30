Sub functionality():
' declare variables for table
Dim ws As Worksheet
Dim PercentageChange As Double
Dim TotalVolume As Double
Dim SummaryRow As Integer
Dim Ticker As String
Dim Value As Double
Dim tickerRowNumber As String
Dim MinTickerRowNumber As String
Dim MaxTotalVolume As Double
' Set the initial values for summary table
    SummaryRow = 2
    TotalVolume = 0
   'set ticker and value headers for table
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    'label % increase, decrease and total volume
    Cells(SummaryRow, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
' greatest % increase is max value of column K
Cells(SummaryRow, 17).Value = Application.WorksheetFunction.Max(Range("k:k"))
' greatest % idecrease in max value of column K
Cells(3, 17).Value = Application.WorksheetFunction.Min(Range("k:k"))
'greatest total volume
Cells(4, 17).Value = Application.WorksheetFunction.Max(Range("L:L"))

'max value ticker symbol
tickerRowNumber = Application.WorksheetFunction.Match(Cells(SummaryRow, 17).Value, Range("k:k"), 0)
Cells(SummaryRow, 16).Value = Cells(tickerRowNumber, 9).Value

'min value ticker symbol
MinTickerRowNumber = Application.WorksheetFunction.Match(Cells(3, 17).Value, Range("k:k"), 0)
Cells(3, 16).Value = Cells(MinTickerRowNumber, 9).Value

' Greatest stock volume
MaxTotalVolume = Application.WorksheetFunction.Match(Cells(4, 17).Value, Range("L:L"), 0)
Cells(4, 16).Value = Cells(MaxTotalVolume, 9).Value
End Sub
