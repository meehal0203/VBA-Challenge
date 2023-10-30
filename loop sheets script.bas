Sub StockAnalysisLoop():

'Loop through all worksheets

For Each ws In ThisWorkbook.Worksheets


    Dim LastRow As Long
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim TotalVolume As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim SummaryRow As Integer
    Dim Start As Double
    
    'Set new column headers

    ws.Range("I1").Value = " Ticker"
    ws.Range("J1").Value = "Yearly_change"
    ws.Range("K1").Value = "Percent_price_change"
    ws.Range("L1").Value = "Total_stock_volume"

    ' Set the initial values for summary table
    SummaryRow = 2
    TotalVolume = 0
     Start = 2
    ' Find the last row with data in column A
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through the data and perform calculations
    For I = 2 To LastRow
    
    ' Add to the Total Stock Volume
            TotalVolume = TotalVolume + ws.Cells(I, 7).Value
            
        If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
            ' Set the ticker symbol
            Ticker = ws.Cells(I, 1).Value

            ' Calculate the Yearly Change
            OpenPrice = ws.Cells(Start, 3).Value
            ClosePrice = ws.Cells(I, 6).Value
            YearlyChange = ClosePrice - OpenPrice

            ' Calculate the Percentage Change
            If OpenPrice <> 0 Then
                PercentageChange = (YearlyChange / OpenPrice) * 100
            Else
                PercentageChange = 0
            End If
            Start = I + 1
            
              ' Output the results in the summary table
            ws.Cells(SummaryRow, 9).Value = Ticker
            ws.Cells(SummaryRow, 10).Value = YearlyChange
            ws.Cells(SummaryRow, 11).Value = PercentageChange
            ws.Cells(SummaryRow, 12).Value = TotalVolume
            
        
            
            'Conditional formating yearly change
             If ws.Cells(SummaryRow, 10).Value < 0 Then
                
                 'Set cell background color to red
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                
            Else
                
                 'Set cell background color to green
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
          End If
          
            'Conditional formating percentage change
             If ws.Cells(SummaryRow, 11).Value < 0 Then
                
                'Set cell background color to red
                ws.Cells(SummaryRow, 11).Interior.ColorIndex = 3
                
            Else
                
                    'Set cell background color to green
            ws.Cells(SummaryRow, 11).Interior.ColorIndex = 4
            
            End If
             
             ' Move to the next row in the summary table
            SummaryRow = SummaryRow + 1
            
            ' Add to the Total Stock Volume
            'TotalVolume = TotalVolume + ws.Cells(I, 7).Value

            ' Reset values for the next ticker
            TotalVolume = 0
            OpenPrice = ws.Cells(I + 1, 3).Value
        
        
           ' TotalVolume = TotalVolume + ws.Cells(I, 7).Value
            If OpenPrice = 0 Then
                OpenPrice = ws.Cells(I, 3).Value
            End If
        End If
    Next I
    
    Next ws
    
End Sub

