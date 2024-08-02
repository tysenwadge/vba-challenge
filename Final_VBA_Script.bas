Attribute VB_Name = "Module1"
Sub stock_loop()

 'loop through worksheets
    For Each ws In Worksheets
    
    'setting variables
        Dim tickersymbol As String
        Dim tickersymbol_row As Integer
        Dim tickervol As Double
        Dim open_value As Double
        Dim close_value As Double
        Dim percent_change As Double
        Dim i As Long
        Dim j As Integer
        Dim row_count As Long
        Dim daily_change As Double
        Dim quarterly_change As Double
        Dim change As Double
    
    'Populating cell header values for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
    'Setting start values to variables
        tickervol = 0
        tickersymbol_row = 2
        open_value = Cells(2, 3).Value
        change = 0
        j = 0
    
    'Find the last row in the worksheet
        row_count = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Loop to go through the tickers
        For i = 2 To row_count
        'Compares when value of next ticker cell is different from current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Add the ticker volume
                tickervol = tickervol + ws.Cells(i, 7).Value
           'Set ticker name
                tickersymbol = ws.Cells(i, 1).Value
            'Prints the ticker name to the summary table
                ws.Range("I" & tickersymbol_row).Value = tickersymbol
            'Prints the stock volume for each ticker in the summary table
                ws.Range("L" & tickersymbol_row).Value = tickervol
            'Stores the close price
                close_value = ws.Cells(i, 6).Value
            'Calculates the change across the quarter
                quarterly_change = (close_value - open_value)
            'Prints the quarterly change to the summary table
                ws.Range("J" & tickersymbol_row).Value = quarterly_change
             'Checks to make sure we don't divide by zero
                If open_value = 0 Then
                    
                    percent_change = 0
                    
                Else
                    
                    percent_change = quarterly_change / open_value
                    
                End If
              'Prints and formats percentage change in summary table
                ws.Range("K" & tickersymbol_row).Value = percent_change
                ws.Range("K" & tickersymbol_row).NumberFormat = "0.00%"
                
            'Resets the row counter
                tickersymbol_row = tickersymbol_row + 1
                
            'Resets stock trade volume
                tickervol = 0
                
            'setting initial stock open price
                open_value = ws.Cells(i + 1, 3)
            
            Else
        
                tickervol = tickervol + ws.Cells(i, 7).Value
            
            End If
        
        Next i
    
        lastrow_summarytbl = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
    'Conditional formatting to highlight negative change in red and positive in green.
        For i = 2 To lastrow_summarytbl
    
            If ws.Cells(i, 10).Value > 0 Then
            
                ws.Cells(i, 10).Interior.ColorIndex = 4
            
            Else
        
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i
    
    'Loop to determine min and max percent change, and largest total stock volume
        For i = 2 To lastrow_summarytbl
    
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summarytbl)) Then
            
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summarytbl)) Then
        
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summarytbl)) Then
        
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
        
    Next ws
    
    
End Sub
