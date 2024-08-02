# vba-challenge
Sub stock_loop()

  #Loops through each worksheet
    For Each ws In Worksheets

    #Naming and assigning variables
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
    
    #Populating summary table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
    #setting start values to variables
        tickervol = 0
        tickersymbol_row = 2
        open_value = Cells(2, 3).Value
        change = 0
        j = 0

    #finding last row in worksheet
        row_count = ws.Cells(Rows.Count, "A").End(xlUp).Row
    #looping through tickers
        For i = 2 To row_count
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                tickervol = tickervol + ws.Cells(i, 7).Value
            
                tickersymbol = ws.Cells(i, 1).Value
            
                ws.Range("I" & tickersymbol_row).Value = tickersymbol
            
                ws.Range("L" & tickersymbol_row).Value = tickervol
            
                close_value = ws.Cells(i, 6).Value
            
                quarterly_change = (close_value - open_value)
            
                ws.Range("J" & tickersymbol_row).Value = quarterly_change
                
                If open_value = 0 Then
                    
                    percent_change = 0
                    
                Else
                    
                    percent_change = quarterly_change / open_value
                    
                End If
                
                ws.Range("K" & tickersymbol_row).Value = percent_change
                ws.Range("K" & tickersymbol_row).NumberFormat = "0.00%"
            
                tickersymbol_row = tickersymbol_row + 1
            
                tickervol = 0
            
                open_value = ws.Cells(i + 1, 3)
            
            Else
        
                tickervol = tickervol + ws.Cells(i, 7).Value
            
            End If
        
        Next i
    
        lastrow_summarytbl = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        For i = 2 To lastrow_summarytbl
    
            If ws.Cells(i, 10).Value > 0 Then
            
                ws.Cells(i, 10).Interior.ColorIndex = 4
            
            Else
        
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i
    
    
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
