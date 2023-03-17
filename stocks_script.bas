Attribute VB_Name = "Module1"
Sub stocks()
For Each ws In Worksheets

    Dim counter As Long
    Dim lastrow As Long
    Dim ticker As String
    Dim openvalue As Double
    Dim closevalue As Double
    Dim total_vol As Double
    ' Dim great_increase As Double
    ' Dim great_decrease As Double
    ' Dim great_total As Double
    Dim percent_change As Double
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' total stock volume starts at zero
    
    total_vol = 0
    
    ' counter starts at top of table
    
    counter = 1
    
    ' set greatest increase, decrease, and total volume
    
    
    For i = 2 To lastrow
    
        ' add volume to total
        
        total_vol = total_vol + ws.Range("G" & i).Value
    
        ' set ticker
        
        ticker = ws.Range("A" & i).Value
        
            
            
        ' ticker is the same as the next one (not the last)
        
        If ticker = ws.Range("A" & i + 1).Value Then
            
            ' ticker is the same as the previous (not the first)
            
            If ticker = ws.Range("A" & i - 1).Value Then
                
            ' ticker is not the same as the previous (new ticker, first row of stock)
            
            Else
            
                ' counter goes up, changing new table position
            
                counter = counter + 1
            
                ' ticker symbol is set to new stock and added to the new table
                        
                ticker = ws.Range("A" & i).Value
                  
                ws.Range("I" & counter) = ticker
                
                ' sets new open value
            
                openvalue = ws.Range("C" & i).Value
                    
            End If
            
        ' last row for this stock
        
        Else
        
            ' sets close value as the value for the last row of each stock
            
            closevalue = ws.Range("F" & i).Value
            
            'finds out change in stock price
            
            ws.Range("J" & counter).Value = closevalue - openvalue
            
            ' determines if value is + or - then assigns green or red fill to cell, respectively
            
            If (closevalue - openvalue) > 0 Then
                ws.Range("J" & counter).Interior.ColorIndex = 4
            
            ElseIf (closevalue - openvalue) < 0 Then
                ws.Range("J" & counter).Interior.ColorIndex = 3
            
            End If
            
            ' now for percentage change
            
            ws.Range("K" & counter).Value = FormatPercent((closevalue - openvalue) / openvalue)
            
            percent_change = ws.Range("K" & counter).Value
        
            ' put total stock volume into new table, then reset to zero
            
            ws.Range("L" & counter).Value = total_vol
                
            ' now for greatest increase, decrease, and total volume
            
            If percent_change > ws.Range("P2").Value Then
                ws.Range("P2").Value = FormatPercent(percent_change)
                ws.Range("O2").Value = ticker
                
            End If
            
            If percent_change < ws.Range("P3").Value Then
                ws.Range("P3").Value = FormatPercent(percent_change)
                ws.Range("O3").Value = ticker
                
            End If
            
            If total_vol > ws.Range("P4") Then
                ws.Range("P4") = total_vol
                ws.Range("O4").Value = ticker
                
            End If
            
            ' reset total volume to zero
            
            total_vol = 0
        
        End If
    
    Next i
        
Next ws
        
End Sub
    

