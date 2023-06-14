Sub stockmarket()

Dim Summary_Table_Row As Integer
Dim Stockvol_total As Double
Dim Ticker As String
Dim LR As Long
Dim ws As Worksheet



For Each ws In Worksheets

    Summary_Table_Row = 2
    Stockvol_total = 0

    'Determine the last row
    LR = Cells(Rows.Count, "A").End(xlUp).Row

    'Add 4 columns and their respective labels & values to each sheet

    'Ticker column
    ws.Range("J1").EntireColumn.Insert
    ws.Cells(1, 10).Value = "Ticker"
    
     'Yearly change column
    ws.Range("K1").EntireColumn.Insert
    ws.Cells(1, 11).Value = "Yearly Change"
    
    'Percentage change column
    ws.Range("L1").EntireColumn.Insert
    ws.Cells(1, 12).Value = "Percentage Change"
    
    'Total volume stock column
    ws.Range("M1").EntireColumn.Insert
    ws.Cells(1, 13).Value = "Total Stock Volume"
    
    'Greatest % increase/decrease table
    ws.Range("Q1").EntireColumn.Insert
    ws.Cells(1, 17).Value = "Ticker"
    
    ws.Range("R1").EntireColumn.Insert
    ws.Cells(1, 18).Value = "Value"
    
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    
    ws.Range("J1:M1").EntireColumn.AutoFit
    
    For i = 2 To LR
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         ' Set stock name
            Ticker = ws.Cells(i, 1).Value
            'Set Close price
            Close_price = ws.Cells(i, 6).Value
            'Set Open_price
            Open_price = ws.Cells(i - 250, 3).Value
            
            Yearly_change = Close_price - Open_price
            Perc_change = Yearly_change / Open_price
          
          'Add to stock volume total
            Stockvol_total = Stockvol_total + ws.Cells(i, 7)
           
           'Print stock name in the ticker column
          
            ws.Range("J" & Summary_Table_Row).Value = Ticker
            
            'Print calculated Yearly change and Percent change
            ws.Range("K" & Summary_Table_Row).Value = Yearly_change
            ws.Range("L" & Summary_Table_Row).Value = FormatPercent(Perc_change)
            
            'Print total stock volume to Total Stock Volumn column
            ws.Range("M" & Summary_Table_Row).Value = Stockvol_total
          
             'Add one to summary table
            Summary_Table_Row = Summary_Table_Row + 1
            
            Stockvol_total = 0
          
         Else
         
         ' add to the Stock vol total
         Stockvol_total = Stockvol_total + ws.Cells(i, 7).Value
        
        End If
        
    Next i
    
    'Conditional formatting
    
    For r = 2 To 3001
        For c = 11 To 12
        
            If ws.Cells(r, c).Value >= 0 Then
                ws.Cells(r, c).Interior.ColorIndex = 4 'Green
            Else
                
                ws.Cells(r, c).Interior.ColorIndex = 3 'Red
                
            End If
        Next c
    Next r
    
    'Greatest % increae, decrease and greatest total volume
    'Bonus
    
     For i = 2 To 3001
    
         Ticker = ws.Cells(i, 10).Value
        
         ws.Range("R2") = WorksheetFunction.Max(Range("L2:L3001"))
         ws.Cells(2, 17).Value = Ticker
          
         ws.Range("R3") = WorksheetFunction.Min(Range("L2:L3001"))
         ws.Cells(3, 17).Value = Ticker
         
         ws.Range("R4") = WorksheetFunction.Max(Range("M2:M3001"))
         ws.Cells(4, 17).Value = Ticker
         
      Next i
    
Next ws
    
End Sub
