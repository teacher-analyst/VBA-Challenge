Sub stockmarket()

Dim Summary_Table_Row As Integer
Dim Stockvol_total As Double
Dim Ticker As String
Dim LR As Long
Dim ws As Worksheet
Dim Opening_Price As Double
Dim stockPriceCaptured As Boolean


For Each ws In Worksheets

    Summary_Table_Row = 2
    Stockvol_total = 0
    Opening_Price = ws.Cells(2, 3).Value

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
    
        If stockPriceCaptured = False Then
            Opening_Price = ws.Cells(i, 3).Value
            
            stockPriceCaptured = True
        End If
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         
            stockPriceCaptured = False
         
         ' Set stock name
            Ticker = ws.Cells(i, 1).Value
            
            'Print ticker in summary table
          
            ws.Range("J" & Summary_Table_Row).Value = Ticker
            
            'Add or subtract closing price from first opening price
            Yearly_Change = ws.Cells(i, 6).Value - Opening_Price
    
            ' Calculate percent change from opening to closing price
            Percent_Change = Yearly_Change / Opening_Price
          
          'Add to stock volume total
            Stockvol_total = Stockvol_total + ws.Cells(i, 7)
           
            'Print total stock volume to summary table
            ws.Range("M" & Summary_Table_Row).Value = Stockvol_total
            'Print yearly change in summary table
            ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
            
            'Print percentage change in summary table
            ws.Range("L" & Summary_Table_Row).Value = FormatPercent(Percent_Change)
          
            'conditional formatting percetnage change column
        
            Select Case Percent_Change
            Case Is > 0
                ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4 'Green
            Case Else
                ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3  'Red
      
            End Select
            
            'conditional formatting yearly change column
            Select Case Yearly_Change
            Case Is > 0
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4 'Green
            Case Else
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3  'Red
      
            End Select
            'Add one to summary table
            Summary_Table_Row = Summary_Table_Row + 1
            
            Stockvol_total = 0
            
        
         Else
         
         ' add to the Stock vol total
         Stockvol_total = Stockvol_total + ws.Cells(i, 7).Value
        
         End If  
        
    Next i
     
        
Next ws
         
End Sub
