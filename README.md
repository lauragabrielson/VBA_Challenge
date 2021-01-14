# VBA_Challenge
Wall Street VBA Homework

Sub WallStreet()

    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock_volume As Double
    Dim LastRow As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim Summary_Table_Row As Integer
    Dim ws As Worksheet
    
   
'loop through worksheets
For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    Summary_Table_Row = 2
    total_stock_volume = 0
    
    'establish last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
    'set open_price for first stock
    open_price = ws.Cells(2, 3).Value
    
  
  'loop through worksheet
    For i = 2 To LastRow
        
        'Set ticker
        ticker = ws.Cells(i, 1).Value
              
        'Check if we are still within the same ticker, if not
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
             'set close price
            close_price = ws.Cells(i, 6).Value
            
            'calculate yearly change
            yearly_change = close_price - open_price
            
            'calculate percent change
            If yearly_change <> 0 And open_price <> 0 Then
    
                percent_change = yearly_change / open_price
            
            Else: percent_change = 0
            
            End If
            
            'print values in summary table
            ws.Range("J" & Summary_Table_Row).Value = yearly_change
            
            ws.Range("K" & Summary_Table_Row).Value = percent_change
                
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            'set open price for next stock
            open_price = ws.Cells(i + 1, 3).Value
            
            
           'add total stock volume
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            
            
            'print values in Summary Table
            ws.Range("I" & Summary_Table_Row).Value = ticker
            
            ws.Range("L" & Summary_Table_Row).Value = total_stock_volume
            
            'Add one to summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            'reset total stock volume and yearly change
            total_stock_volume = 0
            yearly_change = 0
            
        'If the cell immediately following a row is the same ticker
           Else
           
            'Add to the total stock volume
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            End If
            
             'set colors for conditional formatting
             If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
                
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
                
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 2
                
            End If
            
            
            'conditional formatting for percent change - not shown in example so excluded
            'If ws.Cells(i, 11).Value > 0 Then
                'ws.Cells(i, 11).Interior.ColorIndex = 4
                
            'ElseIf ws.Cells(i, 11).Value < 0 Then
                'ws.Cells(i, 11).Interior.ColorIndex = 3
                
            'Else
                'ws.Cells(i, 11).Interior.ColorIndex = 2
                                                    
             'End If
      
    Next i
    
Next ws

End Sub








Sub WallStreetBonus()



Dim Ticker As String
Dim LastRow As Long
Dim ws As Worksheet
Dim max_increase As Double
Dim max_decrease As Double
Dim max_value As Double
   
'loop through worksheets
For Each ws In Worksheets

    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'find values
    max_increase = ws.Application.WorksheetFunction.Max(Columns("K"))
    max_decrease = ws.Application.WorksheetFunction.Min(Columns("K"))
    max_volume = ws.Application.WorksheetFunction.Max(Columns("L"))
    
    'insert values
    ws.Cells(2, 17).Value = max_increase
    ws.Cells(3, 17).Value = max_decrease
    ws.Cells(4, 17).Value = max_volume
    
    'format values
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
        

    
    'establish last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop through worksheet
    For i = 2 To LastRow
        
        'Set ticker
        Ticker = ws.Cells(i, 9).Value
        
        'find ticker
        If ws.Cells(i, 11).Value = max_increase Then
            ws.Cells(2, 16).Value = Ticker
        End If
            
        If ws.Cells(i, 11).Value = max_decrease Then
            ws.Cells(3, 16).Value = Ticker
        End If
            
        If ws.Cells(i, 12).Value = max_volume Then
            ws.Cells(4, 16).Value = Ticker
            
        End If
      
    Next i
    
Next ws

End Sub



