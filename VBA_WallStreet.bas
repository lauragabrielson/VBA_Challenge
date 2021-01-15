Attribute VB_Name = "Module1"
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


