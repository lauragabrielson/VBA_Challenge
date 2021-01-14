# VBA_Challenge
Wall Street VBA Homework
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



