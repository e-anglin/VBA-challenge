Attribute VB_Name = "Module2"

Sub Greatest()

For Each ws In ThisWorkbook.Worksheets
Dim Sheet As String

'print the column names
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

'print the row names
    ws.Cells(2, 15).Value = "Greatest % Increase "
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
  'Reset Loop for New Table
    lastrow9 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Declare values for the New Table
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total_Volume As Double
    
    'Find values of highest and lowest performing stocks
    Greatest_Increase = ws.Cells(2, 11).Value
    Greatest_Decrease = ws.Cells(2, 11).Value
    Greatest_Total_Volume = ws.Cells(2, 12).Value
    
    'For Loop for final IF Functions
    For i = 2 To lastrow9
    
        'Commence IF for Greatest Percent Increase
        If ws.Cells(i, 11).Value > Greatest_Increase Then
            Greatest_Increase = ws.Cells(i, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
            Else
            Greatest_Increase = Greatest_Increase
            
        End If
        
        'Commence IF for Greatest Percent Decrease
        If ws.Cells(i, 11).Value < Greatest_Decrease Then
            Greatest_Decrease = ws.Cells(i, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
            Else
            Greatest_Decrease = Greatest_Decrease
            
        End If
        
        'Commence IF for Greatest Total Volume
        If ws.Cells(i, 12).Value > Greatest_Total_Volume Then
            Greatest_Total_Volume = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            Else
            Greatest_Total_Volume = Greatest_Total_Volume
            
        End If
        
      'Don't forget about that percent format!
      Range("Q2:Q3").NumberFormat = "0.00%"
      
      Next i
      
    Next ws
    
End Sub
