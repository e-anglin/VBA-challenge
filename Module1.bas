Attribute VB_Name = "Module1"

'_____________________________________________________


Sub Multi_Year_Stock_Data()
'set the initial worksheet Loop through all sheets
For Each ws In ThisWorkbook.Worksheets
Dim Sheet As String


    'print the column names
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Year Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

   


    


    'Set up the second loop for the ticker values (year, percentage change, and total stock volume)
    ticker_value = 2
    opening_value = 2

    'Set up the last row to reset the loop
    lastrow1 = Cells(Rows.Count, 1).End(xlUp).Row

    'Commence For Loop for If Functions
    For i = 2 To lastrow1
        
        'Commence IF to find the stock values
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Identify the stock
            ws.Cells(ticker_value, 9).Value = ws.Cells(i, 1).Value
            'Findthe Yearly Change
            ws.Cells(ticker_value, 10).Value = ws.Cells(i, 6).Value - ws.Cells(opening_value, 3).Value
        
        'Commence IF for color values
        If ws.Cells(ticker_value, 10).Value > 0 Then
            ws.Cells(ticker_value, 10).Interior.Color = RGB(128, 255, 0)
            Else
            ws.Cells(ticker_value, 10).Interior.Color = RGB(220, 20, 60)
        
        End If
        
        'Commence IF to calculate the Percent Change
        If ws.Cells(opening_value, 3).Value <> 0 Then
            percentchange = ((ws.Cells(i, 6).Value - ws.Cells(opening_value, 3).Value) / ws.Cells(opening_value, 3).Value)
            ws.Cells(ticker_value, 11).Value = Format(percentchange, "Percent")
            Else
            ws.Cells(ticker_value, 11).Value = Format(0, "Percent")
            
        End If
        
        'Calculate Total Stock Volume
        ws.Cells(ticker_value, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(opening_value, 7), ws.Cells(i, 7)))
        ticker_value = ticker_value + 1
        opening_value = i + 1
        
        End If
        
    Next i
    
Next ws
    
    
        
End Sub

