Attribute VB_Name = "Module1"
'calculate the yearly change
'Column "k" = Cells("F252" - "C2").Value


'subtract opening price from closing price for Year Change

'Retrieve and store the data values in each variable

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

    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"



    'print the row names
    ws.Cells(2, 15).Value = "Greatest % Increase "
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"


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
            ws.Cells(ticker_value, 10).Interior.Color = RGB(255, 255, 0)
        
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

