Sub year_stock():
    'worksheets is an array of sheets
    For Each ws In Worksheets
    
    'set the header of Summary Table
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
      
    'MsgBox Worksheets.Count
     
    'set an initial variable for holding the ticker
     Dim Ticker_Name As String
    
    'Set an initial variable for opening price and closing price for the same ticker
     Dim Opening_price, Closing_price As Double
     
    'Set an initial variable for yearly change and percent change as double
     Dim Yearly_Change, Percent_Change As Double
     
    'Set an initial variable for total stock volume
     Dim Total_Stock As Double
    
    'Keep track of the location for each ticker in the summary table
     Dim Summary_Table_Row As Integer
     Summary_Table_Row = 2
     
    'Set the first opening price for the calculation
     Opening_price = ws.Cells(2, 3).Value
     Closing_price = 0
     
    'Set the last row
     Last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    'Loop through all ticker data
     For i = 2 To Last_row
        
            'check if we are still within the same ticker, if is not
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                'Set the ticker
                 Ticker_Name = ws.Cells(i, 1).Value
    
                'Print the ticker in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
            
                'Set the closing price for the calculation
                Closing_price = ws.Cells(i, 6).Value
                
                'Calculate the yearly change
                Yearly_Change = Closing_price - Opening_price
                
                'Conditional formating for yearly Change
                If Yearly_Change > 0 Then
                
                    'print the yearly change in Green to the summary Table
                     ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                     ws.Range("J" & Summary_Table_Row).FormatConditions.Delete
                     ws.Range("J" & Summary_Table_Row).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=0
                     ws.Range("J" & Summary_Table_Row).FormatConditions(1).Interior.ColorIndex = 4
                Else
                    'print the yearly change in Red to the summary Table
                     ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                     ws.Range("J" & Summary_Table_Row).FormatConditions.Delete
                     ws.Range("J" & Summary_Table_Row).FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, Formula1:=0
                     ws.Range("J" & Summary_Table_Row).FormatConditions(1).Interior.ColorIndex = 3
                End If
                
                'Calculate the percent change
                Percent_Change = Yearly_Change / Opening_price

                
                'print the percent change in the summary table
                 ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                 
                 'Set the format for percent change
                 ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                 
                'set open price of next ticker's row
                Opening_price = ws.Cells(i + 1, 3).Value
                
                'Add stock volume to Total stock
                Total_Stock = Total_Stock + ws.Cells(i, 7).Value
                
                'print the total stock volume in the summary table
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock
                
                'Reset total stock volue equal to zero
                 Total_Stock = 0
                 
                'Add one to the summary table
                 Summary_Table_Row = Summary_Table_Row + 1
        
            'IF the cell immediately follwoing a row is the same ticker
            Else
        
                'Add stock volume to Total stock
                
                Total_Stock = Total_Stock + ws.Cells(i, 7).Value
                
           End If

    Next i
 
    'Set variables for Greatest % Increase and its row
    Dim Max_Inc, Min_Inc, Total_Vol As Double
    Dim MAX_row, Min_row, Vol_row As Integer
    Dim Max_Ticker, Min_Ticker, Vol_Ticker As String
        

    'Get Greatest% Increase and Its row
     Max_Inc = Application.WorksheetFunction.Max(ws.Range("K2", "K" & Last_row))
     ws.Cells(2, 17).Value = Max_Inc
     ws.Cells(2, 17).NumberFormat = "0.00%"
     MAX_row = WorksheetFunction.Match(ws.Range("Q2"), ws.Range("K2:K" & Last_row), 0) + 1
     Max_Ticker = ws.Cells(MAX_row, 9).Value
     ws.Cells(2, 16).Value = Max_Ticker

    'Get minimum % Increase and its row
     Min_Inc = Application.WorksheetFunction.Min(ws.Range("K2", "K" & Last_row))
     ws.Cells(3, 17).Value = Min_Inc
     ws.Cells(3, 17).NumberFormat = "0.00%"
     Min_row = WorksheetFunction.Match(ws.Range("Q3"), ws.Range("K2:K" & Last_row), 0) + 1
     Min_Ticker = ws.Cells(Min_row, 9).Value
     ws.Cells(3, 16).Value = Min_Ticker

    'Get Greatest Total Volume and its row
     Total_Vol = Application.WorksheetFunction.Max(ws.Range("L2", "L" & Last_row))
     ws.Cells(4, 17).Value = Total_Vol
     Vol_row = WorksheetFunction.Match(ws.Range("Q4"), ws.Range("L2:L" & Last_row), 0) + 1
     Vol_Ticker = ws.Cells(Vol_row, 9).Value
     ws.Cells(4, 16).Value = Vol_Ticker
     

 Next ws

End Sub



