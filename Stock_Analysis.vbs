Sub Stock_Analysis()
'loop through all sheets
For Each ws In Worksheets

    'determine the last row
    Dim LastRow As Long
    Dim i As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'define variables and initial values
    Dim ticker As String
    Dim open_price As Double
    open_price = 0
    Dim close_price As Double
    close_price = 0
    Dim price_change As Double
    price_change = 0
    Dim percent_change As Double
    percent_change = 0
    
    
    'define variable and initial value for volume
    Dim total_vol As LongLong
    total_vol = 0
    
    'keep track of the ticker ID in the summary table
    Dim sum_table_row As Long
    sum_table_row = 2
    
    'setup table headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'set initial value of opening price for the first ticker in the ws
    open_price = ws.Cells(2, 3).Value
    
    'loop through transactions
    For i = 2 To LastRow
        'Check if we're still within the same ticker id, if not
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'set ticker id
            ticker = ws.Cells(i, 1).Value
            
            'calculate annual price change and percent change
            close_price = ws.Cells(i, 6).Value
            price_change = close_price - open_price
            'check if divisible by 0 just in case
            If open_price <> 0 Then
                percent_change = (price_change / open_price) * 100
            Else
                MsgBox ("Cannot divide by 0 for " & ticker)
            End If
                                
            'add to stock volume
            total_vol = total_vol + ws.Cells(i, 7).Value
            
            'print ticker id to summary table
            ws.Range("I" & sum_table_row).Value = ticker
            'print yearly change to summary table
            ws.Range("J" & sum_table_row).Value = price_change
            
            'format color for price_change
                If (price_change > 0) Then
                    ws.Range("J" & sum_table_row).Interior.ColorIndex = 4
                ElseIf (price_change <= 0) Then
                    ws.Range("J" & sum_table_row).Interior.ColorIndex = 3
                End If
                
            'print annual percent change to summary table
            ws.Range("k" & sum_table_row).Value = percent_change
            'print total volume to summary table
            ws.Range("L" & sum_table_row).Value = total_vol
            
            'add one to the table row
            sum_table_row = sum_table_row + 1
            'reset totals for next ticker
            price_change = 0
            close_price = 0
            'grab next ticker's open price
            open_price = ws.Cells(i + 1, 3).Value
            'reset percent and volume
            percent_change = 0
            total_vol = 0
            
        'If the next cell is the same ticker id
        Else
            'Add to the total volume
            total_vol = total_vol + ws.Cells(i, 7).Value
        End If
            
    Next i

Next ws

End Sub