Sub Stock_analyzer()
For Each ws In Worksheets

'Variables-------------------------------
    Dim ticker As String
    Dim volume As Double
    Dim Summary_table_row As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim price_change As Double
    Dim Percent_change As Double
    Dim last_amount As Long
    Dim last_row As Long

'Variables values-------------------------
    volume = 0
    Summary_table_row = 2
    last_amount = 2
'Headers for stock columns-------------------
    ws.Range("I1").Value = "ticker"
    ws.Range("J1").Value = "price_change"
    ws.Range("K1").Value = "percent_change"
    ws.Range("L1").Value = "total_volume"

'looping through for tickers and calculating the total volume------------------
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To last_row
    volume = volume + ws.Cells(i, 7).Value
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ticker = ws.Cells(i, 1).Value
        ws.Range("I" & Summary_table_row).Value = ticker
        ws.Range("l" & Summary_table_row).Value = volume
        volume = 0
'calcuating the price change from yearly open and close prices--------------------
        open_price = ws.Range("c" & last_amount)
        close_price = ws.Range("F" & i)
        price_change = close_price - open_price
        ws.Range("j" & Summary_table_row).Value = price_change
        
'calculating the yearly percentage change and to determin if there is a percent --------------
        If open_price = 0 Then
            Percent_change = 0
        Else
            open_price = ws.Range("C" & last_amount)
            Percent_change = price_change / open_price
        End If
        
'formating percentage and percent change cells to have color for (-)and (+) changes and to include percentage symbol
        ws.Range("K" & Summary_table_row).NumberFormat = "0.00%"
        ws.Range("K" & Summary_table_row).Value = Percent_change
    
        If ws.Range("J" & Summary_table_row).Value >= 0 Then
            ws.Range("J" & Summary_table_row).Interior.ColorIndex = 4
        Else
            ws.Range("J" & Summary_table_row).Interior.ColorIndex = 3
        End If
    
            Summary_table_row = Summary_table_row + 1
            last_amount = i + 1
        End If
    Next i
Next ws

End Sub
