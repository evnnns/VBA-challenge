Attribute VB_Name = "Module1"
Sub StockLoop()

    For Each ws In Worksheets

    Dim lastRow As String
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim ticker_count As Long
    ticker_count = 0
    
    Dim year_open As Double
    
    Dim year_close As Double
    
    Dim yearly_change As Double
    
    Dim percent_change As Double
    
    Dim percent_change_dec As Double
    
    Dim stock_volume As Single
    
    'Set column names
    ws.Cells(1, 9).Value = "Ticker Name"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    For ticker_count = 1 To 1
    
        For i = 2 To lastRow
        
            'Record first year open
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                year_open = ws.Cells(i, 3).Value
                stock_volume = ws.Cells(i, 7).Value
                
            End If

            'Write tickers to new column
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ticker_count = ticker_count + 1
                ws.Cells(ticker_count, 9).Value = ws.Cells(i, 1).Value
                year_close = ws.Cells(i, 6).Value
                'Write Yearly Change
                ws.Cells(ticker_count, 10).Value = year_close - year_open
                yearly_change = ws.Cells(ticker_count, 10).Value
                
                'Percentage Change
                percent_change = yearly_change / year_open
                Cells(ticker_count, 11).Value = percent_change
                                                              
            End If
            
            'Format percentages
            Cells(i, 11).NumberFormat = "0.00%"
            
            'Shade cells red/green if they are negative/positive for yearly change
            If yearly_change < 0 Then
                ws.Cells(ticker_count, 10).Interior.ColorIndex = 3
                
            ElseIf yearly_change > 0 Then
                ws.Cells(ticker_count, 10).Interior.ColorIndex = 4
                
            Else
                ws.Cells(ticker_count, 10).Interior.ColorIndex = 0
                
            End If
            
            'Sum Total Stock Volume
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                stock_volume = stock_volume + ws.Cells(i + 1, 7).Value
                ws.Cells(ticker_count + 1, 12).Value = stock_volume
            End If
                                                            
        Next i
    
    Next ticker_count
    
    Next ws
    
End Sub
