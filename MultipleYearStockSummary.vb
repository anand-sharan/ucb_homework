Sub MultipleYearStockSummary()

    'Create and initialize the local variables of Sub Multiple Year Stock Summary
    
    'The Stock Name of the current stock begin processed
    Dim stock_name As String
    stock_name = ""
    
    'Yearly open price of the current stock begin processed
    Dim yearly_open_price As Double
    yearly_open_price = 0
    
    'Yearly close price of the current stock begin processed
    Dim yearly_close_price As Double
    yearly_close_price = 0
    
    'Yearly price change (close - open) of the current stock begin processed
    Dim yearly_price_change As Double
    yearly_price_change = 0
        
    ' Yearly price change ((close - open ) / open)*100 of the current stock begin processed
    Dim yearly_price_change_percentage As Double
    yearly_price_change_percentage = 0
    
    'Yearly stock volume of the current stock begin processed
    Dim yearly_total_stock_vol As Double
    yearly_total_stock_vol = 0
    
    'The last Row number of data in the current worksheet
    Dim last_row_of_data As Long
    last_row_of_data = 0
    
    'The last Row number of data in the summary section on the current worksheet
    Dim last_row_of_summary As Long
    last_row_of_summary = 1
    
    'Set default color index to White
    Dim color_index As Integer
    color_index = 2
    
    'Greatest percentage increase
    Dim greatest_per_incr As Double
    greatest_per_incr = 0
    
    'Ticker of Greatest percentage increase
    Dim ticker_greatest_per_incr As String
    ticker_greatest_per_incr = ""
    
    'Greatest percentage decrease
    Dim greatest_per_decr As Double
    greatest_per_decr = 0
    
    'Ticker of Greatest percentage decrease
    Dim ticker_greatest_per_decr As String
    ticker_greatest_per_decr = ""
    
    'Greatest total volume increase
    Dim greatest_total_vol As Double
    greatest_total_vol = 0
    
    'Ticker of Greatest percentage decrease
    Dim ticker_greatest_total_vol As String
    ticker_greatest_total_vol = ""
    
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------

    For Each ws In Worksheets
        
        ' -----------------------------------------------
        ' ASSIGN HEADERS TO THE SUMMARY SECTION
        ' -----------------------------------------------
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize the first stock name
        stock_name = ws.Cells(2, 1).Value
        
        ' Initialize the first stock's open price
        yearly_open_price = ws.Cells(2, 3).Value
        
        ' Initialize the total stock volume
        yearly_total_stock_vol = ws.Cells(2, 7).Value
        
        ' Find the last row of each worksheet
        last_row_of_data = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        ' Initialize the last row number of summary data
        last_row_of_summary = 1
        ' --------------------------------------------
        ' LOOP THROUGH ALL DATA IN CURRENT WORKSHEET
        ' --------------------------------------------
        
        For i = 3 To last_row_of_data
            
            ' ---------------------------------------------------------------------------------------
            ' CHECK IF NEXT ROW IS FOR DIFFERENT STOCK, IF YES THEN COMPLETE CACULATIONS FOR SUMMARY
            ' ----------------------------------------------------------------------------------------
                        
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            
                yearly_total_stock_vol = yearly_total_stock_vol + ws.Cells(i, 7).Value
            
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Assign the yearly close price
                yearly_close_price = ws.Cells(i, 6).Value
                
                ' Calcuate yearly price change
                yearly_price_change = yearly_close_price - yearly_open_price
                
                ' Assign Green, Red or White color based on Yearly Price Change (>0, <0, =0) values respectively
                If yearly_price_change < 0 Then
                    color_index = 3
                ElseIf yearly_price_change > 0 Then
                    color_index = 4
                Else
                    color_index = 2
                End If
                
                ' Avoid divide by zero error
                If yearly_open_price <> 0 Then
                   yearly_price_change_percentage = (yearly_price_change / yearly_open_price) * 100
                Else
                   yearly_price_change_percentage = 0
                End If
                
                'Add cumulatively to arrive at the year total stock volume
                yearly_total_stock_vol = yearly_total_stock_vol + ws.Cells(i, 7).Value
                
                'Find the next empty row number in the summary section on the current worksheet
                last_row_of_summary = last_row_of_summary + 1
                
                'Append ( +1 ) new row to the summary section
                ws.Cells(last_row_of_summary, 9).Value = stock_name
                ws.Cells(last_row_of_summary, 10).Value = yearly_price_change
                ws.Cells(last_row_of_summary, 11).Value = yearly_price_change_percentage
                ws.Cells(last_row_of_summary, 12).Value = yearly_total_stock_vol
                
                'Set the Cell Colors
                ws.Cells(last_row_of_summary, 10).Interior.ColorIndex = color_index
                
                'Set next stock name
                stock_name = ws.Cells(i + 1, 1).Value
                
                'Set next stock open price
                yearly_open_price = ws.Cells(i + 1, 3).Value
                
                yearly_total_stock_vol = ws.Cells(i + 1, 7).Value
    
            End If
        Next i
    
        ' -----------------------------------------------
        ' COMPLETE THE GREATEST SUMMARY SECTION
        ' -----------------------------------------------
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Greatest percentage increase
        greatest_per_incr = 0
    
        'Greatest percentage decrease
        greatest_per_decr = 0
    
        'Greatest total volume increase
        greatest_total_vol = 0
        
        'Assign the last row of summary section
        'last_row_of_summary = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        ' Loop through Yearly change summary through range of summary rows
        For j = 2 To last_row_of_summary
            If ws.Cells(j, 11).Value > greatest_per_incr Then
               ticker_greatest_per_incr = ws.Cells(j, 9).Value
               greatest_per_incr = ws.Cells(j, 11).Value
            End If
            
            If ws.Cells(j, 11).Value < greatest_per_decr Then
               ticker_greatest_per_decr = ws.Cells(j, 9).Value
               greatest_per_decr = ws.Cells(j, 11).Value
            End If
            
            If ws.Cells(j, 12).Value > greatest_total_vol Then
                ticker_greatest_total_vol = ws.Cells(j, 9).Value
                greatest_total_vol = ws.Cells(j, 12).Value
            End If
        Next j
        
        'Load the Greatest summary section values
        ws.Cells(2, 16).Value = ticker_greatest_per_incr
        ws.Cells(2, 17).Value = greatest_per_incr
        ws.Cells(3, 16).Value = ticker_greatest_per_decr
        ws.Cells(3, 17).Value = greatest_per_decr
        ws.Cells(4, 16).Value = ticker_greatest_total_vol
        ws.Cells(4, 17).Value = greatest_total_vol
        
        ' -----------------------------------------------
        ' FORMAT CURRENT WORKSHEET
        ' -----------------------------------------------
        ' Format percentage columns
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ' Autofit to display data
        ws.Columns("A:Q").AutoFit
    Next ws

End Sub
