Attribute VB_Name = "Module1"
Sub populate_tickerlist()
    Dim ws As Worksheet, i As Long, last_row As Long, result_table_row As Integer
    Dim open_price As Double, close_price As Double, yearly_change As Double, yearly_change_percentage As Double, total_stock_vol As LongLong
    Dim greatest_increase_ticker As String, greatest_increase_percentage As Double, greatest_decrease_ticker As String, greatest_decrease_percentage As Double, greatest_total_ticker As String, greatest_total_volume As LongLong
    
    ''Loop through Worksheets
    For Each ws In Worksheets
        ''Set Result Table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        ''Count number of rows
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ''Initialize Values
        result_table_row = 2
        total_stock_vol = 0
        greatest_increase_ticker = ""
        greatest_increase_percentage = 0
        greatest_decrease_ticker = ""
        greatest_decrease_percentage = 0
        greatest_total_ticker = ""
        greatest_total_volume = 0
        
        ''Print the First Ticker value
        ws.Cells(result_table_row, 9).Value = ws.Cells(2, 1).Value
        
        ''Set Opening Price of the first ticker
        open_price = ws.Cells(2, 3).Value
        
        ''Loop through rows
        For i = 2 To last_row
        total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                '' set close price for the previous ticker and yearly change before open_price gets overridden.
                close_price = ws.Cells(i, 6).Value
                yearly_change = close_price - open_price
                
                '' Div by 0 error handling
                If open_price <> 0 Then
                    yearly_change_percentage = yearly_change / open_price
                Else
                    yearly_change_percentage = 0
                End If
                
                '' Find greatest increase percentage by comparing it with a previous value to find a maximum.
                If yearly_change_percentage > greatest_increase_percentage Then
                    greatest_increase_percentage = yearly_change_percentage
                    greatest_increase_ticker = ws.Cells(i, 1).Value
                End If
                
                '' Find greatest decrease percentage by comparing it with a previous value to find a minimum.
                If yearly_change_percentage < greatest_decrease_percentage Then
                    greatest_decrease_percentage = yearly_change_percentage
                    greatest_decrease_ticker = ws.Cells(i, 1).Value
                End If
                
                If total_stock_vol > greatest_total_volume Then
                    greatest_total_volume = total_stock_vol
                    greatest_total_ticker = ws.Cells(i, 1).Value
                End If
                            
                '' set calculated values to result table
                ws.Cells(result_table_row, 10).Value = yearly_change
                ws.Cells(result_table_row, 11).Value = Format(yearly_change_percentage, "0.00%")
                ws.Cells(result_table_row, 12).Value = total_stock_vol
                
                ' set percentage change cell background color to green for positive values and red for negative values
                If yearly_change > 0 Then
                    ws.Cells(result_table_row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(result_table_row, 10).Interior.ColorIndex = 3
                End If
                
                '' set result_table_row to next row
                result_table_row = result_table_row + 1
                
                '' reset total_stock_vol to next 0 to calculate total for the next ticker
                total_stock_vol = 0
                
                '' print the next ticker value (A, AA, etc.)
                ws.Cells(result_table_row, 9).Value = ws.Cells(i + 1, 1).Value
                
                '' set open price for the next ticker
                open_price = ws.Cells(i + 1, 3).Value
                    
            End If
        Next i
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = greatest_increase_ticker
        ws.Cells(2, 16).Value = Format(greatest_increase_percentage, "0.00%")
        
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = greatest_decrease_ticker
        ws.Cells(3, 16).Value = Format(greatest_decrease_percentage, "0.00%")
        
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = greatest_total_ticker
        ws.Cells(4, 16).Value = greatest_total_volume
    
    Next ws
End Sub




