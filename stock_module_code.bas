Attribute VB_Name = "Module1"
Sub stocks():


    
    Dim w As Integer
    
    For w = 1 To ActiveWorkbook.Worksheets.Count
        
        Worksheets(w).Activate
        
    
        Dim last_row As Double
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
                
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
            
        Dim output_row As Integer
            output_row = 2
                 
        Dim ticker As String
            ticker = Cells(2, 1).Value
        
        Dim open_price As Double
            open_price = Cells(2, 3).Value
        Dim open_date As Long
            open_date = Cells(2, 2).Value
            
        Dim close_price As Double
            close_price = Cells(2, 6).Value
        Dim close_date As Long
            close_date = Cells(2, 2).Value
            
        Dim trading_volume As Double
            trading_volume = Cells(2, 7).Value
            
        Dim this_ticker As String
        
        Dim price_change As Double
        
        Dim percent_change As String
        
                
        For i = 3 To last_row
        
            this_ticker = Cells(i, 1).Value
            
            Dim this_date As Long
            this_date = Cells(i, 2).Value
            
            If this_ticker <> ticker Then
            
                Cells(output_row, 9).Value = ticker
                
                price_change = close_price - open_price
                Cells(output_row, 10).Value = price_change
                    
                If close_price < open_price Then
                    
                    Cells(output_row, 10).Interior.ColorIndex = 3
                
                Else
                    
                    Cells(output_row, 10).Interior.ColorIndex = 4
                
                End If
                
                percent_change = CStr(Round(100 * price_change / open_price, 2))
                Cells(output_row, 11).Value = percent_change & "%"
                Cells(output_row, 12).Value = trading_volume
                
                output_row = output_row + 1
                
                ticker = this_ticker
                
                open_price = Cells(i, 3).Value
                open_date = this_date
                
                close_price = Cells(i, 6).Value
                close_date = this_date
                
                trading_volume = Cells(i, 7).Value
                    
            Else
            
                trading_volume = trading_volume + Cells(i, 7).Value
            
                If this_date < open_date Then
                    
                    open_date = this_date
                    open_price = Cells(i, 3).Value
                    
                End If
                
                If this_date > close_date Then
                        
                    close_date = this_date
                    close_price = Cells(i, 6).Value
                    
                End If
                
                If i = last_row Then
                    
                    Cells(output_row, 9).Value = ticker
                
                    price_change = close_price - open_price
                    Cells(output_row, 10).Value = price_change
                        
                    If close_price < open_price Then
                        
                        Cells(output_row, 10).Interior.ColorIndex = 3
                    
                    Else
                        
                        Cells(output_row, 10).Interior.ColorIndex = 4
                    
                    End If
                    
                    percent_change = price_change / open_price
                    Cells(output_row, 11).Value = percent_change
                    Cells(output_row, 11).NumberFormat = "0.00%"
                    Cells(output_row, 12).Value = trading_volume
                    
                End If
                
            End If
                
        Next i
        
        
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        Dim max_increase As Double
        max_increase = 0
        Dim increase_ticker As String
        
        Dim max_decrease As Double
        max_decrease = 0
        Dim decrease_ticker As String
        
        Dim max_volume As Double
        max_volume = 0
        Dim volume_ticker As String
        
        Dim last_ticker_row As Double
        last_ticker_row = Cells(Rows.Count, 9).End(xlUp).Row
        
        
        For i = 2 To last_ticker_row
            
            this_ticker = Cells(i, 9).Value
            
            Dim this_change_value As Double
            this_change_value = Cells(i, 11).Value
    
            Dim this_volume As Double
            this_volume = Cells(i, 12).Value
            
            If this_change_value > max_increase Then
                
                max_increase = this_change_value
                increase_ticker = this_ticker
                
            End If
            
            If this_change_value < max_decrease Then
            
                max_decrease = this_change_value
                decrease_ticker = this_ticker
                
            End If
            
            If this_volume > max_volume Then
            
                max_volume = this_volume
                volume_ticker = this_ticker
                
            End If
            
            
            Cells(2, 16).Value = increase_ticker
            Cells(2, 17).Value = max_increase
            Cells(2, 17).NumberFormat = "0.00%"
            
            Cells(3, 16).Value = decrease_ticker
            Cells(3, 17).Value = max_decrease
            Cells(3, 17).NumberFormat = "0.00%"
            
            Cells(4, 16).Value = volume_ticker
            Cells(4, 17).Value = max_volume
            
        Next i


    Next w
    
    
        
End Sub
