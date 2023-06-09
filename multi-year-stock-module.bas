Sub stocks():

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
        
    Dim price_change As Double
    Dim percent_change As String
    
            
    For i = 3 To last_row
    
        Dim this_ticker As String
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
                
                percent_change = CStr(Round(100 * price_change / open_price, 2))
                Cells(output_row, 11).Value = percent_change & "%"
                Cells(output_row, 12).Value = trading_volume
                
            End If
            
        End If
            
    Next i
    
End Sub
