Attribute VB_Name = "Module1"
Sub StockLoop()

    ' Declare current as a worksheet object variable
    Dim Current As Worksheet


    ' Loop through all worksheets
    For Each Current In Worksheets
    
        ' Create column and row headings
            Current.Cells(1, 9).Value = "Ticker"
            Current.Range("J1").Value = "Yearly Change"
            Current.Range("K1").Value = "Percent Change"
            Current.Range("L1").Value = "Total Stock Volume"
            Current.Cells(1, 16).Value = "Ticker"
            Current.Cells(1, 17).Value = "Value"
            Current.Range("O2").Value = "Greatest % Increase"
            Current.Range("O3").Value = "Greatest % Decrease"
            Current.Range("O4").Value = "Greatest Total Volume"

            ' Definitions and variables
            Dim Ticker As String
            Ticker = ""
            Dim TickerRow As Long
            TickerRow = 1
            Dim open_price As Double
            open_price = 0
            Dim close_price As Double
            close_price = 0
            Dim stock_volume As Double
            stock_volume = 0
            Dim price_change As Double
            price_change = 0
            Dim price_change_percent As Double
            price_change_percent = 0
            Dim max_increase As Double
            max_positive = 0
            Dim max_decrease As Double
            max_decrease = 0
            Dim LastRow As Long
            LastRow = Range("A:A").SpecialCells(xlCellTypeLastCell).Row
             
        
                ' Loop through all stocks
                For i = 1 To LastRow
            
                    If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then
                    
                        ' Output ticker symbol
                        TickerRow = TickerRow + 1
                        Ticker = Current.Cells(i + 1, 1).Value
                        Current.Cells(TickerRow, 9).Value = Ticker
                        
                        ' Find opening price
                        open_price = Current.Cells(i + 1, 3).Value
                                
                         
                                If i > 1 Then
                                    
                                    close_price = Current.Cells(i, 6).Value
                                    price_change = close_price - open_price
                                    
                                        ' Can't divide by 0
                                        If open_price = 0 Then
                                        
                                            price_change_percent = 0
                                        
                                            
                                        ElseIf open_price <> 0 Then
                                        
                                            price_change_percent = (price_change / open_price) * 100
                            
                                    ' Output difference from opening price column C to closing price column F
                                    Current.Cells(TickerRow, "J").Value = price_change

                                    ' Output percentage represented by that difference
                                    Current.Cells(TickerRow, "K").Value = price_change_percent
                                    
                                    
                                    
                                    End If
                                        
                
                                End If
                
                        End If
                        
                    
            
                Next i
            
            ' Dim range for each ticker to get total stock volume
            Dim rng As Range, cell As Range
            Set rng = Range("A2:A" & LastRow)
                
                ' Loop through each stock range
                For Each cell In rng
    
                    If cell = cell.Offset(0, 8) Then
        
                        stock_volume = stock_volume + cell.Offset(0, 6)
    
                    Else
                
                        stock_volume = 0 + cell.Offset(0, 6)
                        cell.Offset(0, 6) = stock_volume
                        
                        ' Output stock_volume into columnn L
                        cell.Offset(0, 11).Value = stock_volume
                      
                    End If
        
                Next cell
            
            ' Loop for conditional formatting for price_change and price_change_percent
            For i = 2 To LastRow
            
                ' Check if price_change is positive to conditionally format green
                If Current.Cells(i, 10).Value > 0 Then
                                    
                    Current.Cells(i, 10).Interior.ColorIndex = 4
                                        
                ' Check if negative to conditionally format red
                ElseIf Current.Cells(i, 10).Value < 0 Then
                                    
                    Current.Cells(i, 10).Interior.ColorIndex = 3
                                              
                                        
                End If
                
            
                ' Check if price_change_percent is positive to conditionally format green
                If Current.Cells(i, 11).Value > 0 Then
                                    
                    Current.Cells(i, 11).Interior.ColorIndex = 4
                                        
                ' Check if negative to conditionally format red
                ElseIf Current.Cells(i, 11).Value < 0 Then
                                    
                    Current.Cells(i, 11).Interior.ColorIndex = 3
                                            
                                                       
                End If
                
            Next i
                    
    
    Next Current

End Sub
