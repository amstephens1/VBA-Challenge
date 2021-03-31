Attribute VB_Name = "Module1"
Sub alphabetical_testing()
    
 Dim ws As Worksheet
    
        For Each ws In Worksheets
               
            
            Dim ticker_symbol As String
            Dim t_symbol As String
            Dim stock_volume As LongLong
            Dim open_price As Double
            Dim close_price As Double
            Dim counter As Integer
            Dim Summary_Table_Row As Long
            Dim max_increase As Double
            Dim max_decrease As Double
            Dim max_total_volume As LongLong
            
            max_total_volume = 0
            max_decrease = 0
            max_increase = 0
            counter = 0
            Summary_Table_Row = 2
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            
            
                For i = 2 To lastrow
                
                    counter = counter + 1
            
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        
                                    
                        ticker_symbol = ws.Cells(i, 1).Value
                        
                      
                        open_price = ws.Cells(i - (counter - 1), 3).Value
                                           
                        close_price = ws.Cells(i, 6).Value
                        
                        stock_volume = stock_volume + ws.Cells(i, 7).Value
                        ws.Range("I" & Summary_Table_Row).Value = ticker_symbol
                        ws.Range("I1") = "Ticker Symbol"
                        ws.Range("J" & Summary_Table_Row).Value = (close_price - open_price)
                        ws.Range("J1") = "Year Change"
                        If open_price <> 0 Then
                            ws.Range("K" & Summary_Table_Row).Value = FormatPercent((close_price / open_price) - 1)
                        Else
                            ws.Range("K" & Summary_Table_Row).Value = 0
                        End If
                        ws.Range("K1") = "Percent Year Change"
                        ws.Range("L" & Summary_Table_Row).Value = stock_volume
                        ws.Range("L1") = "Total Stock Volume"
                        
                        Summary_Table_Row = Summary_Table_Row + 1
                        
                        stock_volume = 0
                        counter = 0
                        
                    Else
                     
                        stock_volume = stock_volume + ws.Cells(i, 7).Value
                              
                    End If
                    
                Next i
            
                For i = 2 To lastrow
                
                    If ws.Cells(i, 10) > 0 Then
                    
                        ws.Cells(i, 10).Interior.ColorIndex = 4
                        
                    ElseIf ws.Cells(i, 10) < 0 Then
                    
                        ws.Cells(i, 10).Interior.ColorIndex = 3
                        
                    End If
                    
                Next i
                
                    ws.Range("O2") = "Greatest % Increase"
                    ws.Range("O3") = "Greatest % Decrease"
                    ws.Range("O4") = "Greatest Total Volume"
                    ws.Range("P1") = "Ticker Symbol"
                    ws.Range("Q1") = "Value"

                
                For i = 2 To lastrow
                
                
                
                    If ws.Cells(i, 11).Value > max_increase Then
                    
                        max_increase = ws.Cells(i, 11).Value
                        
                        ws.Range("P2") = ws.Cells(i, 9).Value
                       
                    End If
                    
               
               
                    If ws.Cells(i, 11).Value < max_decrease Then
                    
                        max_decrease = ws.Cells(i, 11).Value
                        
                        ws.Range("P3") = ws.Cells(i, 9).Value
                       
                    End If
                    
               
               
                    If ws.Cells(i, 12).Value > max_total_volume Then
                    
                        max_total_volume = ws.Cells(i, 12).Value
                        
                        ws.Range("P4") = ws.Cells(i, 9).Value
                       
                    End If
                    

                Next i
            ws.Range("Q2") = FormatPercent(max_increase)
            ws.Range("Q3") = FormatPercent(max_decrease)
            ws.Range("Q4") = max_total_volume
            ws.UsedRange.EntireColumn.AutoFit
                
        Next ws
      
      
    
End Sub

