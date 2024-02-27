Sub ticker()

    For Each ws in Worksheets

    
        Dim name As String
    
        Dim volume As Double
        volume = 0

        Dim row_num As Integer
        row_num = 2
        
        Dim open_price As Double
        open_price = ws.Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row


        For i = 2 To last_row

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              name = ws.Cells(i, 1).Value

              volume = volume + ws.Cells(i, 7).Value

              ws.Range("I" & row_num).Value = name

              ws.Range("L" & row_num).Value = volume

              close_price = ws.Cells(i, 6).Value

              yearly_change = (close_price - open_price)
              
              ws.Range("J" & row_num).Value = yearly_change

                If open_price = 0 Then
                    percent_change = 0
                
                Else
                    percent_change = yearly_change / open_price
                
                End If

              ws.Range("K" & row_num).Value = percent_change
              ws.Range("K" & row_num).NumberFormat = "0.00%"
   
              row_num = row_num + 1

              volume = 0

              open_price = ws.Cells(i + 1, 3)
            
            Else
              
              volume = volume + ws.Cells(i, 7).Value

            
            End If
        
        Next i


    table_last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        For i = 2 To table_last_row
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i

        For i = 2 To table_last_row
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 10
            
            Else
                ws.Cells(i, 11).Interior.ColorIndex = 3
            
            End If
        
        Next i

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        For i = 2 To table_last_row
        
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & table_last_row)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & table_last_row)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & table_last_row)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    
    Next ws
        
End Sub