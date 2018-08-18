Sub tckrloop():
    Dim tckr As String
    Dim vol As Double
    Dim lastrow As Double
    Dim summary_table_row As Integer
    Dim price_op As Double
    Dim price_cl As Double
    Dim price_diff As Double
    Dim price_pct As Double
    Dim greatincr As Double
    Dim greatdecr As Double
    Dim greatvol As Double
    Dim lastrowsum As Double
            
    For Each ws In Worksheets
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        vol = 0
        summary_table_row = 2
        price_op = 0
        price_cl = 0
        price_diff = 0
        price_pct = 0
        counter = 0
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            tckr = ws.Cells(i, 1).Value
            vol = vol + ws.Cells(i, 7).Value
            price_cl = price_cl + ws.Cells(i, 6).Value
            price_diff = price_diff + (price_cl - price_op)
                If price_op > 0 Then
                price_pct = price_pct + (price_diff / price_op)
                Else: price_pct = price_pct
                End If
            ws.Range("I" & summary_table_row).Value = tckr
            ws.Range("L" & summary_table_row).Value = vol
            
            ws.Range("J" & summary_table_row).Value = price_diff
            ws.Range("K" & summary_table_row).Value = price_pct
           
            summary_table_row = summary_table_row + 1
            
            vol = 0
            price_cl = 0
            counter = 0
            price_op = 0
            price_diff = 0
            price_pct = 0
            Else
            vol = vol + ws.Cells(i, 7).Value
            counter = counter + 1
            End If
            
            If counter = 1 Then
            price_op = price_op + ws.Cells(i, 3).Value
            End If
        Next i
    
        lastrowsum = ws.Cells(Rows.Count, 9).End(xlUp).Row

            greatincr = Application.WorksheetFunction.Max(ws.Range("K:K").Value)
            greatdecr = Application.WorksheetFunction.Min(ws.Range("K:K").Value)
            greatvol = Application.WorksheetFunction.Max(ws.Range("L:L").Value)
        For j = 2 To lastrowsum
                If ws.Cells(j, 11).Value = greatincr Then
                greatincr = ws.Cells(j, 11).Value
                ws.Cells(2, 15).Value = ws.Cells(j, 9).Value
                ws.Cells(2, 16).Value = greatincr
                End If
                
                If ws.Cells(j, 11).Value = greatdecr Then
                greatdecr = ws.Cells(j, 11).Value
                ws.Cells(3, 15).Value = ws.Cells(j, 9).Value
                ws.Cells(3, 16).Value = greatdecr
                End If
                
                If ws.Cells(j, 12).Value = greatvol Then
                greatvol = ws.Cells(j, 12).Value
                ws.Cells(4, 15).Value = ws.Cells(j, 9).Value
                ws.Cells(4, 16).Value = greatvol
                End If
        Next j
        
    Next ws
 
End Sub




