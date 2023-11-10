Attribute VB_Name = "Module1"
Sub Test():

    For Each ws In Worksheets
        Dim new_start, year_end, total_volume, percent_change, summary_row As Integer
        Dim Last_Row As Long
        
        
        
        Last_Row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        new_start = 2
        year_end = 0
        total_volume = 0
        percent_change = 0
        summary_row = 2
    
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        
        ws.Columns("I:Q").AutoFit
      
                
        For i = 2 To Last_Row
            
            year_begin = ws.Cells(new_start, 3).Value
            
            
            If (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value) Then
            
                total_volume = total_volume + ws.Cells(i, 7).Value
                
            
            Else
            
                year_end = ws.Cells(i, 6).Value
                
                
                ws.Cells(summary_row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(summary_row, 10).Value = (year_end - year_begin)
                
                If ws.Cells(summary_row, 10).Value >= 0 Then
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 4
                
                Else
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 3
                    
                End If
                
                percent_change = ((year_end / year_begin) - 1)
                
                
                    
                ws.Cells(summary_row, 11).Value = FormatPercent(percent_change)
                ws.Cells(summary_row, 12).Value = total_volume
                ws.Cells(summary_row, 12).NumberFormat = "0"
                
                
                summary_row = summary_row + 1
                total_volume = 0
                new_start = i + 1
    
            
            End If
            
        
        Next i
        
        
        For i = 2 To ws.Range("K1").End(xlDown).Row
        
            If ws.Cells(i, 11).Value > ws.Cells(2, 17).Value Then
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                    ws.Cells(2, 17).Value = FormatPercent(ws.Cells(i, 11).Value)
        
            End If
                
                
            If ws.Cells(i, 11).Value < ws.Cells(3, 17).Value Then
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                    ws.Cells(3, 17).Value = FormatPercent(ws.Cells(i, 11).Value)
            End If
            
            
            If ws.Cells(i, 12).Value > ws.Cells(4, 17).Value Then
                    ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
            End If
            
    
        Next i
        
    
    

    Next ws
End Sub

