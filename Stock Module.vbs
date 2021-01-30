Attribute VB_Name = "Module1"
Sub stocks()
    'create variables to hold all values
    Dim ticker As String
    Dim dates As Long
    Dim opening As Double
    Dim high As Double
    Dim low As Double
    Dim closing As Double
    Dim vol As Double
    Dim summary_table_row As Integer
    Dim close1 As Double
    
    
    
    Dim summary_table_ticker As String
    Dim Yearly_Change As Double
    Dim percent As Double
    Dim Total_Volume As Double
    
    
    'ticker = Cells(i, 1).Value
    'dates = Cells(i, 2).Value
    'opening = Cells(i, 3).Value
    'high = Cells(i, 4).Value
    'low = Cells(i, 5).Value
    'closing = Cells(i, 6).Value
    'vol = Cells(i, 7).Value
    
    
    For Each ws In Worksheets
    
    
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'lastcol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "percent"
        ws.Cells(1, 12).Value = "Total Volume"
        
        summary_table_row = 2
        

         
         'Loop through all data
            For i = 2 To lastrow
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    opening = ws.Cells(i, 3).Value
                    low = ws.Cells(i, 5).Value
                End If
                
            'If ticker = "A" Then
            'Cells(i, 9).Value = ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Set the ticker
                ticker = ws.Cells(i, 1).Value
                'print the ticker
                
                
                ws.Cells(summary_table_row, 9).Value = ticker
                close1 = ws.Cells(i, 6).Value ' previous ticker
                
                high = ws.Cells(i, 4).Value
                
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                Yearly_Change = (close1 - opening)
                ws.Cells(summary_table_row, 10).Value = Yearly_Change
                
                    If opening <> 0 Then
                        percent = ((close1 - opening) / opening)
                    End If
                    
                ws.Cells(summary_table_row, 11).Value = percent
                ws.Cells(summary_table_row, 11).Value = FormatPercent(ws.Cells(summary_table_row, 11).Value)
                ws.Cells(summary_table_row, 12).Value = Total_Volume
                    If percent < 0 Then
                        ws.Cells(summary_table_row, 11).Interior.ColorIndex = 3
                    Else
                        ws.Cells(summary_table_row, 11).Interior.ColorIndex = 4

                    End If
                
                
                summary_table_row = summary_table_row + 1
                Yearly_Change = 0
                percent = 0
                Total_Volume = 0
                
                
                
                opening = ws.Cells(i, 3).Value
                low = ws.Cells(i, 5).Value
                
            
            Else
            'calculate the yearly open-close
            
            
            
            
            'calculate the percent change high-low
            'percent = (high - low)
            'ws.Cells(summary_table_row, lastcol + 4).Value = percent
'
                
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
            
            
           End If
        
        Next i
    Next ws
    
End Sub


