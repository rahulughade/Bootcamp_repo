Sub multiyear_macro()

'Declare variables - summary counter holds the cells number for summary table and total volume holds aggregate volume per ticker
Dim summarycounter As Double
Dim Totalvolume As Double
Dim stockopen As Double
Dim stockclose As Double
Dim yearlychange As Double
Dim percentchange As Double
    
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
        
'Initialize total volume to zero. Initialize summary counter to 2 to begin from row no. 2
    Totalvolume = 0
    summarycounter = 2
    
'Initialize summary table headers
    ws.Cells(1, 8).Value = "Ticker"
    ws.Cells(1, 9).Value = "Total Stock Volume"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    
'Assign value to stockopen variable
    stockopen = ws.Cells(2, 3).Value

'Last row function counts total no. of rows in all spreadsheets
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
'Loop through all rows
        For i = 2 To lastrow
            
            'Condition to compare ticker value in current and next cells
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Add ticker name to H2 cell
                ws.Range("H" & summarycounter).Value = ws.Cells(i, 1).Value
                
                'Aggregate total volume and store in I2 cell
                Totalvolume = Totalvolume + ws.Cells(i, 7).Value
                ws.Range("I" & summarycounter).Value = Totalvolume
                
                'Reset totalvolume to zero for next cycle
                Totalvolume = 0
                
                'Assign stockclose value
                stockclose = ws.Cells(i, 6).Value
                
                'Enter yearly change in Ji cell
                ws.Range("J" & summarycounter).Value = stockclose - stockopen
                yearlychange = ws.Range("J" & summarycounter).Value
                
                'Enter percent change in Ki cell
                If stockopen > 0 Then
                    ws.Range("K" & summarycounter).Value = yearlychange / stockopen
                    ws.Range("K" & summarycounter).NumberFormat = "0.00%"
                    percentchange = ws.Range("K" & summarycounter).Value
                
                    'Conditional color formatting'
                    If percentchange > 0 Then
                        ws.Range("J" & summarycounter).Interior.ColorIndex = 4
                        
                        Else
                            ws.Range("J" & summarycounter).Interior.ColorIndex = 3
                    End If
                    
                Else
                    percentchange = 0
                    
                End If
                
                'Reassign stock open value
                stockopen = ws.Cells(i + 1, 3).Value
                
                'Increment summary counter by 1
                summarycounter = summarycounter + 1
            
            Else
            'If ticker value is same, then add to total volume
                Totalvolume = Totalvolume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
    
    Next ws
    
End Sub
