# VBA-challenge
For Data Analytics Bootcamp Module 2 Challenge

'Most code was from TA Drew Hoang's speedrun of this assignment, with some changes in variable names. Also has more comments for personal understanding of each of the functions. -John T. Nguyen

```
Sub stocks()

    'labeling plenty of variables
    Dim StockVolume As Double
    Dim RowIndex As Long
    Dim Change As Double
    Dim ColumnIndex As Integer
    Dim Start As Long
    Dim RowCount As Long
    Dim PercentChange As Double
    Dim Days As Integer
    Dim DailyChange As Single
    Dim AverageChange As Double
    Dim ws As Worksheet
        
    'setting for loop to each worksheet
    For Each ws In Worksheets
        ColumnIndex = 0
        StockVolume = 0
        Change = 0
        Start = 2
        DailyChange = 0
        
        'first table's header row for ticker information
        ws.Range("k1").Value = "Ticker"
        ws.Range("l1").Value = "Yearly Change"
        ws.Range("m1").Value = "Percent Change"
        ws.Range("n1").Value = "Total Stock Volume"
        
        'second table's header row/column for tickers with notable stats
        ws.Range("s1").Value = "Ticker"
        ws.Range("t1").Value = "Value"
        ws.Range("r2").Value = "Greatest % Increase"
        ws.Range("r3").Value = "Greatest % Decrease"
        ws.Range("r4").Value = "Greatest Total Volume"
        
        'get row number of last row with data (accounts for sheets having different # of rows)
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For RowIndex = 2 To RowCount
            
            'adds new ticker names to ticker column in first table
            If ws.Cells(RowIndex + 1, 1).Value <> ws.Cells(RowIndex, 1).Value Then
            
                StockVolume = StockVolume + ws.Cells(RowIndex, 7).Value
                
                'formats row in first table for given ticker & corresponding values
                If StockVolume = 0 Then
                
                    ws.Range("K" & 2 + ColumnIndex).Value = Cells(RowIndex, 1).Value
                    ws.Range("L" & 2 + ColumnIndex).Value = 0
                    ws.Range("M" & 2 + ColumnIndex).Value = "%" & 0
                    ws.Range("N" & 2 + ColumnIndex).Value = 0
                    
                    
                Else
                    'Start value can change if a stock's opening value in a given row is 0
                    If ws.Cells(Start, 3) = 0 Then
                    
                        For find_value = Start To RowIndex
                            
                            If ws.Cells(find_value, 3).Value <> 0 Then
                            
                                Start = find_value
                                
                                Exit For
                                
                            End If
                            
                        Next find_value
                        
                    End If
                    
                    'calculations for first table
                    Change = (ws.Cells(RowIndex, 6) - ws.Cells(Start, 3))
                    PercentChange = Change / ws.Cells(Start, 3)
                    
                    Start = RowIndex + 1
                    
                    ws.Range("K" & 2 + ColumnIndex) = ws.Cells(RowIndex, 1).Value
                    ws.Range("L" & 2 + ColumnIndex) = Change
                    ws.Range("L" & 2 + ColumnIndex).NumberFormat = "0.00"
                    ws.Range("M" & 2 + ColumnIndex).Value = PercentChange
                    ws.Range("M" & 2 + ColumnIndex).NumberFormat = "0.00%"
                    ws.Range("N" & 2 + ColumnIndex).Value = StockVolume
                    
                    'Conditional formating for Change in first table (green in 1st case, red in 2nd, white in 3rd)
                    Select Case Change
                        Case Is > 0
                            ws.Range("L" & 2 + ColumnIndex).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("L" & 2 + ColumnIndex).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("L" & 2 + ColumnIndex).Interior.ColorIndex = 0
                    End Select
                
                End If
                    'resets all values except column index, moves down one row in same column
                    StockVolume = 0
                    Change = 0
                    ColumnIndex = ColumnIndex + 1
                    Days = 0
                    DailyChange = 0
                    
            Else
                'if ticker is the same, add to volume
                StockVolume = StockVolume + ws.Cells(RowIndex, 7).Value
            
            End If
        
        Next RowIndex
        
        'take the max and min values and respective tickers to place them in 2nd table
        ws.Range("t2") = "%" & WorksheetFunction.Max(ws.Range("M2:M" & RowCount)) * 100
        ws.Range("t3") = "%" & WorksheetFunction.Min(ws.Range("M2:M" & RowCount)) * 100
        ws.Range("t4") = WorksheetFunction.Max(ws.Range("N2:N" & RowCount))
        
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("M2:M" & RowCount)), ws.Range("M2:M" & RowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("M2:M" & RowCount)), ws.Range("M2:M" & RowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("N2:N" & RowCount)), ws.Range("N2:N" & RowCount), 0)
        
        ws.Range("s2") = ws.Cells(increase_number + 1, 11)
        ws.Range("s3") = ws.Cells(decrease_number + 1, 11)
        ws.Range("s4") = ws.Cells(volume_number + 1, 11)
        
    Next ws
        
End Sub
```
