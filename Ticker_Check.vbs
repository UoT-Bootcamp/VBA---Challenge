
Sub ticker_check()
    
    'Loop through all sheets
    For Each ws In Worksheets
    
        'declare variables for column and set the value to 1 which means we will check the value in column 1
        Dim column As Integer
        column = 1
    
        'Declare variable and determine the last row of each sheet
        Dim RowEnd As Long
        RowEnd = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'set an initial variable for holding the first row position
        Dim row_position As Long
        row_position = 2
    
        'Declare variables for close price, open price, yearly change and percent change
        Dim close_price As Double
        Dim open_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        
        'set an initial variable for holding the first row position of the open price
        Dim start As Long
        start = 2
        
        'set an initial variable for holding the total of volume and sets the value to 0
        Dim volume_total As Double
        volume_total = 0
   
        'Add headers to the selected cells
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        
        'loop through each row of sheet
        For i = 2 To RowEnd
                
                'searches for when the value of the next cell is different than that of the current cell, if it is not
                If ws.Cells(i, column).Value <> ws.Cells(i + 1, column).Value Then
                        
                    ' Set the value of the ticker
                    ws.Cells(row_position, "J").Value = ws.Cells(i, column).Value
                        
                    'set the value of the  close price
                    close_price = ws.Cells(i, "F").Value
                        
                    'set the value of the open price
                    open_price = ws.Cells(start, 3).Value
                        
                    'calculate the value of yearly_change
                    yearly_change = close_price - open_price
                        
                    'set the column with the yearly change value
                    ws.Range("K" & row_position).Value = Round(yearly_change, 2)
            
                    'searches for close price not equals to 0, if it is not then
                    If open_price <> 0 Then
                                
                                'calculate the percent change and round the output upto 2 decimal places
                                percent_change = Round(((yearly_change / open_price) * 100), 2)
                
                    End If
                        
                    'set the value of selected cells with percent change and format it with %
                    ws.Range("L" & row_position).Value = percent_change & "%"
            
                   'searches for yearly change not equals to 0, if it is then
                    If yearly_change > 0 Then
                        
                            'fill the selected cells' interior to green
                            ws.Range("K" & row_position).Interior.ColorIndex = 4
                    Else
                        
                            'fill the selected cells' interior to red
                            ws.Range("K" & row_position).Interior.ColorIndex = 3
                        
                    End If
                        
                    'searches for yearly change not equals to 0, if it is then
                    If percent_change > 0 Then
                        
                        'fill the selected cells' interior to green
                        ws.Range("L" & row_position).Interior.ColorIndex = 4
                    
                    Else
                        
                        'fill the selected cells' interior to green
                         ws.Range("L" & row_position).Interior.ColorIndex = 3
                       
                    End If
                        
                        
                    'calculate the total of volume
                    volume_total = volume_total + ws.Cells(i, 7).Value
                        
                    'set the selected cell with the volume total value
                    ws.Range("M" & row_position).Value = volume_total
                
                    'Update the row position
                    row_position = row_position + 1
                        
                    'reset the value of total of volume to 0 for executing the next loop
                    volume_total = 0
               
                    'update the value of start for the next loop to find the row number for next open price
                    start = i + 1
        
        
        
                'If the cell immediately following a row is the same ...
                Else
                        
                    'Add to the volume total
                    volume_total = volume_total + ws.Cells(i, 7).Value
            
                End If
    
        Next i
   
        'Autofil the columns
        ws.Columns("J:M").AutoFit
        
        'CALCULATION FOR "GREATEST % INCREASE", "GREATEST % DECREASE" & "GREATEST TOTAL VOLUME"

        'set the headers of the selected cells
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
    
        'set the variables for calculating the "Greatest % Increase", "Greatest % Decrease" and "Greatest Total Volume"
        Dim max_percent As Double
        Dim min_percent As Double
        Dim max_volume As Double
    
       'create the pointer for the range in which we will look for the value
        Set Data = ws.Range("L:L")
        Set Volume = ws.Range("M:M")
        
        'Find the "Greatest % Increase"
        max_percent = Application.WorksheetFunction.Max(Data)
        
        'set the greatest % increase value to the selected cell and round it to 2 decimal place with a % sign
        ws.Range("R2").Value = max_percent
        ws.Range("R2").NumberFormat = "0.00%"
    
        'Find the "Greatest % Decrease"
        min_percent = Application.WorksheetFunction.Min(Data)
        
        'set the greatest % decrease value to the selected cell and round it to 2 decimal place with a % sign
        ws.Range("R3").Value = min_percent
        ws.Range("R3").NumberFormat = "0.00%"
    
        'Find the "Greatest Total Volume"
        max_volume = Application.WorksheetFunction.Max(Volume)
       
       'set the greatest volume to the selected cell
        ws.Range("R4").Value = max_volume
    
        'Loop through all rows
        For i = 2 To RowEnd
        
            'check in "percent change" column for the value equals to the greatest % increase, if it is then
            If ws.Cells(i, 12).Value = max_percent Then
            
                'set the value of selected cell with the value of the corresponding ticker symbol
                ws.Range("Q2").Value = ws.Cells(i, 10).Value
                        
            'check in "percent change" column for the value equals to the greatest % decrease, if it is then
            ElseIf ws.Cells(i, 12).Value = min_percent Then
                
                'set the value of selected cell with the value of the corresponding ticker symbol
                ws.Range("Q3").Value = ws.Cells(i, 10).Value
                        
            'check in "percent change" column for the value equals to the greatest % decrease, if it is then
            ElseIf ws.Cells(i, 13).Value = max_volume Then
                    
                'set the value of selected cell with the value of the corresponding ticker symbol
                ws.Range("Q4").Value = ws.Cells(i, 10).Value
            
            End If
            
        Next i
    
        'Autofit the columns
        ws.Columns("P:R").AutoFit
    
    Next ws
    
End Sub






