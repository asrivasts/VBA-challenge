Sub stockAnalyzer()
    
    For Each ws In ActiveWorkbook.Worksheets
    
        'Define variables that will be used during the calculations
        
        'Tracks to see if we are looking at a new Stock
        Dim newStock As Integer
        
        'Stores the current ticker symbol
        Dim ticker As String
        
        'Stores the current ticker volume
        Dim stockVolume As LongLong
        
        'Stores teh current ticker year open and close prices, as well as defines two variables that will be used in calculation
        Dim yearOpen, yearClose, yearlyChange, percentChange As Double
        
        'Calculates the last row to close out our for loop
        Dim lastRow As Long
        
        'Sets the current row you are writing data in
        Dim printRow As Integer
        
        'initialize my variables
        newStock = 0
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        printRow = 2
        
        'Clears all contents and formats in the printing range.  This is just in case I do multiple runs, and made any mistakes
        ws.Range("J:R").Clear
        
        'Set the headings of the table for our calculations
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        ws.Range("J1:M1").Font.Bold = True
        
        
        For i = 2 To lastRow
            
            'Take the ticker, year open, and stock volume of the new stock
            If (newStock = 0) Then
                ticker = ws.Cells(i, 1).Value
                yearOpen = ws.Cells(i, 3).Value
                stockVolume = ws.Cells(i, 7).Value
                newStock = 1
            Else
                'If its not a new stock, just increment the volume of the stock
                stockVolume = stockVolume + ws.Cells(i, 7).Value
            End If
            
            'if the next row is a new ticker, then complete the calculations and print
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                yearClose = ws.Cells(i, 6).Value
                yearlyChange = yearClose - yearOpen
                If (yearOpen <> 0) Then
                    percentChange = CDbl(yearlyChange / yearOpen)
                Else
                    percentChange = 0
                End If
                
                'Print data and format cells
                ws.Cells(printRow, 10).Value = ticker
                ws.Cells(printRow, 11).Value = yearlyChange
                ws.Cells(printRow, 11).NumberFormat = "$#,##0.00"
                ws.Cells(printRow, 12).Value = percentChange
                ws.Cells(printRow, 12).NumberFormat = "0.00%"
                ws.Cells(printRow, 13).Value = stockVolume
                ws.Cells(printRow, 13).NumberFormat = "#,##0"
                
                'if the change is positive, make the cell green
                'if the change is negative, make the cell red
                'or else, leave the cell color transparent
                If (yearlyChange > 0) Then
                    ws.Cells(printRow, 11).Interior.ColorIndex = 35
                ElseIf (yearlyChange < 0) Then
                    ws.Cells(printRow, 11).Interior.ColorIndex = 38
                End If
                
                'reset newStock to start calculating the new stock's data
                newStock = 0
    
                'increment the current print row
                printRow = printRow + 1
                
            End If
        Next i
        
        
        'Set the headings of the table for our calculations
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        ws.Range("P2:P4").Font.Bold = True
        ws.Range("Q1:R1").Font.Bold = True
        
        're-calculate the last row based on independent tickers
        lastRow = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
        
        'new variabels for calculations
        Dim greatInc, greatDec As Double
        Dim tikInc, tikDec, tikVol As String
        Dim greatVol As LongLong
        
        'reset values
        greatInc = 0
        greatDec = 0
        greatVol = 0
            
        For i = 2 To lastRow
    
            'If the positive change is greater than the current max, store the new ticker and % change value
            If (ws.Cells(i, 12).Value > greatInc) Then
                greatInc = ws.Cells(i, 12).Value
                tikInc = ws.Cells(i, 10).Value
            End If
            
            'If the negative change is greater than the current max, store the new ticker and % change value
            If (ws.Cells(i, 12).Value < greatDec) Then
                greatDec = ws.Cells(i, 12).Value
                tikDec = ws.Cells(i, 10).Value
            End If
            
            'If the volume is greater than the current max, store the new ticker and volume value
            If (ws.Cells(i, 13).Value > greatVol) Then
                greatVol = ws.Cells(i, 13).Value
                tikVol = ws.Cells(i, 10).Value
            End If
            
        Next i
        
        'Print the results of our analysis
        ws.Range("Q2").Value = tikInc
        ws.Range("Q3").Value = tikDec
        ws.Range("Q4").Value = tikVol
        ws.Range("R2").Value = greatInc
        ws.Range("R2").NumberFormat = "0.00%"
        ws.Range("R3").Value = greatDec
        ws.Range("R3").NumberFormat = "0.00%"
        ws.Range("R4").Value = greatVol
        ws.Range("R4").NumberFormat = "#,##0"
    Next
End Sub

