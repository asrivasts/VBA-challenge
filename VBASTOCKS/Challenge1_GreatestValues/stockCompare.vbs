Sub stockAnalyzer()
    
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
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    printRow = 2
    
    'Clears all contents and formats in the printing range.  This is just in case I do multiple runs, and made any mistakes
    Range("J:R").Clear
    
    'Set the headings of the table for our calculations
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Stock Volume"
    Range("J1:M1").Font.Bold = True
    
    
    For i = 2 To lastRow
        
        'Take the ticker, year open, and stock volume of the new stock
        If (newStock = 0) Then
            ticker = Cells(i, 1).Value
            yearOpen = Cells(i, 3).Value
            stockVolume = Cells(i, 7).Value
            newStock = 1
        Else
            'If its not a new stock, just increment the volume of the stock
            stockVolume = stockVolume + Cells(i, 7).Value
        End If
        
        'if the next row is a new ticker, then complete the calculations and print
        If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
            yearClose = Cells(i, 6).Value
            yearlyChange = yearClose - yearOpen
            
            If (yearOpen <> 0) Then
                percentChange = CDbl(yearlyChange / yearOpen)
            Else
                percentChange = 0
            End If
            
            'Print data and format cells
            Cells(printRow, 10).Value = ticker
            Cells(printRow, 11).Value = yearlyChange
            Cells(printRow, 11).NumberFormat = "$#,##0.00"
            Cells(printRow, 12).Value = percentChange
            Cells(printRow, 12).NumberFormat = "0.00%"
            Cells(printRow, 13).Value = stockVolume
            Cells(printRow, 13).NumberFormat = "#,##0"
            
            'if the change is positive, make the cell green
            'if the change is negative, make the cell red
            'or else, leave the cell color transparent
            If (yearlyChange > 0) Then
                Cells(printRow, 11).Interior.ColorIndex = 35
            ElseIf (yearlyChange < 0) Then
                Cells(printRow, 11).Interior.ColorIndex = 38
            End If
            
            'reset newStock to start calculating the new stock's data
            newStock = 0

            'increment the current print row
            printRow = printRow + 1
            
        End If
    Next i
    
    
    'Set the headings of the table for our calculations
    Range("P2").Value = "Greatest % Increase"
    Range("P3").Value = "Greatest % Decrease"
    Range("P4").Value = "Greatest Total Volume"
    Range("Q1").Value = "Ticker"
    Range("R1").Value = "Value"
    Range("P2:P4").Font.Bold = True
    Range("Q1:R1").Font.Bold = True
    
    're-calculate the last row based on independent tickers
    lastRow = Cells(Rows.Count, 10).End(xlUp).Row
    
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
        If (Cells(i, 12).Value > greatInc) Then
            greatInc = Cells(i, 12).Value
            tikInc = Cells(i, 10).Value
        End If
        
        'If the negative change is greater than the current max, store the new ticker and % change value
        If (Cells(i, 12).Value < greatDec) Then
            greatDec = Cells(i, 12).Value
            tikDec = Cells(i, 10).Value
        End If
        
        'If the volume is greater than the current max, store the new ticker and volume value
        If (Cells(i, 13).Value > greatVol) Then
            greatVol = Cells(i, 13).Value
            tikVol = Cells(i, 10).Value
        End If
        
    Next i
    
    'Print the results of our analysis
    Range("Q2").Value = tikInc
    Range("Q3").Value = tikDec
    Range("Q4").Value = tikVol
    Range("R2").Value = greatInc
    Range("R2").NumberFormat = "0.00%"
    Range("R3").Value = greatDec
    Range("R3").NumberFormat = "0.00%"
    Range("R4").Value = greatVol
    Range("R4").NumberFormat = "#,##0"
End Sub
