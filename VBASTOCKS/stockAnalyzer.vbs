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
    
End Sub
