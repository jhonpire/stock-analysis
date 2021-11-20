Sub AllStockAnalysis()
    '1) Format the output sheet on All Stocks Analysis worksheet
    Range("A1").Value = "All Stocks (2018)"
    
    'Create Header Row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    '2) Initialize array of all tickers
    Dim tickers(11) As String
    
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
        
        
    '3a) Initialize variables for starting price and ending price
    Dim startingPrice As Double
    Dim endingPrice As Double

   '3b) Activate data worksheet
    Worksheets("2018").Activate
   
   '3c) Get the number of rows to loop over
    rowstart = 2
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    totalVolume = 0
    
    '4) Loop through tickers
    For i = 0 To 11

        '5) loop through rows in the data
        Ticker = tickers(i)
        For j = 2 To RowCount
        
            '5a) Get total volume for current ticker
            If Cells(j, 1).Value = Ticker Then
            
                totalVolume = totalVolume + Cells(j, 8).Value

            End If
            '5b) get starting price for current ticker
            If Cells(j - 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then

                startingPrice = Cells(j, 6).Value
            End If
            '5c) get ending price for current ticker
            If Cells(j + 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then

                endingPrice = Cells(j, 6).Value
            End If

        Next j
     
       'Output results
       Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = Ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
                   
    Next i
    
End Sub
