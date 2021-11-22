Sub MacroCheck()


    Dim testMessage As String
    
    testMessage = "Hello World!"
    
    MsgBox (testMessage)
    
    
End Sub
Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQUO (Ticker: DQ)"
    
    'Create a Header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    Worksheets("2018").Activate
    
    rowstart = 2
    'DELETE rowend = 3013
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    totalVolume = 0
    
    Dim startingPrice As Double
    Dim endingPrice As Double
        
    For i = rowstart To rowEnd
        'Increase totalvolume if Ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
                    
        End If
        
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then

            startingPrice = Cells(i, 6).Value

        End If
        
        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            endingPrice = Cells(i, 6).Value

        End If
        
    Next i
    
    MsgBox totalVolume
                
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = endingPrice / startingPrice - 1
    
End Sub
Sub AllStockAnalysis()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the Analysis on?")

        startTime = Timer

    '1) Format the output sheet on All Stocks Analysis worksheet
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
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
    Sheets(yearValue).Activate
   
   '3c) Get the number of rows to loop over
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '4) Loop through tickers
    For i = 0 To 11
        
        Ticker = tickers(i)
        totalVolume = 0
        
        '5) loop through rows in the data
        
        Sheets(yearValue).Activate
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
    
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub
Sub formatAllStocksAnalysisTable()

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Font.Italic = True
    Range("A3:C3").Font.Size = 14
    Range("A3:C3").Font.Color = _
 RGB(100, 150, 150)
    Range("B4:B15").NumberFormat = "$#,##.00"
    Range("C4:C15").NumberFormat = "0.00%"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i
    
End Sub
Sub ClearWorksheet()

    Cells.Clear
    
End Sub
Sub yearValueAnalysis()

    yearValue = InputBox("What year would you like to run the Analysis on?")
    

End Sub
