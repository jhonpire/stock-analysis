Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQUO (Ticker: DQ)"
    
    'Create a Header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    rowstart = 2
    rowend = 3013
    totalVolume = 0
    
    Worksheets("2018").Activate
    For i = rowstart To rowend
        'Increase totalvolume if Ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
                    
        End If
    Next i
    
    MsgBox totalVolume
                
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    
End Sub