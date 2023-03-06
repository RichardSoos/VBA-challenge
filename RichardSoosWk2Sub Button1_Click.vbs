Sub Button1_Click()
Dim h_pctchg As Double
Dim l_pctchg As Double
Dim xmax As Double
Dim sgrtstpctinc As Double
Dim sgrtstpctdec As Double
Dim r As Range
Dim s As Range
Dim h_ticker As String
Dim l_ticker As String
Dim NewTicTally As Integer
Dim SheetCount As Integer
Dim sheetrowcounting As Double
Dim TotStockVol As Double
Dim Ticker As String
Dim OpnPrcSecFrstofYr As Double
Dim ClsPrcSecLstofYr As Double
Dim YearlyChange As Long
Dim PctChange As Long
SheetCount = Application.Worksheets.Count
For i = 1 To SheetCount
    Worksheets(i).Activate
    'Add Column Headers I Ticker, J Yearly Change, K Percent Change L Total Stock Volume
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 13).Value = "Greatest % Increase"
    Cells(3, 13).Value = "Greatest % Decrease"
   ' soos Cells(4, 13).Value = "Greatest Total Volume"
    sheetrowcounting = Range("A" & Rows.Count).End(xlUp).Row
    'Find first ticker code in list
    Ticker = Cells(2, 1).Value
    Range("I2").Value = Ticker
    TotStockVol = Cells(2, 7).Value
    NewTicTally = 2
    OpnPrcSecFrstofYr = Cells(2, 3).Value
    ClsPrcSecLstofYr = Cells(2, 6).Value
   
   
    h_pctchg = 0
    l_pctchg = 0
    h_ticker = Ticker
    
    For j = 3 To sheetrowcounting
        
        If Cells(j, 1).Value <> Ticker Then 'Then we have a new ticker, print previous totstockval
        Cells(NewTicTally + 1, 9).Value = Cells(j, 1).Value ' Get the old ticker symbol and print it
         Cells(NewTicTally, 12).Value = TotStockVol
         'TotStockVol = 0  'reset the TotStockVol to zero soos
        TotStockVol = Cells(j, 7).Value
         
            
            ClsPrcSecLstofYr = Cells(j - 1, 6).Value
         
        YearlyChange = 100 * (ClsPrcSecLstofYr - OpnPrcSecFrstofYr)
        
        Cells(NewTicTally, 10).Value = YearlyChange / 100
        
        
        PctChange = 100 * (YearlyChange / OpnPrcSecFrstofYr)
        
        
        If PctChange > h_pctchg Then
           h_pctchg = PctChange
           h_ticker = Ticker
        End If
        
          If PctChange < l_pctchg Then
           l_pctchg = PctChange
           l_ticker = Ticker
        End If
        
        
        Cells(NewTicTally, 11).Value = PctChange / 10000
        'Now set the cell colour of yrly change, +bve being green, -ve being red
           
          If YearlyChange > 0 Then
             Cells(NewTicTally, 10).Interior.ColorIndex = 4
          ElseIf YearlyChange < 0 Then
             Cells(NewTicTally, 10).Interior.ColorIndex = 3
          End If
          
           
           
         OpnPrcSecFrstofYr = Cells(j, 3)
         
         NewTicTally = NewTicTally + 1
        Ticker = Cells(j, 1).Value

        Else
        TotStockVol = TotStockVol + Cells(j, 7).Value
       
         ClsPrcSecLstofYr = Cells(j, 6).Value
         
         
           If PctChange > h_pctchg Then
           h_pctchg = PctChange
           h_ticker = Ticker
           End If
        
           If PctChange < l_pctchg Then
           l_pctchg = PctChange
           l_ticker = Ticker
           End If
       
        End If
    
Next j
's
        YearlyChange = 100 * (ClsPrcSecLstofYr - OpnPrcSecFrstofYr)
        
        Cells(NewTicTally, 10).Value = YearlyChange / 100
        
        
        PctChange = 100 * (YearlyChange / OpnPrcSecFrstofYr)
        
        Cells(NewTicTally, 11).Value = PctChange / 10000
          If YearlyChange > 0 Then
             Cells(NewTicTally, 10).Interior.ColorIndex = 4
          ElseIf YearlyChange < 0 Then
             Cells(NewTicTally, 10).Interior.ColorIndex = 3
          End If
         
         Cells(NewTicTally, 12).Value = TotStockVol
         Range("N2").Value = h_ticker
         Range("N3").Value = l_ticker

's
 NewTicTally = 2

Set r = Range("G2:G" & Rows.Count)
xmax = Application.WorksheetFunction.Max(r)
' soos Cells(4, 15).Value = xmax

Set s = Range("K2:K" & Rows.Count)
sgrtstpctinc = Application.WorksheetFunction.Max(s)
sgrtstpctdec = Application.WorksheetFunction.Min(s)
Cells(2, 15).Value = sgrtstpctinc
Cells(3, 15).Value = sgrtstpctdec



    Next i
 
'Return to first worksheet to present to viewer
Worksheets("2018").Activate
Worksheets("2018").Cells(1, 1).Select

End Sub
