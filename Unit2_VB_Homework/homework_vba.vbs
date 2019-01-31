
Sub ComputeVolume()
    
    'doing hard version now....looping thru ALL the stocks on ALL the sheets and putting summary info for
    'each year on each sheet.
    
    Dim CumTickerVolume, TickerVolume, GreatestTotalVolume, lastRowIndex, TickerSummaryRow As Double
    Dim PercentChange, GreatestIncPerChange, GreatestDecPerChange  As Double
    Dim StartOfYearOpeningTickerPrice, EndOfYearClosingTickerPrice, YearlyChange As Double
    Dim CurrTicker, GreatestDecTicker, GreatestIncTicker, sheetName As String
    Dim WS_Count As Integer      'number of worksheets in our workbook
    
    Application.StatusBar = "Starting ticker colume computation..."
    
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    For ws = 1 To WS_Count
        
        'init these running values for every new sheet
        GreatestIncPerChange = 0
        GreatestDecPerChange = 0
        GreatestTotalVolume = 0
       
        sheetName = ActiveWorkbook.Worksheets(ws).Name    'start with the first sheet
        TickerSummaryRow = 2    'start at row 2 each time for each sheet
        
        'clear out Cols I:Q since I'll be putting summary data there...be sure it's clean
        Sheets(sheetName).Select
        Columns("I:Q").ClearContents
        Columns("I:Q").ClearFormats
        
        'make summary headers
        Range("I1").Value = "Ticker"    'define header in THE FIRST SHEET sheet...THIS IS FOR ALL THE STOCKS
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        'select that sheet
        Sheets(sheetName).Select
        
        'find last row of data set
        lastRowIndex = Sheets(sheetName).Cells(Rows.Count, 1).End(xlUp).Row
        
        'do a sort first to guarantee that the tickers are sorted with similar groups clustered, so loop and inner loop work
        Range("A1:G" & lastRowIndex).Sort key1:=Range("A1:G" & lastRowIndex), order1:=xlAscending, Header:=xyYes
        
        CumTickerVolume = 0     'cumulative ticker volume
        
        StartOfYearOpeningTickerPrice = (Range("C2").Value)    'first starting price
            
        For i = 2 To lastRowIndex       'go thru each row in the sheet
            CurrTicker = Range("A" & i).Value
                NextTicker = Range("A" & i + 1).Value
                TickerVolume = (Range("G" & i).Value)  'we have a unique ticker....so get this volume
                    
                If CurrTicker <> NextTicker Then
                    'this is the end of the current ticker section, so put the ticker info on that sheet
                    CumTickerVolume = CumTickerVolume + TickerVolume    'add to our running total
                    Range("I" & TickerSummaryRow).Value = CurrTicker
                    Range("L" & TickerSummaryRow).Value = CumTickerVolume
                    EndOfYearClosingTickerPrice = (Range("F" & i).Value)
                    YearlyChange = EndOfYearClosingTickerPrice - StartOfYearOpeningTickerPrice
                    Range("J" & TickerSummaryRow).Value = YearlyChange
                    
                    If StartOfYearOpeningTickerPrice <> 0 Then   'if it stayed 0, then the whole year it had no trading price
                        PercentChange = YearlyChange / StartOfYearOpeningTickerPrice
                    Else
                        PercentChange = 0
                    End If
                    
                    'hard part:  keep the greatest negative and the greatest positive % change and its ticker
                    If PercentChange > 0 And PercentChange > GreatestIncPerChange Then
                        GreatestIncPerChange = PercentChange   'set this to new % increase one
                        GreatestIncTicker = CurrTicker
                    ElseIf PercentChange < 0 And PercentChange < GreatestDecPerChange Then
                        GreatestDecPerChange = PercentChange   'set this to new % decrease one
                        GreatestDecTicker = CurrTicker
                    End If
                    
                    'hard part:  keep the greatest total volume, ticker and value
                    If CumTickerVolume > 0 And CumTickerVolume > GreatestTotalVolume Then
                        GreatestTotalVolume = CumTickerVolume
                        GreateatTotVolTicker = CurrTicker
                    End If
                        
                    Range("K" & TickerSummaryRow).Value = PercentChange
                    Range("K" & TickerSummaryRow).NumberFormat = "0.00%"    'format it to %
                    
                    'conditional format + green, - red
                    If YearlyChange < 0 Then
                        'make the cell background color RED  (negative)      '3
                        Range("J" & TickerSummaryRow).Interior.ColorIndex = 3
                    ElseIf YearlyChange > 0 Then
                        'make the cell background color GREEN   (positive)     '4
                        Range("J" & TickerSummaryRow).Interior.ColorIndex = 4
                    End If
                    
                    TickerSummaryRow = TickerSummaryRow + 1
                    CumTickerVolume = 0      're-init
                    
                    'get next new tickers starting year price
                    StartOfYearOpeningTickerPrice = (Range("C" & i + 1).Value)
                    
                Else
                    'if in here....Curr Ticker still = next ticker
                    CumTickerVolume = CumTickerVolume + TickerVolume    'add to our running total
                    'below trick to replace a zero for StartOfYearOpeningTickerPrice if 0 from the logic above
                    'when the ticker changes and 01Jan was a 0
                    If StartOfYearOpeningTickerPrice = 0 Then    'will continue to check and replace with a non-zero if possible
                        StartOfYearOpeningTickerPrice = (Range("C" & i).Value)
                    Else
                    End If
                
                End If
        Next i
        
        'put year end summary % data for the year in the top right corner
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        'greatest % increase
        Range("O2").Value = "Greatest % Increase"
        Range("P2").Value = GreatestIncTicker       'ticker
        Range("Q2").Value = GreatestIncPerChange                'value
        Range("Q2").NumberFormat = "0.00%"    'format it to %
        
        'greatest % decrease
        Range("O3").Value = "Greatest % Decrease"
        Range("P3").Value = GreatestDecTicker       'ticker
        Range("Q3").Value = GreatestDecPerChange                'value
        Range("Q3").NumberFormat = "0.00%"  'format it to %
        
        'greatest total volume
        Range("O4").Value = "Greatest Total Volume"
        Range("P4").Value = GreateatTotVolTicker       'ticker
        Range("Q4").Value = GreatestTotalVolume                'value
        
        'do a column fit here for the cols I:L
        Columns("I:Q").EntireColumn.AutoFit
                    
        Application.StatusBar = "Finished looking at sheet: " & sheetName
        
        Range("A1").Select      'select A1 cell to show sheet from top
        
    Next ws
        
    Application.StatusBar = "Finished!"
    MsgBox "Finished!"

End Sub


