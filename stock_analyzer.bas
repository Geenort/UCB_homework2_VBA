Attribute VB_Name = "Module1"
Sub stockAnalyzer()
For Each ws In Worksheets
    Dim lastRow As Long
    Dim lastSummaryRow As Long
    Dim nextTickerPos As Integer
    Dim volume As Double
    
    Dim yearOpen As Single
    Dim yearClose As Single
    Dim yearChange As Single
    Dim yearPercentChange As Single
    
    Dim rowOfGInc As Integer
    Dim rowOfGDec As Integer
    Dim rowOfGTVolume As Integer
    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    nextTickerPos = 1
    volume = 0
    yearOpen = ws.Range("C2")
    
    'Create headers for all new columns
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    For i = 2 To lastRow
        'If next ticker is same, just add volume to total volume
        If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
            volume = volume + ws.Range("G" & i)
            
        'If next row is new ticker, add current ticker to new row in "Ticker" column.
        ElseIf ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
            nextTickerPos = nextTickerPos + 1
            ws.Range("I" & nextTickerPos) = ws.Cells(i, 1)
            
            'Still add volume to totaled volume, then write totaled volume to "Total Stock Volume" column, then reset volume to 0
            volume = volume + ws.Range("G" & i)
            ws.Range("L" & nextTickerPos) = volume
            volume = 0
            
            'Grab end of year close price. Calculate changes
            yearClose = ws.Range("F" & i)
            yearChange = yearClose - yearOpen
            'bypass overflow error, when denominator == 0
            If yearOpen <> 0 Then
                yearPercentChange = ((yearClose - yearOpen) / yearOpen)
            Else
                yearPercentChange = 0
            End If
            
            'Write to "Yearly Change" and "Percent Change" columns, and format
            ws.Range("J" & nextTickerPos) = yearChange
            If yearChange < 0 Then
                ws.Range("J" & nextTickerPos).Interior.Color = RGB(255, 0, 0)
            Else
                ws.Range("J" & nextTickerPos).Interior.Color = RGB(0, 255, 0)
            End If
            
            ws.Range("K" & nextTickerPos) = Format(yearPercentChange, "Percent")
            
            'Set yearOpen to first open price of next ticker, BUT not until after performing previous calculations
            yearOpen = ws.Range("C" & (i + 1))
        End If
    Next i
    
    'HARD solution
    'Find last row of new summary table
    lastSummaryRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    
    'Find rows of tickers in summary table with greatest % increase, greatest % decrease, and greatest total volume. Add 1 for header offset
    rowOfGInc = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastSummaryRow)), ws.Range("K2: K" & lastSummaryRow), 0) + 1
    rowOfGDec = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastSummaryRow)), ws.Range("K2: K" & lastSummaryRow), 0) + 1
    rowOfGTVolume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastSummaryRow)), ws.Range("L2: L" & lastSummaryRow), 0) + 1
    
    'Fill in "greatest" summary table using row values found above
    ws.Range("P2") = ws.Range("I" & rowOfGInc)
    ws.Range("Q2") = Format(ws.Range("K" & rowOfGInc), "Percent")
    
    ws.Range("P3") = ws.Range("I" & rowOfGDec)
    ws.Range("Q3") = Format(ws.Range("K" & rowOfGDec), "Percent")
    
    ws.Range("P4") = ws.Range("I" & rowOfGTVolume)
    ws.Range("Q4") = ws.Range("L" & rowOfGTVolume)
Next ws
End Sub
