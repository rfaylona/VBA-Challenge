Sub StockMarket():

' loop through all sheets
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    
        'Set variables
        Dim stockVol As Double
        stockVol = 0
        Dim yrOpen As Double
        Dim yrClose As Double
        Dim yrChange As Double
        Dim pctChange As Double
        Dim lastRow As Long
        Dim ticker As String
        Dim summaryTableRow As Long
        summaryTableRow = 2
    
        'set last row of data
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
 
        'set summary headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
    
        'set open price
        yrOpen = Cells(2, 3).Value
    
        'loop through each of the row of the data
        For i = 2 To lastRow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'create summary table
            ticker = Cells(i, 1).Value
            'set close price
            yrClose = Cells(i, 6).Value
            ' set year change
            yrChange = yrClose - yrOpen
            stockVol = stockVol + Cells(i, 7).Value
            ' set percent change
                If (yrOpen = 0 And yrClose = 0) Then
                    pctChange = 0
                ElseIf (yrOpen = 0 And yrClose <> 0) Then
                    pctChange = 1
                Else
                    pctChange = yrChange / yrOpen
                End If
            
            'write summary table
            Range("I" & summaryTableRow).Value = ticker
            Range("L" & summaryTableRow).Value = stockVol
            Range("J" & summaryTableRow).Value = yrChange
            Range("K" & summaryTableRow).Value = pctChange
            Range("K" & summaryTableRow).NumberFormat = "0.00%"
            summaryTableRow = summaryTableRow + 1
    
            'reset total
            stockVol = 0
            yrOpen = Cells(2, 3)
        
            Else
            'add to total
            stockVol = stockVol + Cells(i, 7).Value
    
        End If
    Next i
    
    'find last row of yearly change per WS
    yrcLastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    ' loop through each of the coloum data
    For j = 2 To yrcLastRow
        
        ' set colors
        If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
            Cells(j, 10).Interior.ColorIndex = 4
        ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
        End If
    Next j
    
    ' write 2nd summary table
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Result"
    
    ' find find tickers with greatest % value/volume and least % value
    For m = 2 To yrcLastRow
        
        If Cells(m, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & yrcLastRow)) Then
            Cells(2, 16).Value = Cells(m, 9).Value
            Cells(2, 17).Value = Cells(m, 11).Value
            Cells(2, 17).NumberFormat = "0.00%"
        ElseIf Cells(m, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & yrcLastRow)) Then
                Cells(3, 16).Value = Cells(m, 9).Value
                Cells(3, 17).Value = Cells(m, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
        ElseIf Cells(m, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & yrcLastRow)) Then
                Cells(4, 16).Value = Cells(m, 9).Value
                Cells(4, 17).Value = Cells(m, 12).Value
        End If
    Next m
    
    Next WS
       
End Sub