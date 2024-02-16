Attribute VB_Name = "Module1"
Sub StockTickerLoop()

    For Each ws In Worksheets

    Dim ticker As String
    Dim vol As Double
    Dim Summary_Table_Row As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percentage_change As Double
    
        
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Vol"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
     
    Summary_Table_Row = 2
    vol = 0
    
    For i = 2 To 797712
        If year_open = 0 Then
            year_open = ws.Cells(i, 3).Value
        End If
        
        If ws.Cells(i - 1, 1) = ws.Cells(i, 1) And ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            year_close = ws.Cells(i, 6).Value
            yearly_change = year_close - year_open
            ticker = ws.Cells(i, 1).Value
            vol = vol + ws.Cells(i, 7).Value
            percentage_change = Round((yearly_change / year_open) * 100, 2)
                                    
            ws.Range("I" & Summary_Table_Row).Value = ticker
            ws.Range("J" & Summary_Table_Row).Value = yearly_change
            ws.Range("K" & Summary_Table_Row).Value = percentage_change
            ws.Range("L" & Summary_Table_Row).Value = vol
                                          
            Summary_Table_Row = Summary_Table_Row + 1
            vol = 0
        Else
            vol = vol + ws.Cells(i, 7).Value
        
        End If
        
        If yearly_change > 0 Then
           ws.Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 4
        Else
           ws.Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 3
        End If
        
        
    Next i
        
        
        
        
        ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K:K"))
        ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K:K"))
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L:L"))
        'Range("P2") = Cells(WorksheetFunction.Max(Range("K:K")), 9)
        'Range("P2") = WorksheetFunction.Max(Range("K:K").Offset(, 2))
        'Range("P3") = Range("I:I" + WorksheetFunction.Max(Range("K:K")))
        'Range("P4") = Range("I:I" + WorksheetFunction.Max(Range("L:L")))
           
    Next ws
End Sub

