Sub Stock_Data()
    Dim ws As Worksheet
    Dim Last_Row_First As Long
    Dim i As Long
    Dim j As Long
    Dim Ticker_Row As Long
    Dim Percent_Change As Double
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As Double
   
    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        Ticker_Row = 2
        j = 2
        Last_Row_First = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To Last_Row_First
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(Ticker_Row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(Ticker_Row, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                If ws.Cells(Ticker_Row, 10).Value < 0 Then
                    ws.Cells(Ticker_Row, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(Ticker_Row, 10).Interior.ColorIndex = 4
                End If
                If ws.Cells(j, 3).Value <> 0 Then
                    Percent_Change = (ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value
                    ws.Cells(Ticker_Row, 11).Value = Format(Percent_Change, "Percent")
                Else
                    ws.Cells(Ticker_Row, 11).Value = Format(0, "Percent")
                End If
                ws.Cells(Ticker_Row, 12).Value = WorksheetFunction.Sum(ws.Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                Ticker_Row = Ticker_Row + 1
                j = i + 1
            End If
        Next i
        Greatest_Increase = ws.Cells(2, 11).Value
        Greatest_Decrease = ws.Cells(2, 11).Value
        Greatest_Volume = ws.Cells(2, 12).Value
        Last_Row_Last = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For i = 2 To Last_Row_Last
            If ws.Cells(i, 11).Value > Greatest_Increase Then
                Greatest_Increase = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 11).Value < Greatest_Decrease Then
                Greatest_Decrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 12).Value > Greatest_Volume Then
                Greatest_Volume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
        Next i
        ws.Cells(2, 17).Value = Format(Greatest_Increase, "Percent")
        ws.Cells(3, 17).Value = Format(Greatest_Decrease, "Percent")
        ws.Cells(4, 17).Value = Format(Greatest_Volume, "Scientific")
    Next ws
End Sub
