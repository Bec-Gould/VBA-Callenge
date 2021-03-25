Attribute VB_Name = "Module1"
Sub Challenge()
    
    'Set Ranges
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("P2").Value = WorksheetFunction.Max(Range("K:K"))
    Range("P3").Value = WorksheetFunction.Min(Range("K:K"))
    Range("P4").Value = WorksheetFunction.Max(Range("L:L"))

        'Last row
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row

        For q = 2 To lastRow

            If Cells(q, 11).Value = Cells(2, 16).Value Then
                Cells(2, 15).Value = Cells(q, 9).Value
                ElseIf Cells(q, 11).Value = Cells(3, 16).Value Then
                    Cells(3, 15).Value = Cells(q, 9).Value
                ElseIf Cells(q, 12).Value = Cells(4, 16).Value Then
                    Cells(4, 15).Value = Cells(q, 9).Value
            End If
        Next q

    Range("P2:P3").NumberFormat = "0.00%"

End Sub
