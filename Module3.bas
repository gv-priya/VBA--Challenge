Attribute VB_Name = "Module3"
Sub conditional()
    Dim ws As Worksheet
    Dim Column As Double
    Dim j As Double
    Column = 1
    Dim Yearlych_lastrow As Double
    For Each ws In ThisWorkbook.Worksheets
        Yearlych_lastrow = ws.Cells(Rows.Count, Column + 9).End(xlUp).Row
            For j = 2 To Yearlych_lastrow
                If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                    Cells(j, Column + 9).Interior.ColorIndex = 10
                ElseIf Cells(j, Column + 9).Value < 0 Then
                    Cells(j, Column + 9).Interior.ColorIndex = 3
                End If
            Next j
    'Bonus question: greatest of yearly change and volume
    For z = 2 To Yearlych_lastrow
        If Cells(z, Column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Yearlych_lastrow)) Then
            Cells(2, Column + 15).Value = Cells(z, Column + 8).Value
            Cells(2, Column + 14).Value = Cells(z, Column + 12).Value
            Cells(2, Column + 14).NumberFormat = "0.00%"
        ElseIf Cells(z, Column + 10).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Yearlych_lastrow)) Then
            Cells(3, Column + 15).Value = Cells(z, Column + 8).Value
            Cells(3, Column + 14).Value = Cells(z, Column + 12).Value
            Cells(3, Column + 14).NumberFormat = "0.00%"
        ElseIf Cells(z, Column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Yearlych_lastrow)) Then
            Cells(4, Column + 15).Value = Cells(z, Column + 8).Value
            Cells(4, Column + 14).Value = Cells(z, Column + 11).Value
        End If
    Next z
    Next ws
End Sub
