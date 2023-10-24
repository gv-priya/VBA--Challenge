Attribute VB_Name = "Module4"
Private Sub Stockcalc()

    Dim ticker As String
    Dim vol As Double
    vol = 0
    Dim ws As Worksheet
    Dim Summary_Table_Row As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim Rows_num As Long
    Dim Column As Double
    Column = 1
    Dim Yearlych_lastrow As Double
    Dim k As Double
    Dim z As Double
    Cells(1, 9).Value = "ticker"
    Cells(1, 10).Value = "Yearly_change"
    Cells(1, 12).Value = "Total Stock Vol"
    Cells(1, 11).Value = "Yearly_percentage"
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells(2, Column + 13) = "Greatest % Increase"
        ws.Cells(3, Column + 13) = "Greatest % Decrease"
        ws.Cells(4, Column + 13) = "Greatest Volume"
        ws.Cells(1, Column + 14) = "Ticker"
        ws.Cells(1, Column + 15) = "Value"
        Rows_num = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        For i = 2 To Rows_num

            If year_open = 0 Then

                year_open = Cells(i, 3).Value
            End If
            'get unique ticker
            If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                year_close = Cells(i, 6).Value
                yearly_change = year_close - year_open
                year_percent = yearly_change / year_open
                ticker = Cells(i, 1).Value
                vol = vol + Cells(i, 7).Value
                Range("j" & Summary_Table_Row).Value = yearly_change
                Range("I" & Summary_Table_Row).Value = ticker
                Range("K" & Summary_Table_Row).Value = year_percent
                Range("L" & Summary_Table_Row).Value = vol
                Summary_Table_Row = Summary_Table_Row + 1
                vol = 0
            Else
                vol = vol + Cells(i, 7).Value
            End If
        Next i
        'conditional formatting
              
        Yearlych_lastrow = ws.Cells(Rows.Count, Column + 9).End(xlUp).Row
            For k = 2 To Yearlych_lastrow
                If (Cells(k, Column + 9).Value > 0 Or Cells(k, Column + 9).Value = 0) Then
                    Cells(k, Column + 9).Interior.ColorIndex = 10
                ElseIf Cells(k, Column + 9).Value < 0 Then
                    Cells(k, Column + 9).Interior.ColorIndex = 3
                End If
            Next k
        'Bonus question: greatest of yearly change and volume
         For z = 2 To Yearlych_lastrow
            If Cells(z, Column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Yearlych_lastrow)) Then
                Cells(2, Column + 14).Value = Cells(z, Column + 8).Value
                Cells(2, Column + 15).Value = Cells(z, Column + 11).Value
                Cells(2, Column + 15).NumberFormat = "0.00%"
            ElseIf Cells(z, Column + 10).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Yearlych_lastrow)) Then
                Cells(3, Column + 14).Value = Cells(z, Column + 8).Value
                Cells(3, Column + 15).Value = Cells(z, Column + 11).Value
                Cells(3, Column + 15).NumberFormat = "0.00%"
            ElseIf Cells(z, Column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Yearlych_lastrow)) Then
                Cells(4, Column + 14).Value = Cells(z, Column + 8).Value
                Cells(4, Column + 15).Value = Cells(z, Column + 11).Value
            End If
        Next z
    Next ws
End Sub
