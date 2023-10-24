Attribute VB_Name = "Module2"
Sub dateformatter()
    Dim WS As Worksheet
    Dim lastRow_num As Long
    'Dim Rnge As Range
    Dim Cell As Range
    Dim str As String
    Dim i As Long, j As Long
           
    For Each WS In ThisWorkbook.Worksheets
        lastRow_num = WS.Cells(WS.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow_num 'Loop through each row of data starting from row 2
            WS.Cells(i, 2).Value = DateSerial(Left(WS.Cells(i, 2).Value, 2), Mid(WS.Cells(i, 2).Value, 5, 2), Right(WS.Cells(i, 2).Value, 2))
            WS.Cells(i, 2).NumberFormat = "dd/mm/yyyy" 'Change the number format to your desired format
        Next i
        
    Next WS
    
    
End Sub
