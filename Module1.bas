Attribute VB_Name = "Module1"
Sub UniqueValues()
    Dim WS As Worksheet
    Dim lastrow As Long
    Dim arr() As Variant
    Dim dict As Object
    Dim i As Long, j As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each WS In ThisWorkbook.Worksheets
        lastrow = WS.Cells(WS.Rows.Count, 1).End(xlUp).Row
        arr = WS.Range("A1:A" & lastrow).Value
        
        'Clear dictionary for each worksheet
        dict.RemoveAll
        
        'Add unique values to dictionary
        For i = 1 To UBound(arr)
            If Not dict.Exists(arr(i, 1)) Then
                dict.Add arr(i, 1), ""
            End If
        Next i
        
        'Paste unique values to column H of the same worksheet
        WS.Range("H1").Resize(dict.Count) = Application.Transpose(dict.keys)
    Next WS
    
End Sub
