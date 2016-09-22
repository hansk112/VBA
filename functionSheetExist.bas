Attribute VB_Name = "functionSheetExist"
Function e(n As String) As Boolean
e = False
For Each ws In ThisWorkbook.Worksheets
    If n = ws.Name Then
        e = True
    End If
Next ws

End Function
