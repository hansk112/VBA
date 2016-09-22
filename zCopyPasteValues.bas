Attribute VB_Name = "zCopyPasteValues"

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''' Copy pastes all values in the worksheet ''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub copyPasteValuesWS()

Dim ws As Worksheet
 For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ws.Cells.Copy
        ws.Range("A1").PasteSpecial Paste:=xlValues
        ws.Range("K9").Select
 Next ws
Application.CutCopyMode = False

End Sub

