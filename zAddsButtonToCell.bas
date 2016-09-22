Attribute VB_Name = "zAddsButtonToCell"
Sub calcWS()


Dim ws As Worksheet
Dim wsName As String
Dim i As Integer
Dim nRows As Long
Dim rng As Range

Set ws = Worksheets("mapCustomer")

nRows = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Worksheets("PSROTORUA").Range("F29").Value

'For i = 1 To nRows
'If wsName = ws.Range("A" & i).Value = 1 Then
'MsgBox ("u")
'End If
'MsgBox (wsName)
'Next i
ws.Range("AA3").Formula = "=D4*$K4"

Dim psAlbany As String
psAlbany = ws.Range("A4").Value
'Worksheets(psAlbany).Range("b7").Value

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''' NO LONGER REQUIRED''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''

Sub createRangeOfButtons()
  Dim btn As Button
  Application.ScreenUpdating = False
  
  Dim t As Range
  For i = 1 To 37
    Set t = ActiveSheet.Range(Cells(i, 37), Cells(i, 37))
    Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
      .OnAction = "btn"
      .Caption = "Btn " & i
      .Name = "Btn" & i
    End With
  Next i
  Application.ScreenUpdating = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''' ADDS BUTTON TO SELECTED CELL '''''
'''''''''''''''''''''''''''''''''''''''''''''''''

Sub buttonInACell()
Dim btn As Button


   Set t = ActiveSheet.Cells(42, 39)
   Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)

With btn
    .Caption = "NWKaikohe"
    .Name = "NWKaikohe"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlTop
End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''' ADDS BUTTON TO SELECTED CELL '''''
'''''''''''''''''''''''''''''''''''''''''''''''''

Sub btnAddCell()

Dim btn1 As Button
Dim tRng As Range
Dim nWS As String
Dim num As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''CHANGE NUM TO THE ROW NUMBER THAT YOU WANT TO INSERT BUTTON '''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

num = 39
nWS = ActiveSheet.Range("A" & num).Value


Set tRng = ActiveSheet.Cells(num, 39)
Set btn1 = ActiveSheet.Buttons.Add(tRng.Left, tRng.Top, tRng.Width, tRng.Height)

With btn1
    .Caption = nWS
    .Name = nWS
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlTop
End With

End Sub

