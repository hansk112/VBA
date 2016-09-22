Attribute VB_Name = "Module1"

Sub main()
Attribute main.VB_ProcData.VB_Invoke_Func = " \n14"

'
Dim ws As Worksheet
Dim nrows As Long
Dim code As String, coyname As String, secname As String, sRec As String


Application.DisplayAlerts = False

Set ws = Worksheets("list")

ws.Activate

nrows = Cells(Rows.Count, 1).End(xlUp).Row
 
For i = 1 To nrows
   

   code = ws.Range("A" & i)
   coyname = ws.Range("A" & i).Offset(0, 1)
   secname = ws.Range("A" & i).Offset(0, 2)



'   If checkExists(code) = True Then
'   Worksheets(code).Delete
'   End If
   
   
'   ThisWorkbook.Sheets.Add.Name = code
   
'   incomeStatement (code)
'   balancesheet (code)
'   cashFlow (code)
'   finRatios (code)
   recoStock (code)
   Debug.Print code
'   keyStock (code)
'   formatSheet code, coyname, secname

 '   recommendSum

Next i

End Sub

Function checkExists(shName As String)
Application.DisplayAlerts = False
checkExists = False

For Each ws In ThisWorkbook.Worksheets
    If shName = ws.Name Then
        checkExists = True
    End If
Next ws

Application.DisplayAlerts = False

End Function

Function ConvertToLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function

Sub financialStatmentAnalysis()
Dim fsanaysis As String

fsanaysis = "FS Anaysis"

If checkExists(fsanaysis) = True Then
Worksheets(fsanaysis).Delete
End If
Sheets.Add.Name = fsanaysis




End Sub
