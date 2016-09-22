Attribute VB_Name = "validateBaking"
Sub validatTotionCheckBake()

Dim ws As Worksheet
Dim i As Long

Dim arrAgTy As Variant
Dim arrPyFq As Variant
Dim arrWsNm As Variant
Dim j As Integer
Dim k As Integer
Dim curPeriod As Integer

Dim icol As Integer
Dim periodRow As String

Dim rngFndRow As Range
Dim rngFndCSF As Range

Set arrAgTy = Range("agmtType")
Set arrPyFq = Range("payFreq")
Set arrWsNm = Range("wsName")
Set ws = Worksheets("mapCustomer")
Set wsData = Worksheets("data")
Set lookRange = wsData.Range("rowPeriod")

 j = 1
 k = 1


Application.ScreenUpdating = False

ws.Activate

'count number of riws
nRows = ws.Cells(Rows.Count, 1).End(xlUp).Row
curPeriod = ws.Range("curPeriod")


Debug.Print curPeriod

' set period to currentPeriod
wsData.Activate
periodRow = Application.WorksheetFunction.VLookup(curPeriod, lookRange, 11, 0)
Debug.Print periodRow

ws.Activate
For i = 3 To nRows


ws.Activate
If arrAgTy(j, 1) = "bake" And arrPyFq(k, 1) = "Qtr" Then
Dim name2 As String

name2 = Range("A" & i).Value
Worksheets(name2).Activate

Dim actShName As String

actShName = ActiveSheet.Name

'handle period
If Range("C11").Value <> "Period" Or Range("C11").Value = "" Then
    If Range("C11").Value = "" Then
        Dim period As String
        period = "Blank"
        GoTo messagePeriod:
    End If
    If Range("C11").Value <> "Period" Then
        period = Range("C11").Value
    End If
messagePeriod:
res = MsgBox("Worksheet " & actShName & " Cell C11 value is: " & period & vbCr & _
"Worksheet " & actShName & " Cell C11 value needs to be: Period" & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that Period is on column C and row 11" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct ", vbOKCancel, "File Violation")
If res = vbCancel Then
    Worksheets(actShName).Activate
    Worksheets(actShName).Range("C11").Select
    Exit Sub
End If
End If

'handle total branded loaf & Occ Bake
If Range("C39").Value <> "Total Branded Loaf & Occ Bake" Or _
Range("C39").Value = "" Then
    
Dim totBrndLoaf As String
totBrndLoaf = Range("C39").Value
    If totBrndLoaf = "" Then
        totBrndLoaf = "Blank"
    End If
res = MsgBox("Worksheet " & actShName & " Cell C39 is: " & totBrndLoaf & vbCr & _
"Worksheet " & actShName & " Cell C39 needs to be: Total Branded Loaf & Occ Bake" & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that Period is on column C and row 39" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct", vbOKCancel, "File Violation Error")
If res = vbCancel Then
    Worksheets(actShName).Activate
    Worksheets(actShName).Range("C39").Select
    Exit Sub
End If
End If


'handle business partnership
If Range("C44").Value <> "Business Partnership" Or _
Range("C44").Value = "" Then

Dim busPart As String
busPart = Range("C44").Value
    If busPart = "" Then
        busPart = "Blank"
    End If

res = MsgBox("Worksheet " & actShName & " Cell C44 is: " & busPart & vbCr & _
"Worksheet " & actShName & " Cell C44 needs to be: Business Partnership" & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that Business Partnership is on column C and row 44" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct", vbOKCancel, "File Violation Error")
If res = vbCancel Then
    Worksheets(actShName).Activate
    Worksheets(actShName).Range("C44").Select
    Exit Sub
End If
End If


'rebate branded loaf & occ bake
If Range("C49").Value <> "Rebate Branded Loaf & Occ Bake" Or Range("C49").Value = "" Then

busPart = Range("C49").Value
    If busPart = "" Then
        busPart = "Blank"
    End If

res = MsgBox("Worksheet " & actShName & " Cell C49 is: " & busPart & vbCr & _
"Worksheet " & actShName & " Cell C49 needs to be: Business Partnership" & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that 'Rebate Branded Loaf & Occ Bake' is on column C and row 49" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct", vbOKCancel, "File Violation Error")
If res = vbCancel Then
    Worksheets(actShName).Activate
    Worksheets(actShName).Range("C49").Select
    Exit Sub
End If
End If

'category support fund
If Range("C56").Value <> "Category Support Fund" Or Range("C56").Value = "" Then

busPart = Range("C56").Value
    If busPart = "" Then
        busPart = "Blank"
    End If

res = MsgBox("Worksheet " & actShName & " Cell C49 is: " & busPart & vbCr & _
"Worksheet " & actShName & " Cell C49 needs to be: Category Support Fund" & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that 'Category Support Fund' is on column C and row 56" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct", vbOKCancel, "File Violation Error")
If res = vbCancel Then
    Worksheets(actShName).Activate
    Worksheets(actShName).Range("C56").Select
    Exit Sub
End If
End If

'closing balance
If Range("C61").Value <> "Closing Balance" Or Range("C61").Value = "" Then

busPart = Range("C61").Value
    If busPart = "" Then
        busPart = "Blank"
    End If

res = MsgBox("Worksheet " & actShName & " Cell C61 is: " & busPart & vbCr & _
"Worksheet " & actShName & " Cell C61 needs to be: Closing Balance" & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that 'Closing Balance' is on column C and row 61" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct", vbOKCancel, "File Violation Error")
If res = vbCancel Then
    Worksheets(actShName).Activate
    Worksheets(actShName).Range("C61").Select
    Exit Sub
End If
End If

'baking - category support fund
If Range("B63").Value <> "Baking - Category Support Fund" Or Range("B63").Value = "" Then

busPart = Range("B63").Value
    If busPart = "" Then
        busPart = "Blank"
    End If

res = MsgBox("Worksheet " & actShName & " Cell B63 is: " & busPart & vbCr & _
"Worksheet " & actShName & " Cell B63 needs to be: Baking - Category Support Fund" & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that 'Baking - Category Support Fund' is on column B and row 63" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct", vbOKCancel, "File Violation Error")
If res = vbCancel Then
    Worksheets(actShName).Activate
    Worksheets(actShName).Range("B63").Select
    Exit Sub
End If
End If


' handle/check that Baking - Catgory Support Fund exists

Set rngFndCSF = ActiveSheet.Cells.Find(What:="Baking - Category Support Fund", After:=ActiveCell, _
    LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
If rngFndCSF Is Nothing Then

res = MsgBox("Worksheet " & actShName & " Cell B63 needs to be: " & vbCr & _
"'Baking - Category Support Fund'" & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that 'Baking - Category Support Fund' is on column B and row 63" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct", vbOKCancel, "File Violation Error")
    If res = vbCancel Then
        Worksheets(actShName).Activate
        Worksheets(actShName).Range("B63").Select
        Exit Sub
    End If

End If
    


' handle/check that the month exists for CSF import
Set rngFndRow = ActiveSheet.Cells.Find(What:=periodRow, _
    LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
'  On Error GoTo dateErr:
 
If rngFndRow Is Nothing Then

res = MsgBox("Worksheet " & actShName & " needs date value " & periodRow & ": " & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that " & periodRow & " is on " & actShName & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct", vbOKCancel, "File Violation Error")
    If res = vbCancel Then
        Worksheets(actShName).Activate
        Worksheets(actShName).Range("C12").Select
        Exit Sub
    End If


End If
    



'*********
' last IF
'*********
End If

' increment array by 1
j = 1 + j
k = 1 + k


'''' end if statement


Next i





ws.Activate
complete = MsgBox("File validation for Baking" & vbCr & _
vbCr & _
"Completed Successfully!", vbOKOnly, "Validation Complete")


End Sub

Sub u()

Dim arrAgTy As Variant
Dim arrPyFq As Variant
Dim i As Integer
Dim icol As Integer

arrAgTy = Range("agmtType")
arrPyFq = Range("payFreq")

For i = 1 To 3
    For icol = 1 To 3
    MsgBox arrAgTy(i, icol)
    i = i + 1
    Next icol
Next i
End Sub



Sub ut()

'Dim monthRange As String

'mothRange = Application.VLookup(3, "testDate", 5)

Range (testDate)

'moth = Calculate(3, "testDate")
Debug.Print moth

Dim arrAgTy As Variant
Dim arrPyFq As Variant
Dim i As Integer
Dim icol As Integer

arrAgTy = Range("agmtType")
arrPyFq = Range("payFreq")

For i = 1 To 12
  
'    MsgBox arrAgTy(i, 1)
'    MsgBox arrPyFq(i, 1)
'    If arrAgTy(i, 1) = "bake" And arrPyFq(i, 1) = "Qtr" Then
'        MsgBox ("yes")
'    End If

Next i
End Sub







