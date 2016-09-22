Attribute VB_Name = "validateOneGF"

Sub validateOneGF()

Dim stName As String
Dim gpName As String
Dim rebTotal As String
Dim othReb As String
Dim grdTot As String
Dim addPay As String
Dim oneGF As String
Dim clBk As String
Dim clCh As String
Dim periodRow As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim a As String
Dim arrAgTy As Variant
Dim arrPyFq As Variant
Dim arrWsNm As Variant
Dim arrAct As Variant
Dim rngFndCSF As Range


Dim ws As Worksheet

Set arrAgTy = Range("agmtType")
Set arrPyFq = Range("payFreq")
Set arrWsNm = Range("wsName")
Set arrAct = Range("active")
Set ws = Worksheets("mapCustomer")
Set wsData = Worksheets("data")
Set lookRange = wsData.Range("rowPeriod")


gpName = "Group"
rebTotal = "Rebate Total"
othReb = "Other Rebate"
grdTot = "Grand Total "
bPrtPymt = "Business Partnership Payment"
qPymt = "Quarterly Payment incl GST"
addPay = "Additional Payments"
oneGF = "1GF Balance"
clBk = "Closing Balance"
clCh = "Closing Balance"
'a = "A"

ws.Activate

Application.ScreenUpdating = False

curPeriod = ws.Range("curPeriod")

periodRow = Application.WorksheetFunction.VLookup(curPeriod, lookRange, 11, 0)
nrowsOneGF = ws.Cells(Rows.Count, 1).End(xlUp).Row
j = 1
k = 1
l = 1

For i = 3 To nrowsOneGF

ws.Activate
If arrAgTy(j, 1) = "oneGF" And arrPyFq(k, 1) = "Qtr" And arrAct(l, 1) = "Y" Then
Dim name2 As String

name2 = Range("A" & i).Value
Worksheets(name2).Activate
Dim shName As String
shName = ActiveSheet.Name

'handle Group
If Range("A11").Value <> gpName Or Range("A11").Value = "" Then
    If Range("A11").Value = "" Then
        Dim period As String
        period = "Blank"
        GoTo messagePeriod:
    End If
    If Range("A11").Value <> gpName Then
        period = Range("A11").Value
    End If
messagePeriod:
res = MsgBox("Worksheet " & shName & " Cell A11 value is: " & period & vbCr & _
"Worksheet " & shName & " Cell A11 value needs to be: " & gpName & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that " & gpName & " is on column A and row 11" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct ", vbOKCancel, "File Violation")
If res = vbCancel Then
    Worksheets(shName).Activate
    Worksheets(shName).Range("A11").Select
    Exit Sub
End If
End If

'handle rebate total
If Range("A83").Value <> rebTotal Or Range("A83").Value = "" Then
    If Range("A83").Value = "" Then
        Dim bRebTotal As String
        bRebTotal = "Blank"
        GoTo messagePeriod1:
    End If
    If Range("A83").Value <> rebTotal Then
        period = Range("A83").Value
    End If
messagePeriod1:
res = MsgBox("Worksheet " & shName & " Cell A83 value is: " & rebTotal & vbCr & _
"Worksheet " & shName & " Cell A83 value needs to be: " & rebTotal & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that " & rebTotal & " is on column A and row 83" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct ", vbOKCancel, "File Violation")
If res = vbCancel Then
    Worksheets(shName).Activate
    Worksheets(shName).Range("A83").Select
    Exit Sub
End If
End If


'handle other rebate
If Range("A85").Value <> othReb Or Range("A85").Value = "" Then
    If Range("A85").Value = "" Then
        Dim bothReb As String
        bothReb = "Blank"
        GoTo messagePeriod2:
    End If
    If Range("A85").Value <> othReb Then
        period = Range("A85").Value
    End If
messagePeriod2:
res = MsgBox("Worksheet " & shName & " Cell A85 value is: " & bothReb & vbCr & _
"Worksheet " & shName & " Cell A85 value needs to be: " & othReb & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that " & othReb & " is on column A and row 85" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct ", vbOKCancel, "File Violation")
If res = vbCancel Then
    Worksheets(shName).Activate
    Worksheets(shName).Range("A85").Select
    Exit Sub
End If
End If


'handle Grand Total
If Range("A129").Value <> grdTot Or Range("A129").Value = "" Then
    If Range("A129").Value = "" Then
        Dim bgrdTot As String
        bgrdTot = "Blank"
        GoTo messagePeriod3:
    End If
    If Range("A129").Value <> bgrdTot Then
        period = Range("A129").Value
    End If
messagePeriod3:
res = MsgBox("Worksheet " & shName & " Cell A129 value is: " & bgrdTot & vbCr & _
"Worksheet " & shName & " Cell A129 value needs to be: " & grdTot & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that " & grdTot & " is on column A and row 129" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct ", vbOKCancel, "File Violation")
If res = vbCancel Then
    Worksheets(shName).Activate
    Worksheets(shName).Range("A129").Select
    Exit Sub
End If
End If


'handle Business Partnership
If Range("A131").Value <> bPrtPymt Or Range("A131").Value = "" Then
    If Range("A131").Value = "" Then
        Dim bbPrtPymt As String
        bbPrtPymt = "Blank"
        GoTo messagePeriod4:
    End If
    If Range("A131").Value <> bPrtPymt Then
        period = Range("A131").Value
    End If
messagePeriod4:
res = MsgBox("Worksheet " & shName & " Cell A131 value is: " & bbPrtPymt & vbCr & _
"Worksheet " & shName & " Cell A131 value needs to be: " & bPrtPymt & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that " & grdTot & " is on column A and row 131" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct ", vbOKCancel, "File Violation")
If res = vbCancel Then
    Worksheets(shName).Activate
    Worksheets(shName).Range("A131").Select
    Exit Sub
End If
End If


'handle Quarter Payment incl GST
If Range("A142").Value <> qPymt Or Range("A142").Value = "" Then
    If Range("A142").Value = "" Then
        Dim bqPymt As String
        bqPymt = "Blank"
        GoTo messagePeriod5:
    End If
    If Range("A142").Value <> qPymt Then
        qPymt = Range("A142").Value
    End If
messagePeriod5:
res = MsgBox("Worksheet " & shName & " Cell A142 value is: " & bqPymt & vbCr & _
"Worksheet " & shName & " Cell A142 value needs to be: " & qPymt & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that " & qPymt & " is on column A and row 142" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct ", vbOKCancel, "File Violation")
If res = vbCancel Then
    Worksheets(shName).Activate
    Worksheets(shName).Range("A142").Select
    Exit Sub
End If
End If



'handle Additonal Payments
If Range("A144").Value <> addPay Or Range("A144").Value = "" Then
    If Range("A144").Value = "" Then
        Dim baddPay As String
        baddPay = "Blank"
        GoTo messagePeriod6:
    End If
    If Range("A144").Value <> qPymt Then
        baddPay = Range("A144").Value
    End If
messagePeriod6:
res = MsgBox("Worksheet " & shName & " Cell A144 value is: " & baddPay & vbCr & _
"Worksheet " & shName & " Cell A144 value needs to be: " & addPay & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that " & addPay & " is on column A and row 144" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct ", vbOKCancel, "File Violation")
If res = vbCancel Then
    Worksheets(shName).Activate
    Worksheets(shName).Range("A144").Select
    Exit Sub
End If
End If

'handle 1GF Balance
If Range("A149").Value <> oneGF Or Range("A149").Value = "" Then
    If Range("A149").Value = "" Then
        Dim boneGF As String
        boneGF = "Blank"
        GoTo messagePeriod7:
    End If
    If Range("A149").Value <> boneGF Then
        boneGF = Range("A149").Value
    End If
messagePeriod7:
res = MsgBox("Worksheet " & shName & " Cell A149 value is: " & boneGF & vbCr & _
"Worksheet " & shName & " Cell A149 value needs to be: " & oneGF & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that " & oneGF & " is on column A and row 149" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct ", vbOKCancel, "File Violation")
If res = vbCancel Then
    Worksheets(shName).Activate
    Worksheets(shName).Range("A149").Select
    Exit Sub
End If
End If


'handle baking CSF closing balance
If Range("A155").Value <> clBk Or Range("A155").Value = "" Then
    If Range("A155").Value = "" Then
        Dim bclBk As String
        bclBk = "Blank"
        GoTo messagePeriod8:
    End If
    If Range("A155").Value <> bclBk Then
        bclBk = Range("A155").Value
    End If
messagePeriod8:
res = MsgBox("Worksheet " & shName & " Cell A155 value is: " & bclBk & vbCr & _
"Worksheet " & clBk & " Cell A155 value needs to be: " & clBk & vbCr & _
"" & RebavbCr & _
"******************************************************" & vbCr & _
"Please check that " & clBk & " is on column A and row 155" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct ", vbOKCancel, "File Violation")
If res = vbCancel Then
    Worksheets(shName).Activate
    Worksheets(shName).Range("A161").Select
    Exit Sub
End If
End If


'handle Chilled CSF closing balance
If Range("A161").Value <> clCh Or Range("A161").Value = "" Then
    If Range("A155").Value = "" Then
        Dim bclCh As String
        bclCh = "Blank"
        GoTo messagePeriod9:
    End If
    If Range("A155").Value <> bclBk Then
        bclCh = Range("A155").Value
    End If
messagePeriod9:
res = MsgBox("Worksheet " & shName & " Cell A161 value is: " & bclCh & vbCr & _
"Worksheet " & clCh & " Cell A161 value needs to be: " & clCh & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that " & clCh & " is on column A and row 161" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct ", vbOKCancel, "File Violation")
If res = vbCancel Then
    Worksheets(shName).Activate
    Worksheets(shName).Range("A161").Select
    Exit Sub
End If
End If


' handle/check that Baking - Catgory Support Fund exists

Set rngFndCSF = ActiveSheet.Cells.Find(What:="Baking - Category Support Fund", After:=ActiveCell, _
    LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
If rngFndCSF Is Nothing Then

res = MsgBox("Worksheet " & shName & " Cell B151 needs to be: " & vbCr & _
"'Baking - Category Support Fund'" & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that 'Baking - Category Support Fund' is on column B and row 151" & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct", vbOKCancel, "File Violation Error")
    If res = vbCancel Then
        Worksheets(shName).Activate
        Worksheets(shName).Range("B151").Select
        Exit Sub
    End If

End If


Set rngFndRow = ActiveSheet.Cells.Find(What:=periodRow, _
    LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
'  On Error GoTo dateErr:
 
If rngFndRow Is Nothing Then

res = MsgBox("Worksheet " & shName & " needs date value " & periodRow & ": " & vbCr & _
"" & vbCr & _
"******************************************************" & vbCr & _
"Please check that " & periodRow & " is on " & shName & vbCr & _
"******************************************************" & vbCr & _
"" & vbCr & _
"Select OK to continue or Cancel to correct", vbOKCancel, "File Violation Error")
    If res = vbCancel Then
        Worksheets(shName).Activate
        Worksheets(shName).Range("C12").Select
        Exit Sub
    End If


End If





'quarterly payment incl GST

'Dim t As Boolean
'
'res = MsgBox("Worksheet " & shName & t = Range(A & "11").Value <> "Group" & " Cell A11 value is: " & gpName & vbCr & _
'"Worksheet " & shName & " Cell C11 value needs to be: Period" & vbCr & _
'"" & vbCr & _
'"******************************************************" & vbCr & _
'"Please check that Period is on column C and row 11" & vbCr & _
'"******************************************************" & vbCr & _
'"" & vbCr & _
'"Select OK to continue or Cancel to correct ", vbOKCancel, "File Violation")
'If res = vbCancel Then
'    Worksheets(shName).Activate
'    Worksheets(shName).Range("A11").Select
'    Exit Sub
'End If
'End If


Debug.Print arrAgTy(j, 1) & " " & arrPyFq(k, 1) & " " & i & " " & shName

End If

' increment array by 1
j = 1 + j
k = 1 + k
l = l + 1

Next i

complete = MsgBox("File validation for One GF" & vbCr & _
vbCr & _
"Completed Successfully!", vbOKOnly, "Validation Complete")




End Sub

Sub st()


Dim t As Boolean

t = Range("a11").Value <> "Group"

MsgBox (IIf(t = True, "thiis true", "nothin"))

End Sub




