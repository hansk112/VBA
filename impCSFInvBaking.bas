Attribute VB_Name = "impCSFInvBaking"

Function ConvertToLetter(icol As Long) As String
Dim iAlpha As Integer
Dim iRemainder As Integer
iAlpha = Int(icol / 27)
iRemainder = icol - (iAlpha * 26)
If iAlpha > 0 Then
ConvertToLetter = Chr(iAlpha + 64)
End If
If iRemainder > 0 Then
ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
End If
End Function
Sub importDFG()

Dim nRows As Long
Dim ws As Worksheet
Dim mthSearch As String
Dim colMthCSF As Long
Dim arrAgmtType As Variant
Dim tempAgmtType As String
Dim j As Integer
Dim col1 As Long
Dim col2 As Long
Dim copyCol As String
Dim pasteOneGFCol As String
Dim copyOneGFCol As String
Dim startTime As Double
Dim minutesElap As String, bakeCSF As String

startTime = Timer

Application.ScreenUpdating = False
' set ws

On Error GoTo messageError:
Set ws = Worksheets("csfInvoicesBaking")
' set array value to 1
ws.Activate
j = 1

With ws
       .Select
      nRows = .Cells(Rows.Count, 4).End(xlUp).Row
       .Range("J1").Value = "custID"
       .Range("J2").Formula = "=value(B2)"
       .Range("J2:J" & nRows).FillDown
       .Calculate
       .Range("J2:J" & nRows).Copy
       .Range("J2:J" & nRows).PasteSpecial xlValues
       .Range("K1").Value = "invNum"
       .Range("K2").Formula = "=value(c2)"
       .Range("K2:K" & nRows).FillDown
       .Calculate
       .Range("K2:K" & nRows).Copy
       .Range("K2:K" & nRows).PasteSpecial xlValues
       .Range("L1").Value = "Product"
       .Range("G2:G" & nRows).Copy
       .Range("L2:L" & nRows).PasteSpecial xlValues
       .Range("M1").Value = "Date"
       .Range("M2:M" & nRows).Formula = "=TEXT(D2, ""d.mm.yy"")"
       .Range("M2").Select
       Selection.AutoFill Destination:=Range("M2:M" & nRows)
       .Range("N1").Value = "exclGST"
       .Range("N2:N" & nRows).Formula = "=F2-(F2/23*3)"
       .Range("O1").Value = "Month"
       .Range("O2:O" & nRows).Formula = "=VLOOKUP(I2,csfDateLookup,10,FALSE)"
       .Range("O2").Select
       Selection.AutoFill Destination:=Range("O2:O" & nRows)
       .Range("O2:O" & nRows).Copy
       .Range("O2:O" & nRows).PasteSpecial xlValues
       .Range("P1").Value = "wsName"
       .Range("P2:P" & nRows).Formula = "=INDEX(wsName,MATCH(VALUE(J2),sapID,0))"
       .Range("P2").Select
       .Calculate
       Selection.AutoFill Destination:=Range("P2:P" & nRows)
       .Range("Q1").Value = "agmtType"
       .Range("Q2:Q" & nRows).Formula = "=INDEX(agmtType,MATCH(VALUE(csfInvoicesBaking!J2),sapID,0))"
       .Calculate
       'store agreemt in array
    arrAgmtType = Range("Q2:Q" & nRows)
End With

' loop sheet

For i = 2 To nRows

Dim wsCSFName As String
wsCSFName = Worksheets("csfInvoicesBaking").Range("P" & i)
Worksheets(wsCSFName).Activate

sheetName = ActiveSheet.Name

With ActiveSheet
    On Error GoTo bakeCSFError:
    
    bakeCSF = "Baking - Category Support Fund"
    
    If Cells.Find(bakeCSF) Is Nothing Then
    MsgBox "Please check " & sheetName & " for value " & bakeCSF & vbNewLine & _
    "The proceeding CSF invoices will need to be checked " & vbNewLine & _
    "Press OK to continue", vbExclamation = vbOKCancel, bakeCSF & " Not Found"
GoTo unfoundCSF:
    End If
    
     .Cells.Find(What:="Baking - Category Support Fund", After:=ActiveCell, _
    LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
    
    End With
    
    col1 = ActiveCell.Column
    row158 = ActiveCell.Row
    
    Selection.End(xlDown).Select
    
    Dim invNum As String
    
   invNum = ActiveCell.Value
    
    ' hanndle format when no invoices listed
    If invNum = "Invoice No" Or invNum = "Invoice #" Then
        ActiveCell.Offset(1, 0).Select
        rowPast = ActiveCell.Row
        
    ElseIf invNum <> "Invoice No" Or invNum <> "Invoice #" Then
        ActiveCell.Offset(1, 0).Select
        ActiveCell.EntireRow.Insert
        rowPast = ActiveCell.Row
    
    End If
    
   
    ws.Range("K" & i & ":M" & i).Copy
    ActiveCell.PasteSpecial xlPasteValues
    
    ' agmt type is onegf add two cells to the right and copy paste values
        If arrAgmtType(j, 1) = "oneGF" Then
        col1 = col1 + 2
        col2 = col1 + 2
        copyOneGFCol = ConvertToLetter(col1)
        pasteOneGFCol = ConvertToLetter(col2)
        
        ActiveSheet.Range(copyOneGFCol & rowPast).Copy
        ActiveSheet.Range(pasteOneGFCol & rowPast).PasteSpecial xlPasteValues
        End If
    
    ' search for month
    mthSearch = ws.Range("O" & i)
    Selection.Copy
    
    If Cells.Find(mthSearch) Is Nothing Then
    MsgBox "Please check " & sheetName & " for value " & mthSearch & vbNewLine & _
    "The proceeding CSF will need to be removed " & vbNewLine & _
    "Press OK to continue ", vbExclamation = vbOKCancel, mthSearch & " Not Found"
GoTo unfoundCSF:
    End If
    
    Cells.Find(What:=mthSearch, After:=ActiveCell, LookIn:= _
    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False).Activate
    colMthCSF = ActiveCell.Column
    colmthletcsf = ConvertToLetter(colMthCSF)
    
    ws.Range("N" & i & ":N" & i).Copy
    'paste invoice value
    ActiveSheet.Range(colmthletcsf & rowPast).PasteSpecial xlPasteValues

    Debug.Print arrAgmtType(j, 1)
unfoundCSF:
    j = j + 1

Next i

Application.CutCopyMode = False

minutesElap = Format((Timer - startTime) / 86400, "s")
ws.Activate
res = MsgBox("CSF Succesfully Updated in " & vbCr & _
 minutesElap & " seconds", , "CSF Successful Update!")


messageError:
If Err.Number = 9 Then
    MsgBox ("Please create a Worksheet called 'csfInvoicesBaking'" & vbCr _
     & "and dump the csf Invoices in that Worksheet")
    Exit Sub

End If

bakeCSFError:

End Sub










